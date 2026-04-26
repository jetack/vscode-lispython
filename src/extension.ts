import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as net from 'net';
import { execFile, ChildProcess, spawn } from 'child_process';
import {
    CloseAction,
    ErrorAction,
    LanguageClient,
    LanguageClientOptions,
    ServerOptions,
} from 'vscode-languageclient/node';

let client: LanguageClient | undefined;
let statusBarItem: vscode.StatusBarItem;

// ---------------------------------------------------------------------------
// Python discovery
// ---------------------------------------------------------------------------

interface PythonInfo {
    path: string;
    version: string;
    hasLispy: boolean;
    label: string;
}

function findPythonCandidates(): Promise<PythonInfo[]> {
    const candidates = new Set<string>();

    // 1. Current setting
    const config = vscode.workspace.getConfiguration('lispython');
    const configured = config.get<string>('lsp.pythonPath', '');
    if (configured) {
        candidates.add(configured);
    }

    // 2. Common system pythons
    for (const name of ['python3', 'python']) {
        candidates.add(name);
    }

    // 3. Scan workspace folders for .venv/venv
    for (const folder of vscode.workspace.workspaceFolders ?? []) {
        const root = folder.uri.fsPath;
        for (const venvDir of ['.venv', 'venv', '.env', 'env']) {
            const venvPython = path.join(root, venvDir, 'bin', 'python');
            if (fs.existsSync(venvPython)) {
                candidates.add(venvPython);
            }
        }
    }

    // 4. Check parent directories for .venv (e.g., lispython repo)
    for (const folder of vscode.workspace.workspaceFolders ?? []) {
        let dir = folder.uri.fsPath;
        for (let i = 0; i < 3; i++) {
            dir = path.dirname(dir);
            if (dir === '/') break;
            for (const venvDir of ['.venv', 'venv']) {
                const venvPython = path.join(dir, venvDir, 'bin', 'python');
                if (fs.existsSync(venvPython)) {
                    candidates.add(venvPython);
                }
            }
        }
    }

    // Probe each candidate
    return Promise.all(
        [...candidates].map((pythonPath) => probePython(pythonPath))
    ).then((results) => results.filter((r): r is PythonInfo => r !== null));
}

function probePython(pythonPath: string): Promise<PythonInfo | null> {
    return new Promise((resolve) => {
        const script = `
import sys
try:
    import lispy.lsp
    has_lispy = True
except ImportError:
    has_lispy = False
v = sys.version.split()[0]
print(f"{v}|{has_lispy}")
`;
        execFile(
            pythonPath,
            ['-c', script],
            { timeout: 5000 },
            (err, stdout) => {
                if (err) {
                    resolve(null);
                    return;
                }
                const [version, hasLispy] = stdout.trim().split('|');
                const ready = hasLispy === 'True';
                const tag = ready ? '' : ' (lispython not installed)';
                resolve({
                    path: pythonPath,
                    version,
                    hasLispy: ready,
                    label: `Python ${version} — ${pythonPath}${tag}`,
                });
            }
        );
    });
}

// ---------------------------------------------------------------------------
// Status bar
// ---------------------------------------------------------------------------

function createStatusBar(context: vscode.ExtensionContext): void {
    statusBarItem = vscode.window.createStatusBarItem(
        vscode.StatusBarAlignment.Right,
        100
    );
    statusBarItem.command = 'lispython.selectPythonPath';
    context.subscriptions.push(statusBarItem);
    updateStatusBar();
}

function updateStatusBar(): void {
    const config = vscode.workspace.getConfiguration('lispython');
    const pythonPath = config.get<string>('lsp.pythonPath', 'python3');
    const short = pythonPath.includes('/')
        ? path.basename(path.dirname(path.dirname(pythonPath))) + '/' + path.basename(pythonPath)
        : pythonPath;
    statusBarItem.text = `$(symbol-misc) LisPy: ${short}`;
    statusBarItem.tooltip = `LisPython LSP Python: ${pythonPath}\nClick to change`;
    statusBarItem.show();
}

// ---------------------------------------------------------------------------
// Select Python command
// ---------------------------------------------------------------------------

async function selectPythonPath(): Promise<void> {
    const statusMsg = vscode.window.setStatusBarMessage('$(loading~spin) Searching for Python interpreters...');

    let candidates: PythonInfo[];
    try {
        candidates = await findPythonCandidates();
    } finally {
        statusMsg.dispose();
    }

    // Sort: lispy-ready first, then by path
    candidates.sort((a, b) => {
        if (a.hasLispy !== b.hasLispy) return a.hasLispy ? -1 : 1;
        return a.path.localeCompare(b.path);
    });

    const items: (vscode.QuickPickItem & { pythonPath?: string })[] = candidates.map((c) => ({
        label: c.hasLispy ? `$(check) ${c.label}` : `$(warning) ${c.label}`,
        description: c.hasLispy ? 'ready' : 'needs lispython',
        pythonPath: c.path,
    }));

    items.push(
        { label: '', kind: vscode.QuickPickItemKind.Separator },
        { label: '$(file-directory) Browse for Python executable...' },
    );

    const selected = await vscode.window.showQuickPick(items, {
        placeHolder: 'Select Python interpreter for LisPython LSP',
        matchOnDescription: true,
    });

    if (!selected) return;

    let chosenPath: string | undefined;

    if (selected.pythonPath) {
        chosenPath = selected.pythonPath;
    } else {
        // Browse dialog
        const uris = await vscode.window.showOpenDialog({
            canSelectFiles: true,
            canSelectFolders: false,
            canSelectMany: false,
            title: 'Select Python Interpreter',
            openLabel: 'Select',
        });
        if (uris && uris.length > 0) {
            chosenPath = uris[0].fsPath;
        }
    }

    if (!chosenPath) return;

    const config = vscode.workspace.getConfiguration('lispython');
    await config.update('lsp.pythonPath', chosenPath, vscode.ConfigurationTarget.Workspace);

    updateStatusBar();
    await restartLspServer();
}

// ---------------------------------------------------------------------------
// LSP client lifecycle
// ---------------------------------------------------------------------------

async function startLspServer(): Promise<void> {
    const config = vscode.workspace.getConfiguration('lispython');
    const serverCommand = config.get<string>('lsp.pythonPath', 'python3');
    const serverArgs = ['-m', 'lispy.lsp'];

    // Probe first — fail fast with a clear message instead of crash-looping
    const probe = await probePython(serverCommand);
    if (!probe) {
        throw new Error(
            `Cannot run '${serverCommand}'. Click the status bar or run "LisPython: Select Python Interpreter".`
        );
    }
    if (!probe.hasLispy) {
        throw new Error(
            `Python at '${serverCommand}' does not have lispython installed. ` +
            `Install it with 'pip install lispython' or select a different interpreter.`
        );
    }

    const serverOptions: ServerOptions = {
        command: serverCommand,
        args: serverArgs,
    };

    const clientOptions: LanguageClientOptions = {
        documentSelector: [{ scheme: 'file', language: 'lispython' }],
        synchronize: {
            fileEvents: vscode.workspace.createFileSystemWatcher('**/*.lpy'),
        },
        errorHandler: {
            // Close the client after the first unexpected exit — don't crash-loop
            error: () => ({ action: ErrorAction.Shutdown }),
            closed: () => ({ action: CloseAction.DoNotRestart }),
        },
    };

    client = new LanguageClient(
        'lispython',
        'LisPython Language Server',
        serverOptions,
        clientOptions,
    );

    await client.start();
}

async function stopLspServer(): Promise<void> {
    if (client) {
        await client.stop();
        client = undefined;
    }
}

async function restartLspServer(): Promise<void> {
    await stopLspServer();
    await startLspServer();
    vscode.window.showInformationMessage('LisPython LSP server restarted.');
}

// ---------------------------------------------------------------------------
// Run / REPL terminal integration
// ---------------------------------------------------------------------------

let replTerminal: vscode.Terminal | undefined;

function getLpyCommand(): string {
    const config = vscode.workspace.getConfiguration('lispython');
    const explicit = config.get<string>('lpyPath', '');
    if (explicit) return explicit;

    // Prefer the venv's lpy if we're using a venv Python
    const pythonPath = config.get<string>('lsp.pythonPath', '');
    if (pythonPath && pythonPath.includes('/')) {
        const lpyInVenv = path.join(path.dirname(pythonPath), 'lpy');
        if (fs.existsSync(lpyInVenv)) return lpyInVenv;
    }

    return 'lpy';
}

function getOrCreateReplTerminal(): vscode.Terminal {
    if (replTerminal && replTerminal.exitStatus === undefined) {
        return replTerminal;
    }
    replTerminal = vscode.window.createTerminal({
        name: 'LisPython REPL',
        shellPath: getLpyCommand(),
    });
    return replTerminal;
}

async function runFile(): Promise<void> {
    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document.languageId !== 'lispython') {
        vscode.window.showErrorMessage('No LisPython file is active.');
        return;
    }

    // Save first so the file on disk is current
    if (editor.document.isDirty) {
        await editor.document.save();
    }

    const filePath = editor.document.uri.fsPath;
    const lpy = getLpyCommand();
    const terminal = vscode.window.createTerminal({
        name: `Run: ${path.basename(filePath)}`,
        cwd: path.dirname(filePath),
    });
    terminal.show();
    terminal.sendText(`${lpy} ${JSON.stringify(filePath)}`);
}

async function startRepl(): Promise<void> {
    const terminal = getOrCreateReplTerminal();
    terminal.show();
}

// --- S-expression finding ---

const OPEN = new Set(['(', '[', '{']);
const CLOSE = new Set([')', ']', '}']);

/**
 * Given the full source and a character offset, scan forward from `start`
 * past whitespace/comments and find the end of the next s-expression.
 * Returns the offset just after the sexp, or -1 if none.
 */
function sexpEndFrom(src: string, start: number): number {
    let i = start;
    // Skip whitespace and comments
    while (i < src.length) {
        const c = src[i];
        if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { i++; continue; }
        if (c === ';') { while (i < src.length && src[i] !== '\n') i++; continue; }
        break;
    }
    if (i >= src.length) return -1;

    const c = src[i];

    // String
    if (c === '"') {
        i++;
        while (i < src.length) {
            if (src[i] === '\\') { i += 2; continue; }
            if (src[i] === '"') { i++; break; }
            i++;
        }
        return i;
    }

    // Paren/bracket/brace form
    if (OPEN.has(c)) {
        const stack: string[] = [c];
        i++;
        while (i < src.length && stack.length > 0) {
            const ch = src[i];
            if (ch === ';') { while (i < src.length && src[i] !== '\n') i++; continue; }
            if (ch === '"') {
                i++;
                while (i < src.length) {
                    if (src[i] === '\\') { i += 2; continue; }
                    if (src[i] === '"') { i++; break; }
                    i++;
                }
                continue;
            }
            if (OPEN.has(ch)) { stack.push(ch); i++; continue; }
            if (CLOSE.has(ch)) { stack.pop(); i++; continue; }
            i++;
        }
        return i;
    }

    // Quote/quasiquote/unquote prefix — include it in the sexp
    if (c === "'" || c === '`' || c === '~') {
        i++;
        if (src[i] === '@') i++; // unquote-splice ~@
        return sexpEndFrom(src, i);
    }

    // Atom/symbol/number
    while (i < src.length) {
        const ch = src[i];
        if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') break;
        if (OPEN.has(ch) || CLOSE.has(ch)) break;
        if (ch === ';') break;
        i++;
    }
    return i;
}

/**
 * Find the s-expression ending at or just before `offset`.
 * Returns [start, end] or null.
 */
function lastSexpBefore(src: string, offset: number): [number, number] | null {
    // Back up past trailing whitespace from the cursor
    let end = offset;
    while (end > 0) {
        const c = src[end - 1];
        if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { end--; continue; }
        break;
    }
    if (end === 0) return null;

    // We need to find the start of the sexp that ends at `end`.
    // Scan backward tracking paren balance in reverse.
    let i = end - 1;
    const stack: string[] = [];
    let inString = false;

    // If we're looking at a close-bracket, jump to its matching open
    const last = src[i];
    if (CLOSE.has(last)) {
        stack.push(last);
        i--;
        while (i >= 0 && stack.length > 0) {
            const ch = src[i];
            if (ch === '"') {
                // Walk back over string
                i--;
                while (i >= 0) {
                    if (src[i] === '"' && (i === 0 || src[i - 1] !== '\\')) { i--; break; }
                    i--;
                }
                continue;
            }
            if (CLOSE.has(ch)) { stack.push(ch); i--; continue; }
            if (OPEN.has(ch)) { stack.pop(); i--; continue; }
            i--;
        }
        let start = i + 1;
        // Include meta-prefix (', `, ~, ~@) before the open-bracket
        while (start > 0) {
            const pc = src[start - 1];
            if (pc === "'" || pc === '`' || pc === '~' || pc === '@') { start--; continue; }
            break;
        }
        return [start, end];
    }

    // Close-quote: walk back to matching open-quote
    if (last === '"') {
        i--;
        while (i >= 0) {
            if (src[i] === '"' && (i === 0 || src[i - 1] !== '\\')) break;
            i--;
        }
        return [i, end];
    }

    // Atom: walk back until whitespace or bracket
    while (i >= 0) {
        const ch = src[i];
        if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') break;
        if (OPEN.has(ch) || CLOSE.has(ch)) break;
        i--;
    }
    return [i + 1, end];
}

/**
 * Find the outermost (top-level) s-expression containing the given offset.
 * Returns [start, end] of the form, or null if none.
 */
function topLevelSexpAt(src: string, offset: number): [number, number] | null {
    // Parse from start tracking top-level forms; find the one containing offset.
    let i = 0;
    while (i < src.length) {
        // Skip whitespace and comments
        const c = src[i];
        if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { i++; continue; }
        if (c === ';') { while (i < src.length && src[i] !== '\n') i++; continue; }

        const start = i;
        const end = sexpEndFrom(src, i);
        if (end < 0 || end === i) { i++; continue; }

        if (offset >= start && offset <= end) {
            return [start, end];
        }
        i = end;
    }
    return null;
}

async function sendSelectionToRepl(): Promise<void> {
    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document.languageId !== 'lispython') {
        vscode.window.showErrorMessage('No LisPython file is active.');
        return;
    }

    const selection = editor.selection;
    if (!selection.isEmpty) {
        const text = editor.document.getText(selection);
        if (!text.trim()) return;
        const range: [number, number] = [
            editor.document.offsetAt(selection.start),
            editor.document.offsetAt(selection.end),
        ];
        sendTextToRepl(text, editor, range);
        return;
    }

    vscode.window.showWarningMessage(
        'No selection. Use "Send Last S-Expression" or "Send Top-Level Form" instead.'
    );
}

async function sendLastSexp(): Promise<void> {
    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document.languageId !== 'lispython') return;

    const src = editor.document.getText();
    const offset = editor.document.offsetAt(editor.selection.active);
    const range = lastSexpBefore(src, offset);
    if (!range) {
        vscode.window.showInformationMessage('No s-expression before cursor.');
        return;
    }
    const text = src.slice(range[0], range[1]);
    sendTextToRepl(text, editor, range);
    flashRange(editor, range);
}

async function sendTopLevelForm(): Promise<void> {
    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document.languageId !== 'lispython') return;

    const src = editor.document.getText();
    const offset = editor.document.offsetAt(editor.selection.active);
    const range = topLevelSexpAt(src, offset);
    if (!range) {
        vscode.window.showInformationMessage('No top-level form at cursor.');
        return;
    }
    const text = src.slice(range[0], range[1]);
    sendTextToRepl(text, editor, range);
    flashRange(editor, range);
}

async function sendTextToRepl(text: string, editor?: vscode.TextEditor, range?: [number, number]): Promise<void> {
    if (!editor || !range) return;
    if (!nreplSocket || nreplSocket.destroyed) {
        await startNrepl();
    }
    if (nreplSocket && !nreplSocket.destroyed) {
        evalAndShow(text, editor, range);
    }
}

function flashRange(editor: vscode.TextEditor, range: [number, number]): void {
    const start = editor.document.positionAt(range[0]);
    const end = editor.document.positionAt(range[1]);
    const deco = vscode.window.createTextEditorDecorationType({
        backgroundColor: new vscode.ThemeColor('editor.findMatchHighlightBackground'),
    });
    editor.setDecorations(deco, [new vscode.Range(start, end)]);
    setTimeout(() => deco.dispose(), 200);
}

// ---------------------------------------------------------------------------
// nREPL client
// ---------------------------------------------------------------------------

let nreplSocket: net.Socket | undefined;
let nreplProcess: ChildProcess | undefined;
let nreplOutputChannel: vscode.OutputChannel | undefined;
let nreplStatusItem: vscode.StatusBarItem | undefined;

const resultDecorationType = vscode.window.createTextEditorDecorationType({
    after: {
        margin: '0 0 0 1em',
        fontStyle: 'italic',
    },
    isWholeLine: false,
});

let nreplPanel: vscode.WebviewPanel | undefined;

function getNreplOutputChannel(): vscode.OutputChannel {
    if (!nreplOutputChannel) {
        nreplOutputChannel = vscode.window.createOutputChannel('LisPython nREPL');
    }
    return nreplOutputChannel;
}

function getOrCreateReplPanel(): vscode.WebviewPanel {
    if (nreplPanel) {
        return nreplPanel;
    }
    nreplPanel = vscode.window.createWebviewPanel(
        'lispythonRepl',
        'LisPython REPL',
        { viewColumn: vscode.ViewColumn.Two, preserveFocus: true },
        { enableScripts: true, retainContextWhenHidden: true },
    );
    nreplPanel.webview.html = getReplHtml();
    nreplPanel.onDidDispose(() => { nreplPanel = undefined; });
    return nreplPanel;
}

function appendToReplPanel(code: string, result: { value?: string | null; stdout?: string; error?: string | null }): void {
    const panel = getOrCreateReplPanel();
    panel.webview.postMessage({ type: 'eval', code, result });
}

function getReplHtml(): string {
    return `<!DOCTYPE html>
<html>
<head>
<style>
  body {
    font-family: var(--vscode-editor-font-family, monospace);
    font-size: var(--vscode-editor-font-size, 13px);
    padding: 8px;
    color: var(--vscode-editor-foreground);
    background: var(--vscode-editor-background);
  }
  .entry { margin-bottom: 12px; border-bottom: 1px solid var(--vscode-editorGroup-border); padding-bottom: 8px; }
  .code { color: var(--vscode-textPreformat-foreground); white-space: pre-wrap; }
  .prompt { color: var(--vscode-descriptionForeground); }
  .value { color: var(--vscode-debugTokenExpression-number); white-space: pre-wrap; }
  .stdout { color: var(--vscode-terminal-ansiGreen); white-space: pre-wrap; }
  .error { color: var(--vscode-errorForeground); white-space: pre-wrap; }
  #entries { overflow-y: auto; }
</style>
</head>
<body>
<div id="entries"></div>
<script>
  const entries = document.getElementById('entries');
  window.addEventListener('message', (e) => {
    const { type, code, result } = e.data;
    if (type === 'clear') { entries.innerHTML = ''; return; }
    if (type !== 'eval') return;
    const div = document.createElement('div');
    div.className = 'entry';

    let html = '<span class="prompt">repl&gt; </span><span class="code">' + escapeHtml(code) + '</span>';

    if (result.stdout) {
      html += '<br><span class="stdout">' + escapeHtml(result.stdout.replace(/\\n$/, '')) + '</span>';
    }
    if (result.error) {
      html += '<br><span class="error">' + escapeHtml(result.error) + '</span>';
    } else if (result.value && result.value !== 'None') {
      html += '<br><span class="value">' + escapeHtml(result.value) + '</span>';
    }

    div.innerHTML = html;
    entries.appendChild(div);
    div.scrollIntoView({ behavior: 'smooth' });
  });

  function escapeHtml(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
</script>
</body>
</html>`;
}

function updateNreplStatus(connected: boolean): void {
    if (!nreplStatusItem) {
        nreplStatusItem = vscode.window.createStatusBarItem(
            vscode.StatusBarAlignment.Right, 99
        );
        nreplStatusItem.command = 'lispython.startNrepl';
    }
    if (connected) {
        nreplStatusItem.text = '$(plug) nREPL';
        nreplStatusItem.tooltip = 'LisPython nREPL connected. Click to restart.';
        nreplStatusItem.color = new vscode.ThemeColor('statusBarItem.prominentForeground');
    } else {
        nreplStatusItem.text = '$(debug-disconnect) nREPL';
        nreplStatusItem.tooltip = 'LisPython nREPL disconnected. Click to start.';
        nreplStatusItem.color = undefined;
    }
    nreplStatusItem.show();
}

function connectNrepl(port: number): Promise<void> {
    return new Promise((resolve, reject) => {
        if (nreplSocket) {
            nreplSocket.destroy();
            nreplSocket = undefined;
        }
        const sock = new net.Socket();
        sock.connect(port, '127.0.0.1', () => {
            nreplSocket = sock;
            updateNreplStatus(true);
            getNreplOutputChannel().appendLine(`Connected to nREPL on port ${port}`);
            resolve();
        });
        sock.on('error', (err) => {
            updateNreplStatus(false);
            reject(err);
        });
        sock.on('close', () => {
            nreplSocket = undefined;
            updateNreplStatus(false);
        });
    });
}

function disconnectNrepl(): Promise<void> {
    if (nreplSocket) {
        nreplSocket.destroy();
        nreplSocket = undefined;
    }
    const proc = nreplProcess;
    nreplProcess = undefined;
    updateNreplStatus(false);
    if (!proc || proc.exitCode !== null) {
        return Promise.resolve();
    }
    return new Promise((resolve) => {
        proc.on('exit', () => resolve());
        proc.kill('SIGTERM');
        // Force kill after 2 seconds if still alive
        setTimeout(() => {
            try { proc.kill('SIGKILL'); } catch {}
            resolve();
        }, 2000);
    });
}

async function startNrepl(): Promise<void> {
    await disconnectNrepl();
    if (nreplPanel) {
        nreplPanel.webview.postMessage({ type: 'clear' });
    }

    const lpy = getLpyCommand();
    const cwd = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? process.cwd();

    const out = getNreplOutputChannel();
    out.show(true);
    out.appendLine(`Starting nREPL server: ${lpy} --nrepl`);

    nreplProcess = spawn(lpy, ['--nrepl'], { cwd });

    // Parse port from stdout: "nREPL server started on 127.0.0.1:PORT"
    const portPromise = new Promise<number>((resolve, reject) => {
        const timeout = setTimeout(() => reject(new Error('nREPL server did not start in time')), 10000);
        nreplProcess!.stdout?.on('data', (data: Buffer) => {
            const text = data.toString();
            out.append(text);
            const match = text.match(/started on .+:(\d+)/);
            if (match) {
                clearTimeout(timeout);
                resolve(parseInt(match[1], 10));
            }
        });
        nreplProcess!.on('exit', () => {
            clearTimeout(timeout);
            reject(new Error('nREPL server exited before starting'));
        });
    });

    nreplProcess.stderr?.on('data', (data: Buffer) => {
        out.append(data.toString());
    });
    nreplProcess.on('exit', (code) => {
        out.appendLine(`nREPL server exited (code ${code})`);
        nreplProcess = undefined;
        updateNreplStatus(false);
    });

    try {
        const port = await portPromise;
        await connectNrepl(port);
        out.appendLine(`Connected on port ${port}`);
    } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        vscode.window.showErrorMessage(msg);
    }
}

function nreplEval(code: string): Promise<{ value?: string; stdout?: string; error?: string }> {
    return nreplRequest({ op: 'eval', code });
}

/** Send an arbitrary JSON request to the nREPL and return the parsed response. */
function nreplRequest(msg: Record<string, unknown>): Promise<any> {
    return new Promise((resolve, reject) => {
        if (!nreplSocket || nreplSocket.destroyed) {
            reject(new Error('nREPL not connected'));
            return;
        }

        let buf = '';
        const onData = (data: Buffer) => {
            buf += data.toString();
            const newlineIdx = buf.indexOf('\n');
            if (newlineIdx >= 0) {
                nreplSocket!.removeListener('data', onData);
                try {
                    resolve(JSON.parse(buf.slice(0, newlineIdx)));
                } catch (e) {
                    reject(e);
                }
            }
        };
        nreplSocket.on('data', onData);
        nreplSocket.write(JSON.stringify(msg) + '\n');
    });
}

async function evalAndShow(code: string, editor: vscode.TextEditor, range: [number, number]): Promise<void> {
    try {
        const result = await nreplEval(code);
        appendToReplPanel(code, result);

        // Also show inline decoration briefly
        const parts: string[] = [];
        if (result.error) {
            parts.push(`Error: ${result.error}`);
        } else if (result.value && result.value !== 'None') {
            parts.push(result.value);
        } else if (result.stdout) {
            parts.push(result.stdout.trimEnd());
        }

        if (parts.length > 0) {
            const msg = parts.join(' | ');
            const endPos = editor.document.positionAt(range[1]);
            const line = endPos.line;
            const lineEnd = editor.document.lineAt(line).range.end;

            const color = result.error
                ? new vscode.ThemeColor('errorForeground')
                : new vscode.ThemeColor('editorInfo.foreground');

            const decoration: vscode.DecorationOptions = {
                range: new vscode.Range(lineEnd, lineEnd),
                renderOptions: {
                    after: {
                        contentText: ` => ${msg}`,
                        color,
                    },
                },
            };
            editor.setDecorations(resultDecorationType, [decoration]);

            setTimeout(() => {
                if (vscode.window.activeTextEditor === editor) {
                    editor.setDecorations(resultDecorationType, []);
                }
            }, 5000);
        }
    } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        vscode.window.showErrorMessage(`nREPL eval failed: ${msg}`);
    }
}

// ---------------------------------------------------------------------------
// Load File command
// ---------------------------------------------------------------------------

async function loadFileToRepl(): Promise<void> {
    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document.languageId !== 'lispython') {
        vscode.window.showErrorMessage('No LisPython file is active.');
        return;
    }

    if (editor.document.isDirty) {
        await editor.document.save();
    }

    const filePath = editor.document.uri.fsPath;

    // Auto-start nREPL if not connected
    if (!nreplSocket || nreplSocket.destroyed) {
        await startNrepl();
    }
    if (!nreplSocket || nreplSocket.destroyed) {
        vscode.window.showErrorMessage('Could not connect to nREPL.');
        return;
    }

    try {
        const result = await nreplRequest({ op: 'load-file', path: filePath });
        appendToReplPanel(`(load-file "${path.basename(filePath)}")`, {
            value: result.value,
            stdout: result.stdout,
            error: result.error,
        });
    } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        vscode.window.showErrorMessage(`nREPL load-file failed: ${msg}`);
    }
}

// ---------------------------------------------------------------------------
// Macroexpand command
// ---------------------------------------------------------------------------

async function macroexpand(): Promise<void> {
    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document.languageId !== 'lispython') {
        vscode.window.showErrorMessage('No LisPython file is active.');
        return;
    }

    const src = editor.document.getText();
    const offset = editor.document.offsetAt(editor.selection.active);

    // Try top-level form first, then fall back to last sexp before cursor
    const range = topLevelSexpAt(src, offset) ?? lastSexpBefore(src, offset);
    if (!range) {
        vscode.window.showInformationMessage('No s-expression at cursor.');
        return;
    }

    const code = src.slice(range[0], range[1]);
    flashRange(editor, range);

    // Auto-start nREPL if not connected
    if (!nreplSocket || nreplSocket.destroyed) {
        await startNrepl();
    }
    if (!nreplSocket || nreplSocket.destroyed) {
        vscode.window.showErrorMessage('Could not connect to nREPL.');
        return;
    }

    try {
        const result = await nreplRequest({ op: 'macroexpand', code });
        appendToReplPanel(`(macroexpand '${code})`, {
            value: result.expansion,
            error: result.error,
        });
    } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        vscode.window.showErrorMessage(`nREPL macroexpand failed: ${msg}`);
    }
}

// ---------------------------------------------------------------------------
// Completion provider (nREPL-backed)
// ---------------------------------------------------------------------------

class NreplCompletionProvider implements vscode.CompletionItemProvider {
    async provideCompletionItems(
        document: vscode.TextDocument,
        position: vscode.Position,
        _token: vscode.CancellationToken,
    ): Promise<vscode.CompletionItem[] | undefined> {
        // Only provide completions when nREPL is connected
        if (!nreplSocket || nreplSocket.destroyed) {
            return undefined;
        }

        // Extract the prefix: walk backwards from cursor over symbol characters
        const lineText = document.lineAt(position.line).text;
        let start = position.character;
        while (start > 0) {
            const ch = lineText[start - 1];
            if (/[\s()\[\]{}'`,;"]/.test(ch)) break;
            start--;
        }
        const prefix = lineText.slice(start, position.character);
        if (!prefix) return undefined;

        try {
            const result = await nreplRequest({ op: 'complete', prefix });
            if (!result.completions || !Array.isArray(result.completions)) {
                return undefined;
            }
            return result.completions.map((name: string) => {
                const item = new vscode.CompletionItem(name, vscode.CompletionItemKind.Function);
                // Replace the prefix range so the completion replaces what was typed
                item.range = new vscode.Range(
                    position.line, start,
                    position.line, position.character,
                );
                return item;
            });
        } catch {
            return undefined;
        }
    }
}

// ---------------------------------------------------------------------------
// Auto-indent on Enter (bracket-depth aware)
// ---------------------------------------------------------------------------

const SPECIAL_FORMS = new Set([
    'def', 'defn', 'defmacro', 'async-def',
    'if', 'cond', 'conde', 'when', 'unless',
    'while', 'for', 'async-for',
    'do', 'try', 'with', 'async-with',
    'class', 'deco',
    'lambda', 'fn',
    'let', 'match', 'case',
    'except', 'except*', 'finally', 'else',
]);

function calcIndent(text: string): number {
    let inString = false;
    const stack: number[] = [];
    let i = 0;
    const n = text.length;

    while (i < n) {
        const c = text[i];
        if (inString) {
            if (c === '\\') { i += 2; continue; }
            if (c === '"') { inString = false; }
            i++; continue;
        }
        if (c === '"') { inString = true; i++; continue; }
        if (c === ';') { while (i < n && text[i] !== '\n') i++; continue; }

        if (c === '(' || c === '[' || c === '{') {
            const parenCol = colOf(text, i);

            // For [ and {: align to first element
            if (c !== '(') {
                let j = i + 1;
                while (j < n && text[j] === ' ') j++;
                if (j < n && text[j] !== '\n') {
                    stack.push(colOf(text, j));
                } else {
                    stack.push(parenCol + 1);
                }
            } else {
                // For (: special form vs function call
                let j = i + 1;
                while (j < n && text[j] === ' ') j++;
                const opStart = j;
                while (j < n && !' \n\t()[]{}\'\"'.includes(text[j])) j++;
                const op = text.slice(opStart, j);

                if (op === '' || SPECIAL_FORMS.has(op)) {
                    stack.push(parenCol + 2);
                } else {
                    let argPos = j;
                    while (argPos < n && text[argPos] === ' ') argPos++;
                    if (argPos < n && text[argPos] !== '\n') {
                        stack.push(colOf(text, argPos));
                    } else {
                        stack.push(parenCol + 2);
                    }
                }
            }
        } else if (c === ')' || c === ']' || c === '}') {
            if (stack.length > 0) stack.pop();
        }
        i++;
    }
    return stack.length > 0 ? stack[stack.length - 1] : 0;
}

function colOf(text: string, pos: number): number {
    const nl = text.lastIndexOf('\n', pos - 1);
    return pos - nl - 1;
}

async function newlineAndIndent(): Promise<void> {
    const editor = vscode.window.activeTextEditor;
    if (!editor) return;
    const pos = editor.selection.active;
    const textBefore = editor.document.getText(new vscode.Range(0, 0, pos.line, pos.character));
    const indent = calcIndent(textBefore);
    const newText = '\n' + ' '.repeat(indent);
    // Consume whitespace after cursor
    const lineText = editor.document.lineAt(pos.line).text;
    let end = pos.character;
    while (end < lineText.length && lineText[end] === ' ') end++;
    await editor.edit((edit) => {
        edit.replace(new vscode.Range(pos.line, pos.character, pos.line, end), newText);
    });
    const newPos = new vscode.Position(pos.line + 1, indent);
    editor.selection = new vscode.Selection(newPos, newPos);
}

async function reindentLines(): Promise<void> {
    const editor = vscode.window.activeTextEditor;
    if (!editor) return;
    const sel = editor.selection;
    const startLine = sel.start.line;
    const endLine = sel.isEmpty ? startLine : sel.end.line;

    await editor.edit((edit) => {
        for (let line = startLine; line <= endLine; line++) {
            const textBefore = editor.document.getText(new vscode.Range(0, 0, line, 0));
            const indent = calcIndent(textBefore);
            const lineObj = editor.document.lineAt(line);
            const oldIndent = lineObj.firstNonWhitespaceCharacterIndex;
            if (oldIndent !== indent) {
                edit.replace(
                    new vscode.Range(line, 0, line, oldIndent),
                    ' '.repeat(indent),
                );
            }
        }
    });
}

// ---------------------------------------------------------------------------
// Activation / deactivation
// ---------------------------------------------------------------------------

export async function activate(context: vscode.ExtensionContext) {
    // Register commands FIRST so they work even if LSP fails to start
    context.subscriptions.push(
        vscode.commands.registerCommand('lispython.selectPythonPath', selectPythonPath),
        vscode.commands.registerCommand('lispython.restartServer', restartLspServer),
        vscode.commands.registerCommand('lispython.runFile', runFile),
        vscode.commands.registerCommand('lispython.startRepl', startRepl),
        vscode.commands.registerCommand('lispython.sendToRepl', sendSelectionToRepl),
        vscode.commands.registerCommand('lispython.sendLastSexp', sendLastSexp),
        vscode.commands.registerCommand('lispython.sendTopLevelForm', sendTopLevelForm),
        vscode.commands.registerCommand('lispython.startNrepl', startNrepl),
        vscode.commands.registerCommand('lispython.disconnectNrepl', disconnectNrepl),
        vscode.commands.registerCommand('lispython.loadFileToRepl', loadFileToRepl),
        vscode.commands.registerCommand('lispython.macroexpand', macroexpand),
    );

    // Completion provider (nREPL-backed)
    context.subscriptions.push(
        vscode.languages.registerCompletionItemProvider(
            { language: 'lispython', scheme: 'file' },
            new NreplCompletionProvider(),
            '.',
        ),
    );

    // Auto-indent
    context.subscriptions.push(
        vscode.commands.registerCommand('lispython.newlineAndIndent', newlineAndIndent),
        vscode.commands.registerCommand('lispython.reindentLines', reindentLines),
    );

    // Status bar
    createStatusBar(context);
    updateNreplStatus(false);
    context.subscriptions.push(nreplStatusItem!);

    // Watch for config changes
    context.subscriptions.push(
        vscode.workspace.onDidChangeConfiguration((e) => {
            if (e.affectsConfiguration('lispython.lsp.pythonPath')) {
                updateStatusBar();
            }
        })
    );

    // Start LSP (don't let failure break activation)
    try {
        await startLspServer();
    } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        const pick = 'Select Interpreter';
        const choice = await vscode.window.showWarningMessage(
            `LisPython LSP: ${msg}`,
            pick,
        );
        if (choice === pick) {
            vscode.commands.executeCommand('lispython.selectPythonPath');
        }
    }
}

export function deactivate(): Thenable<void> | undefined {
    disconnectNrepl();
    nreplStatusItem?.dispose();
    nreplOutputChannel?.dispose();
    return stopLspServer();
}
