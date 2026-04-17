import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { execFile } from 'child_process';
import {
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

    const serverOptions: ServerOptions = {
        command: serverCommand,
        args: serverArgs,
    };

    const clientOptions: LanguageClientOptions = {
        documentSelector: [{ scheme: 'file', language: 'lispython' }],
        synchronize: {
            fileEvents: vscode.workspace.createFileSystemWatcher('**/*.lpy'),
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
        sendTextToRepl(text);
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
    sendTextToRepl(text);
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
    sendTextToRepl(text);
    flashRange(editor, range);
}

function sendTextToRepl(text: string): void {
    const terminal = getOrCreateReplTerminal();
    terminal.show(true);
    terminal.sendText(text);
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
    );

    // Status bar
    createStatusBar(context);

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
        vscode.window.showWarningMessage(
            `LisPython LSP failed to start: ${err}. Click the status bar to select a Python interpreter with lispython installed.`
        );
    }
}

export function deactivate(): Thenable<void> | undefined {
    return stopLspServer();
}
