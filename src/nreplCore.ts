// Pure nREPL client helpers (no `vscode` runtime import) so they run under
// `node --test` without an extension host.
import type { CompletionItemKind } from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import * as crypto from 'crypto';
import * as net from 'net';
import { spawn, ChildProcess } from 'child_process';

// Mirrors lpy-mode's kind table; vscode has no macro kind, so macros read as functions.
export const COMPLETION_KINDS: Record<string, keyof typeof CompletionItemKind> = {
    class: 'Class',
    function: 'Function',
    variable: 'Variable',
    macro: 'Function',
    keyword: 'Keyword',
};

export const FALLBACK_COMPLETION_KIND: keyof typeof CompletionItemKind = 'Function';

function own(obj: object, key: string): boolean {
    return Object.prototype.hasOwnProperty.call(obj, key);
}

export function completionKindName(
    kinds: Record<string, string> | undefined,
    completion: string,
): keyof typeof CompletionItemKind {
    const kind = kinds && own(kinds, completion) ? kinds[completion] : undefined;
    return kind !== undefined && own(COMPLETION_KINDS, kind)
        ? COMPLETION_KINDS[kind]
        : FALLBACK_COMPLETION_KIND;
}

/** The server writes bare decimal digits, no newline. */
export function parsePortFile(text: string): number | undefined {
    const t = text.trim();
    if (!/^\d+$/.test(t)) return undefined;
    const port = parseInt(t, 10);
    return port >= 1 && port <= 65535 ? port : undefined;
}

export const NREPL_PORT_FILE = '.nrepl-port';

/**
 * Look for `.nrepl-port` in `startDir` and each parent directory. Returns the nearest
 * file found, with `port` undefined when that file is malformed.
 */
function locateNreplPortFile(startDir: string): { file: string; port: number | undefined } | undefined {
    let dir = path.resolve(startDir);
    for (;;) {
        const file = path.join(dir, NREPL_PORT_FILE);
        let text: string | undefined;
        try {
            text = fs.readFileSync(file, 'utf8');
        } catch { /* absent or unreadable: keep walking up */ }
        // The nearest file wins; a malformed one must not fall through to an unrelated server above.
        if (text !== undefined) return { file, port: parsePortFile(text) };
        const parent = path.dirname(dir);
        if (parent === dir) return undefined;
        dir = parent;
    }
}

/** Look for `.nrepl-port` in `startDir` and each parent directory. */
export function findNreplPort(startDir: string): number | undefined {
    return locateNreplPortFile(startDir)?.port;
}

/**
 * The `.nrepl-port` to connect to. With an active file, only the search upward from its
 * directory counts, so a sibling root's server is never picked for it; otherwise each
 * workspace folder is searched in order and the first valid port wins.
 */
export function findNreplPortFor(
    activeFile: string | undefined,
    folders: readonly string[],
): { port: number; file: string } | undefined {
    const starts = activeFile !== undefined ? [path.dirname(activeFile)] : folders;
    for (const start of starts) {
        const found = locateNreplPortFile(start);
        if (found?.port !== undefined) return { port: found.port, file: found.file };
    }
    return undefined;
}

/**
 * The file the nREPL commands act for. A focused file editor is it; a focused untitled
 * editor is a deliberate choice of no file. With no text editor focused (e.g. the REPL
 * webview) or a virtual one (the nREPL output channel, git/diff views), the user is still
 * working on the last LisPython file in focus.
 */
export function resolveActiveFile(
    editor: { scheme: string; fsPath: string } | undefined,
    lastLispyFile: string | undefined,
): string | undefined {
    if (editor?.scheme === 'file') return editor.fsPath;
    if (editor?.scheme === 'untitled') return undefined;
    return lastLispyFile;
}

/**
 * Where the nREPL lives: the active file's workspace folder, else folder 0. A file outside
 * every workspace folder counts only when no folder is open. startNrepl spawns the server
 * in `folder` (so it writes .nrepl-port there) and the connect command searches from
 * `activeFile`, else across the folders, so both agree on one location.
 */
export function chooseNreplHome(
    activeFile: string | undefined,
    owningFolder: string | undefined,
    folders: readonly string[],
): { activeFile: string | undefined; folder: string | undefined } {
    if (activeFile !== undefined && owningFolder !== undefined) return { activeFile, folder: owningFolder };
    if (folders.length > 0) return { activeFile: undefined, folder: folders[0] };
    if (activeFile !== undefined) return { activeFile, folder: path.dirname(activeFile) };
    return { activeFile: undefined, folder: undefined };
}

/**
 * Send one newline-delimited JSON request on `sock` and resolve with the first response
 * line. Rejects if the socket errors or closes first (e.g. the server died or the
 * connection was swapped), so a caller never waits forever; every listener is removed
 * once the request settles. This is the only reader of the socket, so it puts the socket
 * in utf8 mode once.
 */
export function nreplRequestOn(sock: net.Socket, msg: Record<string, unknown>): Promise<any> {
    return new Promise((resolve, reject) => {
        if (sock.destroyed) {
            reject(new Error('nREPL not connected'));
            return;
        }
        // One StringDecoder per socket, not per chunk: a multi-byte character split across
        // reads (or across requests) is carried over instead of decoding to U+FFFD.
        if (sock.readableEncoding !== 'utf8') sock.setEncoding('utf8');
        let buf = '';
        const cleanup = () => {
            sock.removeListener('data', onData);
            sock.removeListener('close', onClose);
            sock.removeListener('error', onError);
        };
        const onData = (data: string) => {
            buf += data;
            const newlineIdx = buf.indexOf('\n');
            if (newlineIdx < 0) return;
            cleanup();
            try {
                resolve(JSON.parse(buf.slice(0, newlineIdx)));
            } catch (e) {
                reject(e);
            }
        };
        const onClose = () => {
            cleanup();
            reject(new Error('nREPL connection closed'));
        };
        const onError = (err: Error) => {
            cleanup();
            reject(new Error(`nREPL connection error: ${err.message}`));
        };
        sock.on('data', onData);
        sock.on('close', onClose);
        sock.on('error', onError);
        sock.write(JSON.stringify(msg) + '\n');
    });
}

/** Rejection reason of a connect() superseded by a later connect() or by disconnect(). */
export class NreplConnectCancelled extends Error {
    constructor() {
        super('nREPL connect superseded');
        this.name = 'NreplConnectCancelled';
    }
}

/**
 * The extension's single nREPL socket. A connect opens a new socket as `pending` and
 * swaps it in only once it has connected, so a stale port (a leftover `.nrepl-port`)
 * fails without dropping a live session. At most one connect is pending: a newer one
 * cancels the older, so overlapping connects cannot leak a socket. Status follows the
 * current socket only; a pending socket's failure never reports "disconnected".
 */
export class NreplConnection {
    private current: net.Socket | undefined;
    private pending: { sock: net.Socket; reject: (err: Error) => void } | undefined;

    constructor(private readonly onStatus: (connected: boolean) => void) {}

    get socket(): net.Socket | undefined {
        return this.current;
    }

    connect(port: number, host = '127.0.0.1'): Promise<void> {
        this.cancelPending();
        return new Promise((resolve, reject) => {
            const sock = new net.Socket();
            const entry = { sock, reject };
            this.pending = entry;
            sock.once('connect', () => {
                if (this.pending !== entry) {
                    sock.destroy();
                    return;
                }
                this.pending = undefined;
                const old = this.current;
                this.current = sock;
                old?.destroy();
                this.onStatus(true);
                resolve();
            });
            // Errors on the current socket surface through its 'close'.
            sock.on('error', (err) => {
                if (this.pending !== entry) return;
                this.pending = undefined;
                reject(err);
            });
            sock.on('close', () => {
                if (this.pending === entry) {
                    this.pending = undefined;
                    reject(new Error('nREPL socket closed before connecting'));
                }
                if (this.current !== sock) return;
                this.current = undefined;
                this.onStatus(false);
            });
            sock.connect(port, host);
        });
    }

    disconnect(): void {
        this.cancelPending();
        const old = this.current;
        this.current = undefined;
        old?.destroy();
        this.onStatus(false);
    }

    private cancelPending(): void {
        const p = this.pending;
        if (!p) return;
        this.pending = undefined;
        p.sock.destroy();
        p.reject(new NreplConnectCancelled());
    }
}

// escapeHtml and linkifyErrors are injected into the REPL webview script via
// Function.prototype.toString (see getReplHtml below), so they must stay
// self-contained: linkifyErrors may reference only escapeHtml, by name. A bundler or
// minifier (esbuild/webpack/terser) would rename or inline them and break the
// injection; the getReplHtml script test in test/ pins this only for the tsc output
// in out/. If a bundle ever becomes `main`, the test must run against the bundle.

/** Escapes text for both element content and double/single-quoted attribute values. */
export function escapeHtml(s: string): string {
    return s
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/**
 * Escapes an error message and wraps each `File "<path>", line N` in a clickable
 * span. Matching runs on the raw text so the path is captured unescaped, then every
 * piece (link text, attribute values, surrounding text) is escaped exactly once.
 */
export function linkifyErrors(text: string): string {
    const re = /File "(.+?)", line (\d+)/g;
    let out = '';
    let last = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(text)) !== null) {
        out += escapeHtml(text.slice(last, m.index));
        out += '<span class="file-link" data-file="' + escapeHtml(m[1]) +
            '" data-line="' + m[2] + '">' + escapeHtml(m[0]) + '</span>';
        last = m.index + m[0].length;
    }
    return out + escapeHtml(text.slice(last));
}

/**
 * VS Code notifications render `[label](command:...)` as clickable links, and its
 * linkedText parser has no escape syntax, so adjacency is broken instead: a
 * zero-width space between `]` and `(` keeps the link regex from matching.
 */
export function neutralizeNotificationLinks(text: string): string {
    return text.replace(/\]\(/g, ']\u200B(');
}

/** Warning text for a file that could not be opened; the path is untrusted REPL output. */
export function openFileWarning(file: unknown): string {
    return `Could not open file: "${neutralizeNotificationLinks(String(file))}"`;
}

/** Fresh per-render CSP nonce. */
export function makeNonce(): string {
    return crypto.randomBytes(16).toString('base64');
}

export function getReplHtml(nonce: string, cspSource: string): string {
    return `<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${cspSource} 'nonce-${nonce}'; script-src 'nonce-${nonce}';">
<style nonce="${nonce}">
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
  .file-link { color: var(--vscode-textLink-foreground); cursor: pointer; text-decoration: underline; }
  .file-link:hover { color: var(--vscode-textLink-activeForeground); }
  #entries { overflow-y: auto; }
</style>
</head>
<body>
<div id="entries"></div>
<script nonce="${nonce}">
  const entries = document.getElementById('entries');
  const vscode = acquireVsCodeApi();
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
      html += '<br><span class="error">' + linkifyErrors(result.error) + '</span>';
    } else if (result.value && result.value !== 'None') {
      html += '<br><span class="value">' + escapeHtml(result.value) + '</span>';
    }

    div.innerHTML = html;
    entries.appendChild(div);
    div.scrollIntoView({ behavior: 'smooth' });
  });

  ${escapeHtml.toString()}

  ${linkifyErrors.toString()}

  document.addEventListener('click', (e) => {
    const link = e.target.closest('.file-link');
    if (link) {
      vscode.postMessage({ type: 'openFile', file: link.dataset.file, line: parseInt(link.dataset.line, 10) });
    }
  });
</script>
</body>
</html>`;
}

export const NREPL_START_TIMEOUT_MS = 10000;

export interface NreplServerHandlers {
    onStdout(text: string): void;
    onStderr(text: string): void;
}

// Spawns `<lpy> --nrepl` and resolves `port` from its "started on host:PORT" line.
// Every settle path (port found, exit, spawn error) clears the timer, so none is
// left pending. Without the 'error' listener a missing binary (ENOENT) or EACCES
// would be unhandled and only surface as the start timeout.
export function startNreplServer(
    lpy: string,
    cwd: string,
    handlers: NreplServerHandlers,
    startTimeoutMs = NREPL_START_TIMEOUT_MS,
): { proc: ChildProcess; port: Promise<number> } {
    const proc = spawn(lpy, ['--nrepl'], { cwd });
    // Decoder per stream, not per chunk, so a character split across reads survives.
    proc.stdout?.setEncoding('utf8');
    proc.stderr?.setEncoding('utf8');
    const port = new Promise<number>((resolve, reject) => {
        const timeout = setTimeout(
            () => reject(new Error('nREPL server did not start in time')), startTimeoutMs);
        proc.stdout?.on('data', (text: string) => {
            handlers.onStdout(text);
            const match = text.match(/started on .+:(\d+)/);
            if (match) {
                clearTimeout(timeout);
                resolve(parseInt(match[1], 10));
            }
        });
        proc.stderr?.on('data', (text: string) => handlers.onStderr(text));
        proc.on('exit', () => {
            clearTimeout(timeout);
            reject(new Error('nREPL server exited before starting'));
        });
        proc.on('error', (err) => {
            clearTimeout(timeout);
            reject(new Error(`Failed to start nREPL server '${lpy}' (cwd '${cwd}'): ${err.message}`));
        });
    });
    return { proc, port };
}


export interface NreplStartOptions extends NreplServerHandlers {
    lpy: string;
    cwd: string;
    /** Runs inside the serialized section, after the previous server is stopped. */
    beforeSpawn?(): void;
    onExit?(code: number | null): void;
    connect(port: number): Promise<void>;
    startTimeoutMs?: number;
}

/**
 * Owns the one nREPL server this extension spawned. Starts are serialized: a start while
 * another is in flight shares its promise instead of spawning, since the second spawn
 * would replace the first process unkilled. A start whose port or connect step fails
 * kills its server, which would otherwise outlive the failure with nothing to stop it.
 * A stop() while a start is still stopping the previous server cancels that start, so
 * a Disconnect or deactivate never leaves a freshly spawned server behind.
 */
export class NreplServerSupervisor {
    private proc: ChildProcess | undefined;
    private inFlight: Promise<void> | undefined;
    /** Bumped by every stop(); a start that sees it move before spawning gives up. */
    private stops = 0;

    get process(): ChildProcess | undefined {
        return this.proc;
    }

    /** True while a start is in flight; a start() now joins it and ignores its options. */
    get starting(): boolean {
        return this.inFlight !== undefined;
    }

    start(opts: NreplStartOptions): Promise<void> {
        if (this.inFlight) return this.inFlight;
        const run = this.run(opts).finally(() => {
            if (this.inFlight === run) this.inFlight = undefined;
        });
        this.inFlight = run;
        return run;
    }

    private async run(opts: NreplStartOptions): Promise<void> {
        const stopped = this.stop();
        const gen = this.stops;
        await stopped;
        if (gen !== this.stops) throw new NreplConnectCancelled();
        opts.beforeSpawn?.();
        if (gen !== this.stops) throw new NreplConnectCancelled();
        const { proc, port } = startNreplServer(opts.lpy, opts.cwd, opts, opts.startTimeoutMs);
        this.proc = proc;
        // A failed spawn never emits 'exit'; drop the dead process here.
        proc.on('error', () => {
            if (this.proc === proc) this.proc = undefined;
        });
        proc.on('exit', (code) => {
            opts.onExit?.(code);
            // A late exit of a replaced server must not forget its successor.
            if (this.proc === proc) this.proc = undefined;
        });
        try {
            const p = await port;
            // Don't connect to a server a stop() is already killing.
            if (gen !== this.stops) throw new NreplConnectCancelled();
            await opts.connect(p);
        } catch (err) {
            // A stop() since the spawn is a cancel, not a failure: whatever it broke (e.g. the
            // server "exited before starting") must not surface. stop() already killed it.
            if (gen !== this.stops) throw err instanceof NreplConnectCancelled ? err : new NreplConnectCancelled();
            // A superseding connect (e.g. to the .nrepl-port this server wrote) is not a
            // failure of the server, so it stays. Otherwise kill it, only if still current:
            // a stop() or newer start already owns what follows.
            if (!(err instanceof NreplConnectCancelled) && this.proc === proc) await this.stop();
            throw err;
        }
    }

    /** SIGTERM the server, SIGKILL after 2 s; resolves on exit. No-op without a live server. */
    stop(): Promise<void> {
        this.stops++;
        const proc = this.proc;
        this.proc = undefined;
        if (!proc || proc.exitCode !== null || proc.signalCode !== null) {
            return Promise.resolve();
        }
        return new Promise((resolve) => {
            const force = setTimeout(() => {
                try { proc.kill('SIGKILL'); } catch {}
                resolve();
            }, 2000);
            proc.once('exit', () => {
                clearTimeout(force);
                resolve();
            });
            proc.kill('SIGTERM');
        });
    }
}
