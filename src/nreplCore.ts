// Pure nREPL client helpers (no `vscode` runtime import) so they run under
// `node --test` without an extension host.
import type { CompletionItemKind } from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import * as crypto from 'crypto';
import * as net from 'net';
import { spawn, ChildProcess } from 'child_process';

export const FALLBACK_COMPLETION_KIND: keyof typeof CompletionItemKind = 'Function';

// LSP CompletionItemKind is 1-based, VS Code's enum is 0-based: LSP n is VS Code value n-1,
// so this table is indexed by `kind - 1`.
const LSP_COMPLETION_KIND_NAMES = [
    'Text', 'Method', 'Function', 'Constructor', 'Field', 'Variable', 'Class', 'Interface',
    'Module', 'Property', 'Unit', 'Value', 'Enum', 'Keyword', 'Snippet', 'Color', 'File',
    'Reference', 'Folder', 'EnumMember', 'Constant', 'Struct', 'Event', 'Operator',
    'TypeParameter',
] as const satisfies readonly (keyof typeof CompletionItemKind)[];

export function lspCompletionKindName(kind: unknown): keyof typeof CompletionItemKind {
    if (typeof kind !== 'number' || !Number.isInteger(kind)
        || kind < 1 || kind > LSP_COMPLETION_KIND_NAMES.length) {
        return FALLBACK_COMPLETION_KIND;
    }
    return LSP_COMPLETION_KIND_NAMES[kind - 1];
}

export type NreplCompletionItem = {
    label: string;
    kind: keyof typeof CompletionItemKind;
    description?: string;
};

/** Completion items of a `complete` response: the LSP-shaped `items` field only. */
export function nreplCompletionItems(result: Record<string, unknown>): NreplCompletionItem[] {
    if (!Array.isArray(result.items)) return [];
    const out: NreplCompletionItem[] = [];
    for (const raw of result.items as unknown[]) {
        if (typeof raw !== 'object' || raw === null) continue;
        const entry = raw as Record<string, unknown>;
        if (typeof entry.label !== 'string') continue;
        const item: NreplCompletionItem = {
            label: entry.label,
            kind: lspCompletionKindName(entry.kind),
        };
        const details = entry.labelDetails;
        if (typeof details === 'object' && details !== null) {
            const description = (details as Record<string, unknown>).description;
            if (typeof description === 'string' && description !== '') item.description = description;
        }
        out.push(item);
    }
    return out;
}

/**
 * The one merge rule for the LSP and nREPL completion lists: on an exact label
 * clash the LSP item wins, so only the nREPL items the LSP did not offer
 * (runtime names) are returned, in nREPL order. It is client-side because only
 * the client holds both answers; the nREPL cannot know what the LSP resolved.
 */
export function runtimeOnlyItems(
    lspLabels: Iterable<string | { label: string }>,
    nreplItems: NreplCompletionItem[],
): NreplCompletionItem[] {
    const seen = new Set<string>();
    for (const l of lspLabels) seen.add(typeof l === 'string' ? l : l.label);
    return nreplItems.filter((item) => !seen.has(item.label));
}

/** An LSP completion result as the client middleware sees it: an array, a list, or nothing. */
export type LspCompletionResult<T> = T[] | { items: T[]; isIncomplete?: boolean } | null | undefined;

/**
 * Merge the LSP outcome with the nREPL items: LSP items first, then the
 * runtime-only nREPL items converted by `convert`. The LSP result keeps its
 * shape (array stays array, list keeps isIncomplete). With nothing to add the
 * LSP result is returned as is; a rejected LSP with nothing to add rethrows.
 */
export function mergeCompletionResults<T extends { label: string | { label: string } }>(
    lsp: PromiseSettledResult<LspCompletionResult<T>>,
    nreplItems: NreplCompletionItem[],
    convert: (item: NreplCompletionItem) => T,
): LspCompletionResult<T> {
    const lspResult = lsp.status === 'fulfilled' ? lsp.value : undefined;
    const lspItems = Array.isArray(lspResult) ? lspResult : lspResult?.items ?? [];
    const extra = runtimeOnlyItems(lspItems.map((i) => i.label), nreplItems).map(convert);
    if (extra.length === 0) {
        if (lsp.status === 'rejected') throw lsp.reason;
        return lspResult;
    }
    const merged = [...lspItems, ...extra];
    if (lspResult && !Array.isArray(lspResult)) {
        return { items: merged, isIncomplete: lspResult.isIncomplete };
    }
    return merged;
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

interface PendingRequest {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any -- untyped nREPL payload
    resolve: (value: any) => void;
    reject: (err: unknown) => void;
    timer: NodeJS.Timeout | undefined;
    /** An `op: "error"` reply resolves instead of rejecting: it is the result of user code. */
    errorIsResult: boolean;
    /** The id-echo probe: a reply that matches no request settles it as "not echoed". */
    probe: boolean;
}

/** Bound on a tooling request (complete, macroexpand, ...) so a stuck server cannot hang it. */
export const NREPL_TOOLING_TIMEOUT_MS = 10000;

// eval and load-file run user code for as long as it takes, so they settle only on a
// reply or on socket close; a timeout would drop the real result when it arrives.
const USER_CODE_OPS: ReadonlySet<unknown> = new Set(['eval', 'load-file']);

/** The default timeout for a request: none (undefined) for user-code ops, else the tooling bound. */
export function nreplTimeoutFor(msg: Record<string, unknown>): number | undefined {
    return USER_CODE_OPS.has(msg.op) ? undefined : NREPL_TOOLING_TIMEOUT_MS;
}

// One reader per socket: every 'data' listener sees every chunk, so a socket has exactly
// one reader, which routes each reply to its request. Replies are matched by `id` only,
// because the lpy server echoes a truthy request id (hence ids start at 1) and sends one
// reply per request line. An id-less reply answers a line that never parsed, which is
// none of our pending requests, so it settles nothing. A server that never replies (an
// older one, for an unknown op) is bounded by the per-request timeout. An `op: "error"`
// reply carrying a request's id rejects that request, except for eval and load-file, where it
// resolves like any other reply so the caller shows its `error` as eval output.

/** Rejection reason of a request that got no reply within its timeout. */
export class NreplRequestTimeout extends Error {
    constructor(timeoutMs: number) {
        super(`nREPL request timed out after ${timeoutMs} ms`);
        this.name = 'NreplRequestTimeout';
    }
}

/** Rejection reason of a request answered by an `op: "error"` reply carrying its id. */
export class NreplOpError extends Error {
    constructor(message: string) {
        super(message);
        this.name = 'NreplOpError';
    }
}

/** Rejection reason of the id-echo probe when the server answered without echoing its id. */
export class NreplIdNotEchoed extends Error {
    constructor() {
        super('nREPL server did not echo the request id');
        this.name = 'NreplIdNotEchoed';
    }
}

class NreplReader {
    private buf = '';
    private nextId = 1;
    private readonly pending = new Map<number, PendingRequest>();

    constructor(private readonly sock: net.Socket) {
        // One StringDecoder per socket, not per chunk: a multi-byte character split across
        // reads (or across requests) is carried over instead of decoding to U+FFFD.
        if (sock.readableEncoding !== 'utf8') sock.setEncoding('utf8');
        sock.on('data', this.onData);
        sock.on('error', this.onError);
        sock.on('close', this.onClose);
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any -- untyped nREPL payload
    request(msg: Record<string, unknown>, timeoutMs: number | undefined, probe = false): Promise<any> {
        return new Promise((resolve, reject) => {
            const id = this.nextId++;
            // Stringify and write before registering: a throw (BigInt, cycle, write) must not
            // leave a pending entry or a live timer behind.
            const line = JSON.stringify({ ...msg, id }) + '\n';
            this.sock.write(line);
            const timer = timeoutMs === undefined ? undefined : setTimeout(() => {
                this.pending.delete(id);
                reject(new NreplRequestTimeout(timeoutMs));
            }, timeoutMs);
            timer?.unref();
            this.pending.set(id, { resolve, reject, timer, errorIsResult: USER_CODE_OPS.has(msg.op), probe });
        });
    }

    private readonly onData = (data: string) => {
        this.buf += data;
        let newlineIdx: number;
        while ((newlineIdx = this.buf.indexOf('\n')) >= 0) {
            const line = this.buf.slice(0, newlineIdx);
            this.buf = this.buf.slice(newlineIdx + 1);
            this.handleLine(line);
        }
    };

    private handleLine(line: string) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any -- untyped nREPL payload
        let reply: any;
        try {
            reply = JSON.parse(line);
        } catch {
            return; // not ours to settle: every request settles by its own id
        }
        const isObject = reply !== null && typeof reply === 'object' && !Array.isArray(reply);
        if (!isObject) return;
        const matched = reply.id === undefined || reply.id === null ? undefined : this.pending.get(reply.id);
        if (matched === undefined) {
            // Settles nothing real, but proves the server answers without echoing our id.
            this.rejectProbes();
            return;
        }
        const { id, ...payload } = reply;
        const req = matched;
        this.pending.delete(id);
        clearTimeout(req.timer);
        if (payload.op === 'error' && !req.errorIsResult) {
            req.reject(new NreplOpError(typeof payload.error === 'string' ? payload.error : 'nREPL error'));
        } else {
            req.resolve(payload.op === 'error' && typeof payload.error !== 'string'
                ? { ...payload, error: 'nREPL error' }
                : payload);
        }
    }

    private rejectProbes() {
        for (const [id, r] of [...this.pending]) {
            if (!r.probe) continue;
            this.pending.delete(id);
            clearTimeout(r.timer);
            r.reject(new NreplIdNotEchoed());
        }
    }

    private rejectAll(err: Error) {
        const reqs = [...this.pending.values()];
        this.pending.clear();
        for (const r of reqs) {
            clearTimeout(r.timer);
            r.reject(err);
        }
    }

    private readonly onError = (err: Error) => {
        this.rejectAll(new Error(`nREPL connection error: ${err.message}`));
    };

    // A closed socket never reads again, so the reader detaches and is forgotten.
    private readonly onClose = () => {
        this.rejectAll(new Error('nREPL connection closed'));
        this.sock.removeListener('data', this.onData);
        this.sock.removeListener('error', this.onError);
        this.sock.removeListener('close', this.onClose);
        readers.delete(this.sock);
    };
}

const readers = new WeakMap<net.Socket, NreplReader>();

// Other side of the contract: lispython src/lispy/tools.lpy (nrepl-eval) raises ValueError past these.
export const NREPL_MAX_EVAL_LINE = 1_000_000;
export const NREPL_MAX_EVAL_COLUMN = 100_000;

type NreplDocId = { uri: { fsPath: string; scheme: string }; isUntitled: boolean };

/**
 * The file path the nREPL server knows `doc` by, or undefined when it has none (untitled or
 * non-`file` scheme). The server keys each file's scope by its normalised path, so eval, complete,
 * macroexpand and load-file must all take the path from here: another spelling would split one
 * file into two scopes. Contract: lispython src/lispy/tools.lpy `_nrepl-dispatch`.
 */
export function nreplFileOf(doc: NreplDocId): string | undefined {
    return doc.isUntitled || doc.uri.scheme !== 'file' ? undefined : doc.uri.fsPath;
}

/**
 * `complete` request answered from `doc`'s file scope on the server (so its own definitions
 * complete); no `file` key for a document without one, which answers from the session scope.
 */
export function completeMessage(prefix: string, doc: NreplDocId): Record<string, unknown> {
    const file = nreplFileOf(doc);
    return file === undefined ? { op: 'complete', prefix } : { op: 'complete', prefix, file };
}

/**
 * `macroexpand` request expanded in `doc`'s file scope on the server (so macros defined in the
 * file expand); no `file` key for a document without one. Same path rule as eval.
 */
export function macroexpandMessage(code: string, doc: NreplDocId): Record<string, unknown> {
    const file = nreplFileOf(doc);
    return file === undefined ? { op: 'macroexpand', code } : { op: 'macroexpand', code, file };
}

/**
 * The nREPL `eval` request for the text of `doc` at `range` (`code` is exactly that text).
 * `file`, `line` and `column` (1-based, the position of the code's first character, `range[0]`)
 * let the server report errors at the editor position instead of `<nrepl-N>`.
 * An untitled or non-file document has no real path, so it sends neither file nor position and the
 * server names it `<nrepl-N>`. A position past the server's caps is dropped together with the file
 * (a file without a position would overlay the code onto lines 1..n of the real file) so the eval
 * still runs. The column counts code points like the server's Python, not UTF-16 units.
 */
export function evalMessage(
    code: string,
    doc: NreplDocId & {
        positionAt(offset: number): { line: number; character: number };
        lineAt(line: number): { text: string };
    },
    range: [number, number],
): Record<string, unknown> {
    const file = nreplFileOf(doc);
    if (file === undefined) return { op: 'eval', code };
    const start = doc.positionAt(range[0]);
    const line = start.line + 1;
    const column = [...doc.lineAt(start.line).text.slice(0, start.character)].length + 1;
    if (line > NREPL_MAX_EVAL_LINE || column > NREPL_MAX_EVAL_COLUMN) return { op: 'eval', code };
    return { op: 'eval', code, file, line, column };
}

/**
 * The request field naming the scope an eval/complete/macroexpand runs in, without making it
 * the eval's filename. Other side of the contract: lispython src/lispy/tools.lpy
 * `_nrepl-request-scope`, which reads this field and rejects a value that is not an absolute
 * path; `file` wins when both are sent.
 */
export const NREPL_NS_FIELD = 'ns';

/** The REPL panel's scope: an absolute file path (as nreplFileOf spells it), or undefined for the session scope. */
export type ReplScope = string | undefined;

export const MAIN_SCOPE_LABEL = '__main__';

// The panel never sends file/line/column: its code is not text of the file, so the eval must
// stay `<nrepl-N>` and leave linecache alone. ns picks the scope only.
function withScope(msg: Record<string, unknown>, scope: ReplScope): Record<string, unknown> {
    return scope === undefined ? msg : { ...msg, [NREPL_NS_FIELD]: scope };
}

/** `eval` request for code typed into the REPL panel, run in `scope`. */
export function panelEvalMessage(code: string, scope: ReplScope): Record<string, unknown> {
    return withScope({ op: 'eval', code }, scope);
}

/** `complete` request for the REPL panel input, answered from `scope`. */
export function panelCompleteMessage(prefix: string, scope: ReplScope): Record<string, unknown> {
    return withScope({ op: 'complete', prefix }, scope);
}

/** The deepest workspace root containing `file`, if any. */
function containingRoot(file: string, workspaceRoots: readonly string[]): string | undefined {
    let best: string | undefined;
    for (const root of workspaceRoots) {
        if (isInside(path.resolve(root), file) && (best === undefined || root.length > best.length)) best = root;
    }
    return best;
}

/** How the panel names `scope`: `__main__`, a path relative to its workspace folder, else the absolute path. */
export function replScopeLabel(scope: ReplScope, workspaceRoots: readonly string[]): string {
    if (scope === undefined) return MAIN_SCOPE_LABEL;
    const root = containingRoot(scope, workspaceRoots);
    return root === undefined ? scope : path.relative(path.resolve(root), scope);
}

/** What the panel header and its entry prompts show for `scope`. */
export function replScopeView(scope: ReplScope, workspaceRoots: readonly string[]): ReplScopeViewData {
    return scope === undefined
        ? { label: MAIN_SCOPE_LABEL, title: `${MAIN_SCOPE_LABEL} (session scope)`, prompt: MAIN_SCOPE_LABEL }
        : { label: replScopeLabel(scope, workspaceRoots), title: scope, prompt: path.basename(scope) };
}

export interface ReplScopeViewData {
    /** Header text after "Scope: ". */
    label: string;
    /** Header tooltip: the full path. */
    title: string;
    /** Prompt of panel-originated entries, shown as `<prompt>> `. */
    prompt: string;
}

export interface ReplScopeChoice {
    label: string;
    /** "session scope", or which of open / in nREPL the file is. */
    description: string;
    scope: ReplScope;
    current: boolean;
}

/**
 * Quick-pick items for Set REPL Scope: `__main__` first, then every file that is open or that
 * the server has a scope for, once each, ordered by label then path. The current scope is
 * always listed, even when it is neither open nor in the nREPL any more.
 */
export function replScopeChoices(
    openFiles: Iterable<string>,
    scopedFiles: Iterable<string>,
    current: ReplScope,
    workspaceRoots: readonly string[],
): ReplScopeChoice[] {
    const open = new Set(openFiles);
    const scoped = new Set(scopedFiles);
    const files = new Set([...open, ...scoped]);
    if (current !== undefined) files.add(current);
    const fileChoices = [...files].map((file): ReplScopeChoice => ({
        label: replScopeLabel(file, workspaceRoots),
        description: [open.has(file) ? 'open' : '', scoped.has(file) ? 'in nREPL' : ''].filter(Boolean).join(', '),
        scope: file,
        current: file === current,
    }));
    const cmp = (a: string, b: string) => (a < b ? -1 : a > b ? 1 : 0);
    fileChoices.sort((a, b) => cmp(a.label, b.label) || cmp(a.scope as string, b.scope as string));
    return [
        { label: MAIN_SCOPE_LABEL, description: 'session scope', scope: undefined, current: current === undefined },
        ...fileChoices,
    ];
}

/**
 * The panel's scope, and the files the connected server has a scope for as far as this
 * extension has seen: every request naming a file (eval/complete/macroexpand `file`, panel
 * `ns`, load-file `path`) creates that file's scope on the server, even when it then fails.
 */
export class ReplScopeState {
    current: ReplScope = undefined;
    private readonly scopedFiles = new Set<string>();

    /** Files the server has a scope for, as far as this extension has seen. */
    get scoped(): ReadonlySet<string> {
        return this.scopedFiles;
    }

    /** Record the scope `msg` created; call once the server has answered it. */
    noteAnswered(msg: Record<string, unknown>): void {
        const field = msg.op === 'load-file' ? 'path' : (typeof msg.file === 'string' ? 'file' : NREPL_NS_FIELD);
        const file = msg[field];
        if (typeof file === 'string') this.scopedFiles.add(file);
    }

    /**
     * Forget the scoped files on every new connection: the new socket may be a different
     * server (Connect via .nrepl-port can attach to any), and the list must never claim a scope
     * the server lacks. Under-reporting is harmless: `ns` creates a scope on demand, which is
     * also why the panel's current scope is kept.
     */
    connectionReset(): void {
        this.scopedFiles.clear();
    }
}

/**
 * Send one newline-delimited JSON request on `sock` and resolve with its reply (the
 * echoed `id` stripped). Requests may overlap: the socket's single reader routes each
 * reply to its request. Rejects if the socket errors or closes first (e.g. the server
 * died or the connection was swapped), when the server answers with an `op: "error"`
 * reply for it (eval and load-file resolve with it instead), or after `timeoutMs` without a reply. The default timeout
 * (nreplTimeoutFor) applies to tooling ops only; eval and load-file wait for their reply.
 */
export function nreplRequestOn(
    sock: net.Socket,
    msg: Record<string, unknown>,
    timeoutMs: number | undefined = nreplTimeoutFor(msg),
    probe = false,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any -- untyped nREPL payload
): Promise<any> {
    if (sock.destroyed) return Promise.reject(new Error('nREPL not connected'));
    let reader = readers.get(sock);
    if (reader === undefined) {
        reader = new NreplReader(sock);
        readers.set(sock, reader);
    }
    return reader.request(msg, timeoutMs, probe);
}

export const NREPL_ID_ECHO_TIMEOUT_MS = 3000;

export const NREPL_NO_ID_ECHO_MESSAGE =
    'The nREPL server does not echo request ids: it is too old. Upgrade lispython.';

export type IdEchoProbe = 'echoed' | 'missing' | 'inconclusive';

/**
 * Probe a fresh connection with one cheap request. An old server still replies, just
 * without the id, so 'missing' needs that positive evidence: a reply that carries no id or
 * one we never sent. A timeout is 'inconclusive' (a busy server is not an old one), as is
 * any other failure such as a closed socket; an error reply that carries the id is 'echoed'.
 */
export async function probeIdEcho(
    sock: net.Socket,
    timeoutMs = NREPL_ID_ECHO_TIMEOUT_MS,
): Promise<IdEchoProbe> {
    try {
        await nreplRequestOn(sock, { op: 'complete', prefix: '' }, timeoutMs, true);
    } catch (err) {
        if (err instanceof NreplIdNotEchoed) return 'missing';
        return err instanceof NreplOpError ? 'echoed' : 'inconclusive';
    }
    return 'echoed';
}

export const NREPL_ID_ECHO_INCONCLUSIVE_MESSAGE =
    'nREPL id-echo check got no reply in time; keeping the connection.';

/**
 * Probe `sock` and act on the verdict. 'missing' reports the server as too old and drops the
 * probed socket, whose replies can never be matched, so pending and later requests fail at once.
 * 'inconclusive' only logs: a busy server is not an old one. A verdict for a socket that is no
 * longer `conn`'s current one says nothing about the server now connected.
 */
export async function checkServerIdEcho(
    conn: { readonly socket: net.Socket | undefined },
    sock: net.Socket,
    notify: { log(text: string): void; tooOld(text: string): void },
    timeoutMs?: number,
): Promise<IdEchoProbe> {
    const verdict = await probeIdEcho(sock, timeoutMs);
    if (conn.socket !== sock) return verdict;
    if (verdict === 'missing') {
        notify.log(NREPL_NO_ID_ECHO_MESSAGE);
        notify.tooOld(NREPL_NO_ID_ECHO_MESSAGE);
        // Not conn.disconnect(): that would also cancel a newer connect still in flight. The
        // socket's own close clears it as current and reports the status.
        sock.destroy();
    } else if (verdict === 'inconclusive') {
        notify.log(NREPL_ID_ECHO_INCONCLUSIVE_MESSAGE);
    }
    return verdict;
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
// Function.prototype.toString (see injectWebviewFunction), so they must stay
// self-contained: an injected function may reference only other injected functions,
// by name. The shipped `main` is a minified esbuild bundle, which renames them; the
// injection re-binds each under its source name. The webview script tests run against
// both the tsc output and a production-minified bundle of this module.

/**
 * Source text declaring `fn` in the webview script under `name`. A minifier renames a
 * function and every reference to it, so the text declares the minified name (which other
 * injected functions call) and binds `name` to it for the hand-written script.
 */
// eslint-disable-next-line @typescript-eslint/no-unsafe-function-type -- any function's source text
export function injectWebviewFunction(fn: Function, name: string): string {
    const src = fn.toString();
    const declared = /^function\s+([\w$]+)\s*\(/.exec(src)?.[1];
    if (declared === undefined) throw new Error(`cannot inject ${name}: not a function declaration`);
    return declared === name ? src : `${src}\nconst ${name} = ${declared};`;
}

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
 * Start of the symbol that ends at `caret` in `text`: the walk back stops at whitespace, a
 * bracket, a quote or a comment char. The one boundary rule for editor and panel completion.
 */
export function symbolPrefixStart(text: string, caret: number): number {
    let start = caret;
    while (start > 0 && !/[\s()[\]{}'`,;"]/.test(text[start - 1])) start--;
    return start;
}

/**
 * The text before the caret that every label of a `complete` reply replaces, given the
 * reply's `word` field and the typed `prefix`. A server that labels by the segment after
 * the prefix's last `.` sends `word` (the typed part of that segment, in JS string units);
 * an older server sends no `word` and full dotted labels, which replace the whole prefix.
 * A `word` the prefix does not end with is untrusted and falls back to the whole prefix.
 * The one replace rule for editor and panel completion; injected into the webview.
 */
export function nreplReplacedText(word: unknown, prefix: string): string {
    return typeof word === 'string' && prefix.endsWith(word) ? word : prefix;
}

/** Longest common prefix of `items` ('' for none). */
export function longestCommonPrefix(items: readonly string[]): string {
    if (items.length === 0) return '';
    let prefix = items[0];
    for (const item of items) {
        let n = 0;
        while (n < prefix.length && n < item.length && prefix[n] === item[n]) n++;
        prefix = prefix.slice(0, n);
    }
    // Never end inside a surrogate pair.
    return /[\uD800-\uDBFF]$/.test(prefix) ? prefix.slice(0, -1) : prefix;
}

/**
 * VS Code notifications render `[label](command:...)` as clickable links, and its
 * linkedText parser has no escape syntax, so adjacency is broken instead: a
 * zero-width space between `]` and `(` keeps the link regex from matching.
 */
export function neutralizeNotificationLinks(text: string): string {
    return text.replace(/\]\(/g, ']\u200B(');
}

/**
 * Status-bar text for a refused openFile. The status bar renders `$(codicon)` syntax, so
 * `$(` is broken as well as notification links; the path is untrusted REPL output.
 */
export function openFileRefusedMessage(file: unknown): string {
    const safe = neutralizeNotificationLinks(String(file)).replace(/\$\(/g, '$\u200B(');
    return `Not opening "${safe}": not an existing .lpy/.py file inside the workspace`;
}

function isInside(root: string, p: string): boolean {
    const rel = path.relative(root, p);
    return rel !== '' && rel !== '..' && !rel.startsWith('..' + path.sep) && !path.isAbsolute(rel);
}

function isUncLike(p: string): boolean {
    return p.startsWith('\\\\') || p.startsWith('//');
}

// The path comes from untrusted REPL output. A lexical pre-filter runs before any fs call so
// a UNC path (SMB/NTLM connection) or a hung mount is never touched unless it already lies
// under a root. Then canonicalize (symlinks, "..") before the real containment check, and
// open the canonical path rather than the original string so the checked file is the opened
// file (short of a symlink swap racing the check).
export function resolveOpenableFile(
    file: string,
    workspaceRoots: readonly string[],
    allowedExtensions: readonly string[] = ['.lpy', '.py'],
): string | undefined {
    if (!path.isAbsolute(file)) { return undefined; }
    if (!allowedExtensions.includes(path.extname(file).toLowerCase())) { return undefined; }
    const lexical = path.resolve(file);
    // Roots are trusted, so realpath them freely: tracebacks often print the real path of a
    // root that was opened through a symlink, so the pre-filter accepts either form.
    const candidateRoots: string[] = [];
    for (const root of workspaceRoots) {
        if (isUncLike(file) && !isUncLike(root)) { continue; }
        let realRoot: string;
        try { realRoot = fs.realpathSync(root); } catch { continue; }
        if (isInside(path.resolve(root), lexical) || isInside(realRoot, lexical)) {
            candidateRoots.push(realRoot);
        }
    }
    if (candidateRoots.length === 0) { return undefined; }
    let real: string;
    try {
        real = fs.realpathSync(file);
        if (!fs.statSync(real).isFile()) { return undefined; }
    } catch {
        return undefined;
    }
    if (!allowedExtensions.includes(path.extname(real).toLowerCase())) { return undefined; }
    for (const realRoot of candidateRoots) {
        if (isInside(realRoot, real)) { return real; }
    }
    return undefined;
}

/** Warning text for a file that could not be opened; the path is untrusted REPL output. */
export function openFileWarning(file: unknown): string {
    return `Could not open file: "${neutralizeNotificationLinks(String(file))}"`;
}

/** Fresh per-render CSP nonce. */
export function makeNonce(): string {
    return crypto.randomBytes(16).toString('base64');
}

/** `value` as a JS expression safe inside an inline <script> (no `</script>` or `<!--`). */
function scriptLiteral(value: unknown): string {
    return JSON.stringify(value).replace(/</g, '\\u003c');
}

/**
 * The REPL panel: a scope header (click: Set REPL Scope), the entries, and an input line.
 * Webview → extension messages: `eval {code}`, `complete {prefix, requestId}`, `pickScope`,
 * `openFile {file, line}`. Extension → webview: `eval {code, result, prompt?}` (prompt
 * defaults to `repl`), `clear` (entries only), `scope {label, title}`,
 * `completions {requestId, items, word?}` (`word`: the reply's replaced text, see
 * nreplReplacedText; absent means the whole prefix), `focusInput`.
 */
export function getReplHtml(
    nonce: string,
    cspSource: string,
    scope: ReplScopeViewData = replScopeView(undefined, []),
): string {
    return `<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${cspSource} 'nonce-${nonce}'; script-src 'nonce-${nonce}';">
<style nonce="${nonce}">
  html, body { height: 100%; }
  body {
    font-family: var(--vscode-editor-font-family, monospace);
    font-size: var(--vscode-editor-font-size, 13px);
    margin: 0;
    padding: 0;
    display: flex;
    flex-direction: column;
    color: var(--vscode-editor-foreground);
    background: var(--vscode-editor-background);
  }
  #scope-bar {
    position: sticky; top: 0; flex: none; padding: 4px 8px;
    border-bottom: 1px solid var(--vscode-editorGroup-border);
    background: var(--vscode-editor-background);
  }
  #scope {
    background: none; border: none; padding: 0; margin: 0; font: inherit; cursor: pointer;
    color: var(--vscode-descriptionForeground); text-align: left; max-width: 100%;
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  }
  #scope:hover { color: var(--vscode-textLink-foreground); }
  #scope:focus-visible { outline: 1px solid var(--vscode-focusBorder); }
  #entries { flex: 1 1 auto; overflow-y: auto; padding: 8px; }
  #input-bar { flex: none; padding: 4px 8px 8px; border-top: 1px solid var(--vscode-editorGroup-border); }
  #input {
    display: block; width: 100%; box-sizing: border-box; resize: none; font: inherit; padding: 4px;
    color: var(--vscode-input-foreground); background: var(--vscode-input-background);
    border: 1px solid var(--vscode-input-border, transparent);
  }
  #input:focus { outline: 1px solid var(--vscode-focusBorder); outline-offset: -1px; }
  #candidates {
    color: var(--vscode-descriptionForeground); white-space: pre-wrap; padding-top: 2px;
    max-height: 6em; overflow-y: auto;
  }
  .entry { margin-bottom: 12px; border-bottom: 1px solid var(--vscode-editorGroup-border); padding-bottom: 8px; }
  .code { color: var(--vscode-textPreformat-foreground); white-space: pre-wrap; }
  .prompt { color: var(--vscode-descriptionForeground); }
  .value { color: var(--vscode-debugTokenExpression-number); white-space: pre-wrap; }
  .stdout { color: var(--vscode-terminal-ansiGreen); white-space: pre-wrap; }
  .error { color: var(--vscode-errorForeground); white-space: pre-wrap; }
  .file-link { color: var(--vscode-textLink-foreground); cursor: pointer; text-decoration: underline; }
  .file-link:hover { color: var(--vscode-textLink-activeForeground); }
</style>
</head>
<body>
<div id="scope-bar"><button id="scope" type="button"></button></div>
<div id="entries"></div>
<div id="input-bar">
<textarea id="input" rows="1" spellcheck="false" aria-label="Evaluate in the REPL scope"></textarea>
<div id="candidates" hidden></div>
</div>
<script nonce="${nonce}">
  const vscode = acquireVsCodeApi();
  const entries = document.getElementById('entries');
  const scopeButton = document.getElementById('scope');
  const input = document.getElementById('input');
  const candidatesLine = document.getElementById('candidates');
  const MAX_HISTORY = 100;
  const MAX_CANDIDATES_SHOWN = 50;
  const inputHistory = [];
  // inputHistory.length while editing a fresh line; historyDraft keeps that line.
  let historyIndex = 0;
  let historyDraft = '';
  let completionSeq = 0;
  let pendingCompletion = null;

  function renderScope(view) {
    scopeButton.textContent = 'Scope: ' + String(view.label);
    scopeButton.title = String(view.title);
  }

  function fitInput() {
    input.style.height = 'auto';
    input.style.height = Math.min(input.scrollHeight + 2, 200) + 'px';
  }

  function setInput(text, caret) {
    input.value = text;
    input.setSelectionRange(caret, caret);
    fitInput();
  }

  function showCandidates(items) {
    const shown = items.slice(0, MAX_CANDIDATES_SHOWN).join('  ');
    candidatesLine.textContent = items.length > MAX_CANDIDATES_SHOWN ? shown + '  …' : shown;
    candidatesLine.hidden = items.length === 0;
  }

  function submitInput() {
    const code = input.value;
    if (code.trim() === '') return;
    vscode.postMessage({ type: 'eval', code: code });
    if (inputHistory[inputHistory.length - 1] !== code) inputHistory.push(code);
    if (inputHistory.length > MAX_HISTORY) inputHistory.shift();
    historyIndex = inputHistory.length;
    historyDraft = '';
    pendingCompletion = null;
    showCandidates([]);
    setInput('', 0);
  }

  function recallHistory(step) {
    const next = historyIndex + step;
    if (next < 0 || next > inputHistory.length) return false;
    if (historyIndex === inputHistory.length) historyDraft = input.value;
    historyIndex = next;
    const text = next === inputHistory.length ? historyDraft : inputHistory[next];
    setInput(text, text.length);
    return true;
  }

  function requestCompletion() {
    const caret = input.selectionStart;
    if (caret !== input.selectionEnd) return false;
    const start = symbolPrefixStart(input.value, caret);
    const prefix = input.value.slice(start, caret);
    if (prefix === '') return false;
    completionSeq += 1;
    pendingCompletion = { requestId: completionSeq, caret: caret, value: input.value, prefix: prefix };
    vscode.postMessage({ type: 'complete', prefix: prefix, requestId: completionSeq });
    return true;
  }

  function applyCompletions(requestId, items, word) {
    const pending = pendingCompletion;
    if (pending === null || pending.requestId !== requestId) return;
    pendingCompletion = null;
    // Edited since the request: the answer is for text that is gone.
    if (input.value !== pending.value || input.selectionStart !== pending.caret) return;
    const names = Array.isArray(items) ? items.filter((s) => typeof s === 'string') : [];
    const replaced = nreplReplacedText(word, pending.prefix);
    const start = pending.caret - replaced.length;
    const replaceWith = (text) => setInput(
      pending.value.slice(0, start) + text + pending.value.slice(pending.caret),
      start + text.length);
    if (names.length === 1) {
      replaceWith(names[0]);
      showCandidates([]);
      return;
    }
    const common = longestCommonPrefix(names);
    if (common.length > replaced.length && common.startsWith(replaced)) replaceWith(common);
    showCandidates(names);
  }

  input.addEventListener('keydown', (e) => {
    if (e.isComposing) return;
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      submitInput();
    } else if (e.key === 'Tab' && !e.shiftKey) {
      if (requestCompletion()) e.preventDefault();
    } else if (e.key === 'ArrowUp' && !input.value.slice(0, input.selectionStart).includes('\\n')) {
      if (recallHistory(-1)) e.preventDefault();
    } else if (e.key === 'ArrowDown' && !input.value.slice(input.selectionEnd).includes('\\n')) {
      if (recallHistory(1)) e.preventDefault();
    }
  });

  input.addEventListener('input', () => {
    pendingCompletion = null;
    showCandidates([]);
    fitInput();
  });

  scopeButton.addEventListener('click', () => vscode.postMessage({ type: 'pickScope' }));

  window.addEventListener('focus', () => input.focus());

  window.addEventListener('message', (e) => {
    const msg = e.data;
    if (msg.type === 'clear') { entries.innerHTML = ''; return; }
    if (msg.type === 'scope') { renderScope(msg); return; }
    if (msg.type === 'focusInput') { input.focus(); return; }
    if (msg.type === 'completions') { applyCompletions(msg.requestId, msg.items, msg.word); return; }
    if (msg.type !== 'eval') return;
    const { code, result } = msg;
    const prompt = typeof msg.prompt === 'string' ? msg.prompt : 'repl';
    const div = document.createElement('div');
    div.className = 'entry';

    let html = '<span class="prompt">' + escapeHtml(prompt) + '&gt; </span><span class="code">' + escapeHtml(code) + '</span>';

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

  ${injectWebviewFunction(escapeHtml, 'escapeHtml')}

  ${injectWebviewFunction(linkifyErrors, 'linkifyErrors')}

  ${injectWebviewFunction(symbolPrefixStart, 'symbolPrefixStart')}

  ${injectWebviewFunction(longestCommonPrefix, 'longestCommonPrefix')}

  ${injectWebviewFunction(nreplReplacedText, 'nreplReplacedText')}

  document.addEventListener('click', (e) => {
    const link = e.target.closest('.file-link');
    if (link) {
      vscode.postMessage({ type: 'openFile', file: link.dataset.file, line: parseInt(link.dataset.line, 10) });
      return;
    }
    // A click in the entries focuses the input, unless it ended a text selection there.
    if (entries.contains(e.target) && String(window.getSelection()) === '') input.focus();
  });

  renderScope(${scriptLiteral({ label: scope.label, title: scope.title })});
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
        // Match complete lines only: a pipe read can end mid-banner, and a partial
        // "started on 127.0.0.1:12" would resolve the wrong port. The real server
        // prints a newline-terminated line, so an unterminated tail is never matched.
        // Once the port is found, parsing stops so the buffer cannot grow with
        // newline-less output from a long-running server.
        let pending = '';
        let found = false;
        proc.stdout?.on('data', (text: string) => {
            handlers.onStdout(text);
            if (found) return;
            const lines = (pending + text).split('\n');
            pending = lines.pop() ?? '';
            for (const line of lines) {
                const match = line.match(/started on .+:(\d+)/);
                if (match) {
                    found = true;
                    pending = '';
                    clearTimeout(timeout);
                    resolve(parseInt(match[1], 10));
                    return;
                }
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
                try { proc.kill('SIGKILL'); } catch { /* process already exited */ }
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
