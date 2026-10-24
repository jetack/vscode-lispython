// Pure nREPL client helpers (no `vscode` runtime import) so they run under
// `node --test` without an extension host.
import type { CompletionItemKind } from 'vscode';
import * as fs from 'fs';
import * as path from 'path';

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

/** Look for `.nrepl-port` in `startDir` and each parent directory. */
export function findNreplPort(startDir: string): number | undefined {
    let dir = path.resolve(startDir);
    for (;;) {
        try {
            const port = parsePortFile(fs.readFileSync(path.join(dir, NREPL_PORT_FILE), 'utf8'));
            if (port !== undefined) return port;
        } catch { /* absent or unreadable: keep walking up */ }
        const parent = path.dirname(dir);
        if (parent === dir) return undefined;
        dir = parent;
    }
}
