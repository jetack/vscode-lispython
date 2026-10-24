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
        let text: string | undefined;
        try {
            text = fs.readFileSync(path.join(dir, NREPL_PORT_FILE), 'utf8');
        } catch { /* absent or unreadable: keep walking up */ }
        // The nearest file wins; a malformed one must not fall through to an unrelated server above.
        if (text !== undefined) return parsePortFile(text);
        const parent = path.dirname(dir);
        if (parent === dir) return undefined;
        dir = parent;
    }
}

// escapeHtml and linkifyErrors are injected into the REPL webview script via
// Function.prototype.toString (see getReplHtml in extension.ts), so they must stay
// self-contained: linkifyErrors may reference only escapeHtml, by name.

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
