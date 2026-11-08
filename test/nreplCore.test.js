const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const net = require('net');
const core = require('../out-test/nreplCore');

test('parsePortFile', () => {
    assert.strictEqual(core.parsePortFile('5555'), 5555);
    assert.strictEqual(core.parsePortFile('5555\n'), 5555);
    for (const bad of ['', 'abc', '12x', '0', '70000', '-1']) {
        assert.strictEqual(core.parsePortFile(bad), undefined, bad);
    }
});

test('findNreplPort walks upward', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'nrepl-'));
    const sub = path.join(root, 'a', 'b');
    fs.mkdirSync(sub, { recursive: true });
    fs.writeFileSync(path.join(root, '.nrepl-port'), '4321');
    assert.strictEqual(core.findNreplPort(sub), 4321);
    fs.writeFileSync(path.join(root, 'a', '.nrepl-port'), 'junk');
    assert.strictEqual(core.findNreplPort(sub), undefined);
    fs.writeFileSync(path.join(root, 'a', '.nrepl-port'), '99');
    assert.strictEqual(core.findNreplPort(sub), 99);
    fs.rmSync(root, { recursive: true });
});

const unescapeHtml = (s) => s
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'").replace(/&amp;/g, '&');

test('escapeHtml covers & < > " \'', () => {
    assert.strictEqual(core.escapeHtml(`&<>"'`), '&amp;&lt;&gt;&quot;&#39;');
    assert.strictEqual(core.escapeHtml('&lt;'), '&amp;lt;');
});

test('linkifyErrors escapes path and text, keeps one intact link', () => {
    const file = `/tmp/a"b<c>&d'e.py`;
    const err = `Traceback:\n  File "${file}", line 12, in <module>\nValueError: "x" < 1 & y > 2`;
    const html = core.linkifyErrors(err);

    const spans = html.match(/<span class="file-link" data-file="([^"]*)" data-line="(\d+)">/g) || [];
    assert.strictEqual(spans.length, 1, html);
    assert.strictEqual((html.match(/<\/span>/g) || []).length, 1, html);

    const [, rawFile, line] = html.match(/data-file="([^"]*)" data-line="([^"]*)"/);
    assert.strictEqual(unescapeHtml(rawFile), file);
    assert.strictEqual(line, '12');

    // Outside the generated markup, no raw metacharacters survive and entities are well-formed.
    const text = html.replace(/<span class="file-link" data-file="[^"]*" data-line="\d+">|<\/span>/g, '');
    assert.doesNotMatch(text, /[<>"']/, text);
    assert.doesNotMatch(text, /&(?!amp;|lt;|gt;|quot;|#39;)/, text);
    assert.strictEqual(unescapeHtml(text), err);
});

test('webview-injected sources are self-contained', () => {
    // getReplHtml injects these via toString(); they must run with no module scope.
    const linkify = new Function(
        `${core.escapeHtml.toString()}\n${core.linkifyErrors.toString()}\nreturn linkifyErrors;`)();
    const err = 'File "/x/y.py", line 3';
    assert.strictEqual(linkify(err), core.linkifyErrors(err));
    assert.match(linkify(err), /data-file="\/x\/y\.py" data-line="3"/);
});

// Copied from vs/base/common/linkedText.ts (VS Code), the parser behind notification links.
const LINK_RE = /\[([^\]]+)\]\(((?:https?:\/\/|command:|file:)[^\)\s]+)(?: (["'])(.+?)(\3))?\)/i;

test('openFileWarning defuses notification command links', () => {
    const evil = '[x](command:workbench.action.terminal.sendSequence)';
    assert.ok(LINK_RE.test(evil));
    const warning = core.openFileWarning(evil);
    assert.ok(!LINK_RE.test(warning));
    assert.ok(warning.startsWith('Could not open file: "') && warning.endsWith('"'));
    assert.strictEqual(core.openFileWarning('/home/a b/x.py'), 'Could not open file: "/home/a b/x.py"');
    assert.strictEqual(core.openFileWarning(42), 'Could not open file: "42"');
});

test('openFileRefusedMessage defuses notification links and codicon syntax', () => {
    const evil = '[x](command:workbench.action.terminal.sendSequence)';
    const msg = core.openFileRefusedMessage(evil);
    assert.ok(!LINK_RE.test(msg));
    assert.ok(msg.startsWith('Not opening "'));
    assert.ok(!core.openFileRefusedMessage('/a/$(sync)/b.py').includes('$('));
    assert.ok(core.openFileRefusedMessage('/a/$(sync)/b.py').includes('$​(sync)'));
    assert.strictEqual(core.openFileRefusedMessage('/a/b.txt'),
        'Not opening "/a/b.txt": not an existing .lpy/.py file inside the workspace');
});

// Builds base/{ws/pkg/mod.lpy, ws/pkg/mod.PY, ws/secret.txt, ws2/x.lpy, outside/x.lpy,
// outside/in.lpy -> ws/pkg/mod.lpy, ws/link.lpy -> outside/x.lpy, ws/linkdir -> outside}.
function openFileFixture() {
    const base = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'lispy-open-')));
    const ws = path.join(base, 'ws');
    const ws2 = path.join(base, 'ws2');
    const outside = path.join(base, 'outside');
    for (const d of [path.join(ws, 'pkg'), ws2, outside]) fs.mkdirSync(d, { recursive: true });
    const w = (p) => { fs.writeFileSync(p, 'x'); return p; };
    const fx = { base, ws, ws2, outside };
    fx.lpy = w(path.join(ws, 'pkg', 'mod.lpy'));
    fx.py = w(path.join(ws, 'pkg', 'mod.PY'));
    fx.txt = w(path.join(ws, 'secret.txt'));
    fx.out = w(path.join(outside, 'x.lpy'));
    fx.sibling = w(path.join(ws2, 'x.lpy'));
    fs.symlinkSync(fx.lpy, path.join(outside, 'in.lpy'));
    fs.symlinkSync(fx.out, path.join(ws, 'link.lpy'));
    fs.symlinkSync(outside, path.join(ws, 'linkdir'));
    fx.done = () => fs.rmSync(base, { recursive: true });
    return fx;
}

test('resolveOpenableFile opens .lpy and .py files inside the workspace', () => {
    const fx = openFileFixture();
    try {
        assert.strictEqual(core.resolveOpenableFile(fx.lpy, [fx.ws]), fx.lpy);
        assert.strictEqual(core.resolveOpenableFile(fx.py, [fx.ws]), fx.py);
    } finally { fx.done(); }
});

test('resolveOpenableFile refuses a path outside the workspace', () => {
    const fx = openFileFixture();
    try {
        assert.strictEqual(core.resolveOpenableFile(fx.out, [fx.ws]), undefined);
    } finally { fx.done(); }
});

test('resolveOpenableFile refuses a ../ escape', () => {
    const fx = openFileFixture();
    try {
        assert.strictEqual(core.resolveOpenableFile(fx.ws + '/pkg/../../outside/x.lpy', [fx.ws]), undefined);
    } finally { fx.done(); }
});

test('resolveOpenableFile refuses symlinks that escape the workspace', () => {
    const fx = openFileFixture();
    try {
        assert.strictEqual(core.resolveOpenableFile(path.join(fx.ws, 'link.lpy'), [fx.ws]), undefined);
        assert.strictEqual(core.resolveOpenableFile(path.join(fx.ws, 'linkdir', 'x.lpy'), [fx.ws]), undefined);
    } finally { fx.done(); }
});

test('resolveOpenableFile refuses other non-openable paths', () => {
    const fx = openFileFixture();
    try {
        assert.strictEqual(core.resolveOpenableFile(fx.txt, [fx.ws]), undefined);
        assert.strictEqual(core.resolveOpenableFile('pkg/mod.lpy', [fx.ws]), undefined);
        assert.strictEqual(core.resolveOpenableFile(path.join(fx.ws, 'nope.lpy'), [fx.ws]), undefined);
        fs.mkdirSync(path.join(fx.ws, 'dir.lpy'));
        assert.strictEqual(core.resolveOpenableFile(path.join(fx.ws, 'dir.lpy'), [fx.ws]), undefined);
        assert.strictEqual(core.resolveOpenableFile(fx.sibling, [fx.ws]), undefined);
        assert.strictEqual(core.resolveOpenableFile(fx.lpy, []), undefined);
        assert.strictEqual(core.resolveOpenableFile('\\\\host\\share\\x.py', [fx.ws]), undefined);
        assert.strictEqual(core.resolveOpenableFile('//host/share/x.py', [fx.ws]), undefined);
    } finally { fx.done(); }
});

test('resolveOpenableFile admits files under a workspace root that is a symlink', () => {
    const fx = openFileFixture();
    try {
        const rootLink = path.join(fx.base, 'rootlink');
        fs.symlinkSync(fx.ws, rootLink);
        assert.strictEqual(core.resolveOpenableFile(path.join(rootLink, 'pkg', 'mod.lpy'), [rootLink]), fx.lpy);
        // A traceback prints the real path of a root that was opened through the symlink.
        assert.strictEqual(core.resolveOpenableFile(fx.lpy, [rootLink]), fx.lpy);
    } finally { fx.done(); }
});

test('resolveOpenableFile refuses a path lexically outside every root without resolving symlinks', () => {
    const fx = openFileFixture();
    try {
        // outside/in.lpy resolves into the workspace, but the clicked path is outside it.
        assert.strictEqual(core.resolveOpenableFile(path.join(fx.outside, 'in.lpy'), [fx.ws]), undefined);
    } finally { fx.done(); }
});

test('getReplHtml carries a nonce-based CSP on every script and style tag', () => {
    const html = core.getReplHtml('N0nce+/=', 'vscode-webview:');
    assert.ok(html.includes(`<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src vscode-webview: 'nonce-N0nce+/='; script-src 'nonce-N0nce+/=';">`));
    const tags = html.match(/<(script|style)\b[^>]*>/g);
    assert.strictEqual(tags.length, 2);
    for (const tag of tags) assert.ok(tag.includes('nonce="N0nce+/="'), tag);
    assert.ok(!/<script(?![^>]*nonce=)/.test(html));
    assert.notStrictEqual(core.makeNonce(), core.makeNonce());
});

// The webview script is assembled from function source text, and the shipped `main` is a
// minified bundle that renames those functions, so every webview test runs on both builds.
const coreBundlePath = process.env.LISPY_CORE_BUNDLE
    ? path.resolve(process.cwd(), process.env.LISPY_CORE_BUNDLE)
    : undefined;
const webviewBuilds = [['tsc', () => core], ['minified bundle', () => {
    if (coreBundlePath === undefined) return undefined;
    return require(coreBundlePath);
}]];

function webviewTest(name, fn) {
    for (const [build, load] of webviewBuilds) {
        test(`${name} [${build}]`, (t) => {
            const c = load();
            if (c === undefined) return t.skip('LISPY_CORE_BUNDLE not set; run `bun run test`');
            return fn(c, t);
        });
    }
}

// A minimal DOM for the REPL webview script: just the members the script touches.
function makeElement() {
    const listeners = {};
    return {
        listeners, value: '', innerHTML: '', textContent: '', title: '', hidden: false,
        style: {}, scrollHeight: 0, selectionStart: 0, selectionEnd: 0, children: [],
        addEventListener(type, fn) { (listeners[type] ??= []).push(fn); },
        setSelectionRange(a, b) { this.selectionStart = a; this.selectionEnd = b; },
        appendChild(child) { this.children.push(child); },
        contains: () => false,
        focused: false,
        focus() { this.focused = true; },
        scrollIntoView() {},
    };
}

/** Run the REPL webview script of `core` in node:vm and hand back its DOM and posted messages. */
function loadReplWebview(core, scopeView) {
    const vm = require('node:vm');
    const html = core.getReplHtml('n', 'c', scopeView);
    const script = /<script nonce="n">([\s\S]*?)<\/script>/.exec(html)[1];
    const els = {};
    const posted = [];
    const windowListeners = {};
    const ctx = {
        document: {
            getElementById: (id) => (els[id] ??= makeElement()),
            addEventListener() {},
            createElement: () => makeElement(),
        },
        window: {
            addEventListener: (type, fn) => { (windowListeners[type] ??= []).push(fn); },
            getSelection: () => '',
        },
        // JSON round-trip: objects built in the vm realm fail deepStrictEqual on their prototype.
        acquireVsCodeApi: () => ({ postMessage: (m) => posted.push(JSON.parse(JSON.stringify(m))) }),
    };
    vm.createContext(ctx);
    vm.runInContext(script, ctx);
    const input = els.input;
    return {
        ctx, els, input, posted,
        send: (data) => windowListeners.message.forEach((fn) => fn({ data })),
        type(text, caret = text.length) {
            input.value = text;
            input.selectionStart = input.selectionEnd = caret;
            (input.listeners.input ?? []).forEach((fn) => fn({}));
        },
        key(key, extra = {}) {
            const ev = { key, shiftKey: false, isComposing: false, defaultPrevented: false, ...extra };
            ev.preventDefault = () => { ev.defaultPrevented = true; };
            input.listeners.keydown.forEach((fn) => fn(ev));
            return ev;
        },
    };
}

webviewTest('injected escapeHtml/linkifyErrors work inside the webview script', (core) => {
    const vm = require('node:vm');
    const w = loadReplWebview(core);
    const src = 'File "/a.py", line 3';
    assert.strictEqual(vm.runInContext(`linkifyErrors(${JSON.stringify(src)})`, w.ctx), core.linkifyErrors(src));
    assert.strictEqual(vm.runInContext(`escapeHtml('<b>')`, w.ctx), '&lt;b&gt;');

    w.send({ type: 'eval', code: '<b>', result: { error: src } });
    const div = w.els.entries.children.at(-1);
    assert.ok(div.innerHTML.includes('&lt;b&gt;'));
    assert.ok(div.innerHTML.includes('<span class="file-link" data-file="/a.py" data-line="3">'));
    assert.ok(div.innerHTML.startsWith('<span class="prompt">repl&gt; </span>'), 'editor entries keep the repl prompt');
    w.send({ type: 'eval', code: '1', result: { value: '1' }, prompt: '<i>.lpy' });
    assert.ok(w.els.entries.children.at(-1).innerHTML.startsWith('<span class="prompt">&lt;i&gt;.lpy&gt; </span>'));
});

webviewTest('webview input: Enter posts eval, Shift+Enter and a blank input post nothing', (core) => {
    const w = loadReplWebview(core);
    w.type('(+ 1\n 2)');
    const shift = w.key('Enter', { shiftKey: true });
    assert.strictEqual(shift.defaultPrevented, false, 'Shift+Enter inserts the newline');
    assert.deepStrictEqual(w.posted, []);
    const enter = w.key('Enter');
    assert.strictEqual(enter.defaultPrevented, true);
    assert.deepStrictEqual(w.posted, [{ type: 'eval', code: '(+ 1\n 2)' }]);
    assert.strictEqual(w.input.value, '');
    w.type('  \n ');
    w.key('Enter');
    assert.strictEqual(w.posted.length, 1, 'blank input posts nothing');
});

webviewTest('webview: clear empties the entries only; scope messages set the header as text', (core) => {
    const w = loadReplWebview(core, { label: 'src/</script><b>x.lpy', title: '/w/src/</script><b>x.lpy' });
    assert.strictEqual(w.els.scope.textContent, 'Scope: src/</script><b>x.lpy', 'initial scope renders at once');
    assert.strictEqual(w.els.scope.title, '/w/src/</script><b>x.lpy');
    w.send({ type: 'eval', code: '1', result: { value: '1' } });
    w.type('(draft');
    w.send({ type: 'clear' });
    assert.strictEqual(w.els.entries.innerHTML, '');
    assert.strictEqual(w.input.value, '(draft');
    w.send({ type: 'scope', label: '/a/<img src=x>.lpy', title: '/a/<img src=x>.lpy' });
    assert.strictEqual(w.els.scope.textContent, 'Scope: /a/<img src=x>.lpy');
    assert.ok(!w.els.scope.innerHTML.includes('<img'), w.els.scope.innerHTML);
});

webviewTest('webview: Tab completes the symbol before the caret', (core) => {
    const w = loadReplWebview(core);
    w.type('(foo-ba x)', 7);
    const tab = w.key('Tab');
    assert.strictEqual(tab.defaultPrevented, true);
    const req = w.posted.at(-1);
    assert.deepStrictEqual(req, { type: 'complete', prefix: 'foo-ba', requestId: req.requestId });
    w.send({ type: 'completions', requestId: req.requestId, items: ['foo-bar'] });
    assert.strictEqual(w.input.value, '(foo-bar x)');
    assert.strictEqual(w.input.selectionStart, 8);

    // Several: extend to the common prefix and list them; a stale reply changes nothing.
    w.type('(pr');
    w.key('Tab');
    const stale = w.posted.at(-1).requestId;
    w.key('Tab');
    const fresh = w.posted.at(-1).requestId;
    w.send({ type: 'completions', requestId: stale, items: ['zzz'] });
    assert.strictEqual(w.input.value, '(pr');
    w.send({ type: 'completions', requestId: fresh, items: ['print', 'printf'] });
    assert.strictEqual(w.input.value, '(print');
    assert.strictEqual(w.els.candidates.hidden, false);
    assert.ok(w.els.candidates.textContent.includes('printf'));
});

// `complete` replies captured from the real servers, verbatim. NEW_*: lispython's
// segment-labelled reply with `word` (TASK-1144 branch); OLD_*: lispython main, which
// has no `word` and labels with full dotted names.
const NEW_HEAPQ_NL = { op: 'complete-result', word: 'nl', completions: ['heapq.nlargest'], items: [{ label: 'nlargest', kind: 3 }], id: 10 };
const NEW_HEAPQ_HEAPPU = { op: 'complete-result', word: 'heappu', completions: ['heapq.heappush', 'heapq.heappush-max', 'heapq.heappushpop', 'heapq.heappushpop-max'], items: [{ label: 'heappush', kind: 3 }, { label: 'heappush-max', kind: 3 }, { label: 'heappushpop', kind: 3 }, { label: 'heappushpop-max', kind: 3 }], id: 11 };
// Reply to `heapq.`, trimmed from the capture to its `op`, `word` and `id`.
const NEW_HEAPQ_DOT = { op: 'complete-result', word: '', id: 10 };
const OLD_HEAPQ_NL = { op: 'complete-result', completions: ['heapq.nlargest'], items: [{ label: 'heapq.nlargest', kind: 3 }], id: 10 };
const OLD_HEAPQ_HEAPPU = { op: 'complete-result', completions: ['heapq.heappush', 'heapq.heappushpop'], items: [{ label: 'heapq.heappush', kind: 3 }, { label: 'heapq.heappushpop', kind: 3 }], id: 11 };
const labelsOf = (reply) => core.nreplCompletionItems(reply).map((i) => i.label);

test('nreplReplaceStart: the new reply replaces its word, the old reply the whole prefix', () => {
    // Caret at the end of `(heapq.nl` (9): the prefix `heapq.nl` starts at 1.
    assert.strictEqual(core.nreplReplaceStart(9, 'heapq.nl', NEW_HEAPQ_NL.word), 7);
    assert.strictEqual(core.nreplReplaceStart(9, 'heapq.nl', OLD_HEAPQ_NL.word), 1);
    // An empty word is a real answer (the members of `heapq.`), not a missing one.
    assert.strictEqual(core.nreplReplaceStart(7, 'heapq.', NEW_HEAPQ_DOT.word), 7);
    // A word the typed prefix does not end with cannot be replaced: whole prefix.
    assert.strictEqual(core.nreplReplaceStart(9, 'heapq.nl', 'xy'), 1);
});

/** Tab on `typed` in the panel, answered as completeFromPanel does: labels plus the raw `word`. */
function tabWithReply(core, typed, reply) {
    const w = loadReplWebview(core);
    w.type(typed);
    w.key('Tab');
    w.send({ type: 'completions', requestId: w.posted.at(-1).requestId, items: labelsOf(reply), word: reply.word });
    return w;
}

webviewTest('webview: segment labels replace only the reply word', (core) => {
    const w = tabWithReply(core, '(heapq.nl', NEW_HEAPQ_NL);
    assert.strictEqual(w.input.value, '(heapq.nlargest');
    assert.strictEqual(w.input.selectionStart, '(heapq.nlargest'.length);
    // The common prefix of the labels (`heappush`) extends the word, not the whole prefix.
    assert.strictEqual(tabWithReply(core, '(heapq.heappu', NEW_HEAPQ_HEAPPU).input.value, '(heapq.heappush');
});

webviewTest('webview: full dotted labels without word still replace the whole prefix', (core) => {
    assert.strictEqual(tabWithReply(core, '(heapq.nl', OLD_HEAPQ_NL).input.value, '(heapq.nlargest');
    assert.strictEqual(tabWithReply(core, '(heapq.heappu', OLD_HEAPQ_HEAPPU).input.value, '(heapq.heappush');
});

webviewTest('webview: a word the prefix does not end with falls back to the whole prefix', (core) => {
    const w = loadReplWebview(core);
    w.type('(heapq.nl');
    w.key('Tab');
    w.send({ type: 'completions', requestId: w.posted.at(-1).requestId, items: ['heapq.nlargest'], word: 'zz' });
    assert.strictEqual(w.input.value, '(heapq.nlargest');
});

webviewTest('webview: a completions reply is dropped when the caret moved since its request', (core) => {
    const w = loadReplWebview(core);
    w.type('(foo-ba x)', 7);
    w.key('Tab');
    const { requestId } = w.posted.at(-1);
    // Moving the caret fires no input event, so only the reply-time check can catch it.
    w.input.setSelectionRange(1, 1);
    w.send({ type: 'completions', requestId, items: ['foo-bar'] });
    assert.strictEqual(w.input.value, '(foo-ba x)');
});

webviewTest('webview: focusInput focuses the input', (core) => {
    const w = loadReplWebview(core);
    w.send({ type: 'focusInput' });
    assert.strictEqual(w.input.focused, true);
});

test('findNreplPortFor: active file walks up its own root, else folders in order', (t) => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'nrepl-roots-'));
    t.after(() => fs.rmSync(root, { recursive: true }));
    const a = path.join(root, 'A');
    const b = path.join(root, 'B');
    fs.mkdirSync(path.join(a, 'src'), { recursive: true });
    fs.mkdirSync(path.join(b, 'src'), { recursive: true });
    fs.writeFileSync(path.join(b, '.nrepl-port'), '7001');
    const bPort = { port: 7001, file: path.join(b, '.nrepl-port') };

    assert.deepStrictEqual(core.findNreplPortFor(path.join(b, 'src', 'x.lpy'), [a, b]), bPort);
    assert.deepStrictEqual(core.findNreplPortFor(undefined, [a, b]), bPort);
    // A's file must not borrow B's server.
    assert.strictEqual(core.findNreplPortFor(path.join(a, 'src', 'x.lpy'), [a, b]), undefined);

    // Nearest file wins, even when malformed.
    fs.writeFileSync(path.join(b, 'src', '.nrepl-port'), 'junk');
    assert.strictEqual(core.findNreplPortFor(path.join(b, 'src', 'x.lpy'), [a, b]), undefined);
});

test('chooseNreplHome: owning folder, else folder 0, else the loose file', () => {
    const file = path.join('/w', 'B', 'src', 'x.lpy');
    assert.deepStrictEqual(core.chooseNreplHome(file, '/w/B', ['/w/A', '/w/B']),
        { activeFile: file, folder: '/w/B' });
    // A file outside every folder defers to the folders, matching the connect search.
    assert.deepStrictEqual(core.chooseNreplHome('/elsewhere/y.lpy', undefined, ['/w/A', '/w/B']),
        { activeFile: undefined, folder: '/w/A' });
    assert.deepStrictEqual(core.chooseNreplHome(undefined, undefined, ['/w/A', '/w/B']),
        { activeFile: undefined, folder: '/w/A' });
    assert.deepStrictEqual(core.chooseNreplHome('/loose/z.lpy', undefined, []),
        { activeFile: '/loose/z.lpy', folder: '/loose' });
    assert.deepStrictEqual(core.chooseNreplHome(undefined, undefined, []),
        { activeFile: undefined, folder: undefined });
});

test('resolveActiveFile: falls back to the last LisPython file unless a file or untitled editor has focus', () => {
    assert.strictEqual(core.resolveActiveFile(undefined, '/w/B/x.lpy'), '/w/B/x.lpy');
    assert.strictEqual(core.resolveActiveFile({ scheme: 'file', fsPath: '/w/A/y.py' }, '/w/B/x.lpy'), '/w/A/y.py');
    assert.strictEqual(core.resolveActiveFile({ scheme: 'untitled', fsPath: 'Untitled-1' }, '/w/B/x.lpy'), undefined);
    // Virtual documents (the nREPL output channel, git/diff views) keep the last LisPython file.
    assert.strictEqual(core.resolveActiveFile({ scheme: 'output', fsPath: 'extension-output-1' }, '/w/B/x.lpy'), '/w/B/x.lpy');
    assert.strictEqual(core.resolveActiveFile({ scheme: 'git', fsPath: '/w/A/y.lpy' }, '/w/B/x.lpy'), '/w/B/x.lpy');
    assert.strictEqual(core.resolveActiveFile(undefined, undefined), undefined);
});

test('.vscodeignore keeps tests out of the .vsix', () => {
    const lines = fs.readFileSync(path.join(__dirname, '..', '.vscodeignore'), 'utf8')
        .split(/\r?\n/).map((l) => l.trim());
    assert.ok(lines.includes('test/**'), lines.join('\n'));
});

// --- NreplConnection, against real localhost servers ---

// Cleanup is registered with t.after so a failing assertion cannot leave a server
// holding the event loop open.
function listen(t) {
    return new Promise((resolve) => {
        const server = net.createServer();
        const accepted = new Set();
        server.on('connection', (s) => {
            accepted.add(s);
            s.on('error', () => {});
            s.on('close', () => accepted.delete(s));
        });
        t.after(() => new Promise((r) => {
            for (const s of accepted) s.destroy();
            server.close(() => r());
        }));
        server.listen(0, '127.0.0.1', () => resolve(server));
    });
}

function connection(t, statuses) {
    const conn = new core.NreplConnection((c) => statuses.push(c));
    t.after(() => conn.disconnect());
    return conn;
}

async function deadPort() {
    const server = net.createServer();
    await new Promise((r) => server.listen(0, '127.0.0.1', r));
    const { port } = server.address();
    await new Promise((r) => server.close(r));
    return port;
}

const connections = (server) =>
    new Promise((resolve, reject) => server.getConnections((e, n) => (e ? reject(e) : resolve(n))));

async function waitFor(pred, what) {
    const deadline = Date.now() + 2000;
    while (!(await pred())) {
        if (Date.now() > deadline) assert.fail(`timed out waiting for ${what}`);
        await new Promise((r) => setTimeout(r, 10));
    }
}

// A leaked socket may not have been accepted yet when the check first runs, so let any
// in-flight handshake land before requiring the server to be empty.
async function assertNoConnections(server, what) {
    await new Promise((r) => setTimeout(r, 100));
    await waitFor(async () => (await connections(server)) === 0, what);
}

test('NreplConnection: a stale port leaves the live connection intact', async (t) => {
    const a = await listen(t);
    const statuses = [];
    const conn = connection(t, statuses);
    await conn.connect(a.address().port);
    const live = conn.socket;

    await assert.rejects(conn.connect(await deadPort()), (e) => !(e instanceof core.NreplConnectCancelled));
    assert.strictEqual(conn.socket, live);
    assert.strictEqual(live.destroyed, false);
    assert.deepStrictEqual(statuses, [true]);
    assert.strictEqual(await connections(a), 1);
});

test('NreplConnection: overlapping connects, last wins, nothing leaks', async (t) => {
    const a = await listen(t);
    const b = await listen(t);
    const statuses = [];
    const conn = connection(t, statuses);

    const first = conn.connect(a.address().port);
    const second = conn.connect(b.address().port);
    await assert.rejects(first, core.NreplConnectCancelled);
    await second;
    assert.strictEqual(conn.socket.remotePort, b.address().port);
    await assertNoConnections(a, 'server A to drop every connection');
    assert.strictEqual(await connections(b), 1);
    assert.strictEqual(statuses.at(-1), true);
    assert.ok(!statuses.includes(false), String(statuses));
});

test('NreplConnection: connecting while connected swaps in the new socket', async (t) => {
    const a = await listen(t);
    const b = await listen(t);
    const statuses = [];
    const conn = connection(t, statuses);
    await conn.connect(a.address().port);
    const old = conn.socket;

    const next = conn.connect(b.address().port);
    assert.strictEqual(old.destroyed, false, 'old socket must live until the new one connects');
    await next;
    assert.strictEqual(old.destroyed, true);
    assert.strictEqual(conn.socket.remotePort, b.address().port);
    await waitFor(async () => (await connections(a)) === 0, 'server A to see the old socket close');
    await new Promise((r) => setImmediate(r));
    assert.deepStrictEqual(statuses, [true, true]);
});

test('NreplConnection: disconnect cancels a pending connect', async (t) => {
    const a = await listen(t);
    const statuses = [];
    const conn = connection(t, statuses);
    const pending = conn.connect(a.address().port);
    conn.disconnect();
    await assert.rejects(pending, core.NreplConnectCancelled);
    assert.strictEqual(conn.socket, undefined);
    await assertNoConnections(a, 'server A to have no connections');
    assert.strictEqual(conn.socket, undefined);
    assert.deepStrictEqual(statuses, [false]);
});

test('NreplConnection: the server closing the current socket reports disconnected', async (t) => {
    const a = await listen(t);
    const statuses = [];
    const conn = connection(t, statuses);
    const accepted = new Promise((r) => a.once('connection', r));
    await conn.connect(a.address().port);
    assert.deepStrictEqual(statuses, [true]);
    (await accepted).destroy();
    await waitFor(() => conn.socket === undefined, 'current socket to clear');
    assert.deepStrictEqual(statuses, [true, false]);
});

function clientTo(t, server) {
    return new Promise((resolve) => {
        const sock = net.connect(server.address().port, '127.0.0.1', () => resolve(sock));
        sock.on('error', () => {});
        t.after(() => sock.destroy());
    });
}

const requestListeners = (sock) => ['data', 'close', 'error'].map((e) => sock.listenerCount(e));

// Fake server: hands each batch of `n` parsed request lines to `reply`, whose return value
// is written back in one write.
function batchingServer(server, n, reply) {
    server.on('connection', (s) => {
        s.setEncoding('utf8');
        let buf = '';
        let batch = [];
        s.on('data', (d) => {
            buf += d;
            let i;
            while ((i = buf.indexOf('\n')) >= 0) {
                batch.push(JSON.parse(buf.slice(0, i)));
                buf = buf.slice(i + 1);
                if (batch.length === n) {
                    s.write(reply(batch));
                    batch = [];
                }
            }
        });
    });
}

test('nreplRequestOn keeps one reader per socket across sequential requests', async (t) => {
    const a = await listen(t);
    batchingServer(a, 1, ([req]) =>
        JSON.stringify({ id: req.id, value: req.code }) + '\n');
    const sock = await clientTo(t, a);
    const before = requestListeners(sock);
    assert.deepStrictEqual(await core.nreplRequestOn(sock, { op: 'eval', code: 'a' }), { value: 'a' });
    const afterFirst = requestListeners(sock);
    assert.deepStrictEqual(afterFirst, before.map((c) => c + 1));
    for (const code of ['b', 'c', 'd']) {
        assert.deepStrictEqual(await core.nreplRequestOn(sock, { op: 'eval', code }), { value: code });
    }
    assert.deepStrictEqual(requestListeners(sock), afterFirst);
});

test('nreplRequestOn routes out-of-order replies on one socket by id', async (t) => {
    const a = await listen(t);
    batchingServer(a, 2, ([first, second]) =>
        JSON.stringify({ id: second.id, value: second.code }) + '\n' +
        JSON.stringify({ id: first.id, value: first.code }) + '\n');
    const sock = await clientTo(t, a);
    const evalReq = core.nreplRequestOn(sock, { op: 'eval', code: 'eval-result' });
    const complete = core.nreplRequestOn(sock, { op: 'complete', code: 'complete-result' });
    assert.deepStrictEqual(await Promise.all([evalReq, complete]),
        [{ value: 'eval-result' }, { value: 'complete-result' }]);
});

// In-memory socket: `sent` collects request lines, `reply(obj|string)` feeds a reply line.
function fakeSocket() {
    const sock = new (require('events').EventEmitter)();
    sock.destroyed = false;
    sock.readableEncoding = 'utf8';
    sock.setEncoding = () => {};
    sock.sent = [];
    sock.write = (line) => { sock.sent.push(JSON.parse(line)); };
    sock.reply = (r) => sock.emit('data', (typeof r === 'string' ? r : JSON.stringify(r)) + '\n');
    return sock;
}

const outcome = (p) => p.then((v) => ({ ok: v }), (e) => ({ err: e.message }));

test('nreplRequestOn: an id-less reply settles no pending request', async () => {
    const sock = fakeSocket();
    const p = outcome(core.nreplRequestOn(sock, { op: 'bogus' }, 50));
    sock.reply({ op: 'error', error: 'bad json' });
    assert.deepStrictEqual(await p, { err: 'nREPL request timed out after 50 ms' });
});

test('nreplRequestOn: unknown op then an id-less reply does not settle the next request', async () => {
    const sock = fakeSocket();
    const unknown = outcome(core.nreplRequestOn(sock, { op: 'bogus' }, 50));
    sock.reply({ op: 'error', error: 'bad json' }); // id-less: belongs to neither
    const next = outcome(core.nreplRequestOn(sock, { op: 'eval', code: '1' }, 1000));
    sock.reply({ id: sock.sent[1].id, value: '1' });
    assert.deepStrictEqual(await next, { ok: { value: '1' } });
    assert.deepStrictEqual(await unknown, { err: 'nREPL request timed out after 50 ms' });
});

test('nreplRequestOn rejects on an error reply carrying its id', async () => {
    const sock = fakeSocket();
    const p = outcome(core.nreplRequestOn(sock, { op: 'bogus' }, 1000));
    sock.reply({ id: sock.sent[0].id, op: 'error', error: 'unknown op: bogus' });
    assert.deepStrictEqual(await p, { err: 'unknown op: bogus' });
});

test('nreplRequestOn: a timeout removes only its own request and a late reply is ignored', async () => {
    const sock = fakeSocket();
    const slow = outcome(core.nreplRequestOn(sock, { op: 'eval', code: 'slow' }, 30));
    const fast = outcome(core.nreplRequestOn(sock, { op: 'eval', code: 'fast' }, 1000));
    assert.deepStrictEqual(await slow, { err: 'nREPL request timed out after 30 ms' });
    sock.reply({ id: sock.sent[0].id, value: 'late' });
    sock.reply({ id: sock.sent[1].id, value: 'fast' });
    assert.deepStrictEqual(await fast, { ok: { value: 'fast' } });
});

// Records every timer the code under test creates, and which are still live.
function trackTimers(t) {
    const live = new Set();
    const set = global.setTimeout, clear = global.clearTimeout;
    global.setTimeout = (fn, ms, ...a) => { const h = set(fn, ms, ...a); live.add(h); return h; };
    global.clearTimeout = (h) => { live.delete(h); clear(h); };
    t.after(() => { global.setTimeout = set; global.clearTimeout = clear; });
    return live;
}

test('nreplRequestOn: eval and load-file arm no timeout, other ops get the tooling bound', async (t) => {
    const live = trackTimers(t);
    const sock = fakeSocket();
    const evals = ['eval', 'load-file'].map((op) => core.nreplRequestOn(sock, { op }));
    assert.strictEqual(live.size, 0);
    const complete = core.nreplRequestOn(sock, { op: 'complete', prefix: 'x' });
    assert.strictEqual(live.size, 1);
    assert.strictEqual(core.nreplTimeoutFor({ op: 'eval' }), undefined);
    assert.strictEqual(core.nreplTimeoutFor({ op: 'load-file' }), undefined);
    assert.strictEqual(core.nreplTimeoutFor({ op: 'complete' }), core.NREPL_TOOLING_TIMEOUT_MS);
    sock.reply({ id: sock.sent[2].id, completions: [] });
    await complete;
    sock.reply({ id: sock.sent[0].id, value: 'a' });
    sock.reply({ id: sock.sent[1].id, value: 'b' });
    assert.deepStrictEqual(await Promise.all(evals), [{ value: 'a' }, { value: 'b' }]);
});

test('nreplRequestOn: a settled request leaves no live timer', async (t) => {
    const live = trackTimers(t);
    const sock = fakeSocket();
    const ok = core.nreplRequestOn(sock, { op: 'complete' }, 1000);
    const err = outcome(core.nreplRequestOn(sock, { op: 'complete' }, 1000));
    const closed = outcome(core.nreplRequestOn(sock, { op: 'complete' }, 1000));
    assert.strictEqual(live.size, 3);
    sock.reply({ id: sock.sent[0].id, value: 1 });
    await ok;
    sock.reply({ id: sock.sent[1].id, op: 'error', error: 'boom' });
    await err;
    assert.strictEqual(live.size, 1);
    sock.emit('close');
    await closed;
    assert.strictEqual(live.size, 0);
});

test('nreplRequestOn ignores a malformed reply line', async () => {
    const sock = fakeSocket();
    const p = outcome(core.nreplRequestOn(sock, { op: 'eval', code: '1' }, 1000));
    sock.reply('not json');
    sock.reply({ id: sock.sent[0].id, value: 'ok' });
    assert.deepStrictEqual(await p, { ok: { value: 'ok' } });
});

test('nreplRequestOn rejects every pending request when the socket closes', async (t) => {
    const a = await listen(t);
    let lines = 0;
    a.on('connection', (s) => s.on('data', (d) => {
        lines += d.toString().split('\n').length - 1;
        if (lines >= 2) s.destroy();
    }));
    const sock = await clientTo(t, a);
    const outcomes = await Promise.all(['1', '2'].map((code) =>
        core.nreplRequestOn(sock, { op: 'eval', code }).then(() => 'resolved', (e) => e.message)));
    assert.deepStrictEqual(outcomes, ['nREPL connection closed', 'nREPL connection closed']);
});

test('nreplRequestOn rejects when the server closes the socket mid-request', async (t) => {
    const a = await listen(t);
    a.on('connection', (s) => s.on('data', () => s.destroy()));
    const sock = await clientTo(t, a);
    const before = requestListeners(sock);
    const pending = core.nreplRequestOn(sock, { op: 'eval', code: '(sleep 10)' });
    const outcome = await Promise.race([
        pending.then(() => 'resolved', (e) => e.message),
        new Promise((r) => setTimeout(() => r('still pending'), 2000)),
    ]);
    assert.strictEqual(outcome, 'nREPL connection closed');
    assert.deepStrictEqual(requestListeners(sock), before);
});

test('nreplRequestOn reports the socket error that ended the request', async (t) => {
    const a = await listen(t);
    a.on('connection', (s) => s.on('data', () => s.resetAndDestroy()));
    const sock = await clientTo(t, a);
    const err = await core.nreplRequestOn(sock, { op: 'eval', code: '1' }).then(() => undefined, (e) => e);
    assert.strictEqual(err?.message, 'nREPL connection error: read ECONNRESET');
});

test('nreplRequestOn joins a response line split across writes', async (t) => {
    const a = await listen(t);
    a.on('connection', (s) => s.once('data', (d) => {
        const id = JSON.parse(d.toString()).id;
        s.write(`{"id": ${id}, "value": "sp`);
        setTimeout(() => s.write('lit"}\n'), 50);
    }));
    const sock = await clientTo(t, a);
    let chunks = 0;
    sock.on('data', () => chunks++);
    assert.deepStrictEqual(await core.nreplRequestOn(sock, { op: 'eval', code: '1' }), { value: 'split' });
    assert.strictEqual(chunks, 2, 'the response must really arrive in two reads');
});

test('nreplRequestOn keeps a multi-byte character split across reads', async (t) => {
    const bytes = Buffer.from('{"id": 1, "value": "한국어"}\n');
    const cut = bytes.indexOf(Buffer.from('한')) + 1;
    const first = bytes.subarray(0, cut);
    assert.ok(first[first.length - 1] >= 0x80, 'the split must land inside a character');
    const a = await listen(t);
    a.on('connection', (s) => s.once('data', () => {
        s.write(first);
        setTimeout(() => s.write(bytes.subarray(cut)), 50);
    }));
    const sock = await clientTo(t, a);
    let chunks = 0;
    sock.on('data', () => chunks++);
    assert.deepStrictEqual(await core.nreplRequestOn(sock, { op: 'eval', code: '1' }), { value: '한국어' });
    assert.strictEqual(chunks, 2, 'the response must really arrive in two reads');
});

test('nreplRequestOn: an error reply to eval or load-file resolves so it shows as output', async () => {
    const sock = fakeSocket();
    const evalReq = outcome(core.nreplRequestOn(sock, { op: 'eval', code: '(' }));
    const load = outcome(core.nreplRequestOn(sock, { op: 'load-file', path: 'x.lpy' }));
    sock.reply({ id: sock.sent[0].id, op: 'error', error: 'SyntaxError' });
    sock.reply({ id: sock.sent[1].id, op: 'error', error: 'no such file' });
    assert.deepStrictEqual(await evalReq, { ok: { op: 'error', error: 'SyntaxError' } });
    assert.deepStrictEqual(await load, { ok: { op: 'error', error: 'no such file' } });
});

test('nreplRequestOn: an eval error reply without error text gets the fallback text', async () => {
    const sock = fakeSocket();
    const p = outcome(core.nreplRequestOn(sock, { op: 'eval', code: '(' }));
    sock.reply({ id: sock.sent[0].id, op: 'error' });
    assert.deepStrictEqual(await p, { ok: { op: 'error', error: 'nREPL error' } });
});

test('probeIdEcho: echoed when the probe reply carries its id', async () => {
    const sock = fakeSocket();
    const p = core.probeIdEcho(sock, 1000);
    assert.deepStrictEqual(sock.sent.map((m) => m.op), ['complete']);
    sock.reply({ id: sock.sent[0].id, completions: [] });
    assert.strictEqual(await p, 'echoed');
});

test('probeIdEcho: an error reply that carries the id still counts as echoed', async () => {
    const sock = fakeSocket();
    const p = core.probeIdEcho(sock, 1000);
    sock.reply({ id: sock.sent[0].id, op: 'error', error: 'unknown op' });
    assert.strictEqual(await p, 'echoed');
});

test('probeIdEcho: an id-less reply is missing, at once', async () => {
    const sock = fakeSocket();
    const p = core.probeIdEcho(sock, 5000);
    sock.reply({ completions: [] });
    assert.strictEqual(await p, 'missing');
});

test('probeIdEcho: a reply with an id we never sent is missing', async () => {
    const sock = fakeSocket();
    const p = core.probeIdEcho(sock, 5000);
    sock.reply({ id: sock.sent[0].id + 100, completions: [] });
    assert.strictEqual(await p, 'missing');
});

test('probeIdEcho: no reply is inconclusive, and a closed socket too', async () => {
    const quiet = fakeSocket();
    assert.strictEqual(await core.probeIdEcho(quiet, 30), 'inconclusive');
    const closed = fakeSocket();
    const p = core.probeIdEcho(closed, 1000);
    closed.emit('close');
    assert.strictEqual(await p, 'inconclusive');
});

test('probeIdEcho: an id-less reply settles only the probe, not other requests', async () => {
    const sock = fakeSocket();
    const other = outcome(core.nreplRequestOn(sock, { op: 'bogus' }, 50));
    const probe = core.probeIdEcho(sock, 5000);
    sock.reply({ completions: [] });
    assert.strictEqual(await probe, 'missing');
    assert.deepStrictEqual(await other, { err: 'nREPL request timed out after 50 ms' });
});

// Like NreplConnection, the socket's close clears it as current.
function fakeConn(sock) {
    const conn = { socket: sock, destroys: 0 };
    sock.destroy = () => { conn.destroys++; sock.destroyed = true; sock.emit('close'); if (conn.socket === sock) conn.socket = undefined; };
    return conn;
}

function notes() {
    const n = { logs: [], toasts: [] };
    n.log = (t) => n.logs.push(t);
    n.tooOld = (t) => n.toasts.push(t);
    return n;
}

test('checkServerIdEcho: an id-less reply shows the too-old message once and disconnects', async () => {
    const sock = fakeSocket();
    const conn = fakeConn(sock);
    const n = notes();
    const p = core.checkServerIdEcho(conn, sock, n, 5000);
    sock.reply({ completions: [] });
    assert.strictEqual(await p, 'missing');
    assert.deepStrictEqual(n.toasts, [core.NREPL_NO_ID_ECHO_MESSAGE]);
    assert.deepStrictEqual(n.logs, [core.NREPL_NO_ID_ECHO_MESSAGE]);
    assert.strictEqual(conn.destroys, 1);
});

test('checkServerIdEcho: no reply keeps the connection, logs, and shows no toast', async () => {
    const sock = fakeSocket();
    const conn = fakeConn(sock);
    const n = notes();
    assert.strictEqual(await core.checkServerIdEcho(conn, sock, n, 30), 'inconclusive');
    assert.deepStrictEqual(n.toasts, []);
    assert.deepStrictEqual(n.logs, [core.NREPL_ID_ECHO_INCONCLUSIVE_MESSAGE]);
    assert.strictEqual(conn.destroys, 0);
    assert.strictEqual(conn.socket, sock);
});

test('checkServerIdEcho: an echoing server is left alone and silent', async () => {
    const sock = fakeSocket();
    const conn = fakeConn(sock);
    const n = notes();
    const p = core.checkServerIdEcho(conn, sock, n, 1000);
    sock.reply({ id: sock.sent[0].id, completions: [] });
    assert.strictEqual(await p, 'echoed');
    assert.deepStrictEqual([n.logs, n.toasts, conn.destroys], [[], [], 0]);
});

test('checkServerIdEcho: a verdict for a swapped-out socket changes nothing', async () => {
    const sock = fakeSocket();
    const other = fakeSocket();
    const conn = fakeConn(other);
    const n = notes();
    const p = core.checkServerIdEcho(conn, sock, n, 5000);
    sock.reply({ completions: [] });
    assert.strictEqual(await p, 'missing');
    assert.deepStrictEqual([n.logs, n.toasts, conn.destroys], [[], [], 0]);
});

test('startNreplServer rejects at once when the binary cannot be spawned', async () => {
    const bin = 'lpy-does-not-exist-xyz';
    const t0 = Date.now();
    const origSet = global.setTimeout;
    const origClear = global.clearTimeout;
    const startTimers = new Set();
    const cleared = new Set();
    global.setTimeout = (fn, ms, ...rest) => {
        const h = origSet(fn, ms, ...rest);
        if (ms === 10000) startTimers.add(h);
        return h;
    };
    global.clearTimeout = (h) => { cleared.add(h); return origClear(h); };
    try {
        const { port } = core.startNreplServer(bin, process.cwd(), { onStdout() {}, onStderr() {} });
        await assert.rejects(port, (err) => err.message.includes(bin) && err.message.includes('ENOENT') && err.message.includes(process.cwd()));
        assert.ok(Date.now() - t0 < 2000, 'must not wait for the start timeout');
        assert.strictEqual(startTimers.size, 1, 'the 10 s start timer must have been created');
        for (const h of startTimers) assert.ok(cleared.has(h), 'the start timer must be cleared on the error path');
    } finally {
        global.setTimeout = origSet;
        global.clearTimeout = origClear;
        for (const h of startTimers) origClear(h);
    }
});

test('startNreplServer resolves the port from the server banner', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'fakelpy-'));
    const fake = path.join(dir, 'lpy');
    fs.writeFileSync(fake, '#!/bin/sh\necho "nREPL server started on 127.0.0.1:1234"\nexec sleep 30\n', { mode: 0o755 });
    const seen = [];
    const started = core.startNreplServer(fake, process.cwd(), { onStdout: (t) => seen.push(t), onStderr() {} });
    try {
        assert.strictEqual(await started.port, 1234);
        assert.ok(seen.join('').includes('started on'));
    } finally {
        started.proc.kill();
        fs.rmSync(dir, { recursive: true });
    }
});

test('startNreplServer resolves the port from a banner split across reads', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'fakelpy-'));
    const fake = path.join(dir, 'lpy');
    fs.writeFileSync(fake,
        "#!/bin/sh\nprintf 'nREPL server started on 127.0.0.1:1'\nsleep 0.3\nprintf '234\\n'\nexec sleep 30\n",
        { mode: 0o755 });
    const started = core.startNreplServer(fake, process.cwd(), { onStdout() {}, onStderr() {} });
    try {
        assert.strictEqual(await started.port, 1234);
    } finally {
        started.proc.kill();
        fs.rmSync(dir, { recursive: true });
    }
});

// A fake `lpy` that appends its pid to `spawns` on every launch, then optionally prints
// the banner, then keeps running like a real server.
function fakeServer(banner) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'fakelpy-'));
    const bin = path.join(dir, 'lpy');
    const spawns = path.join(dir, 'spawns');
    const announce = banner ? 'sleep 0.2\necho "nREPL server started on 127.0.0.1:1234"\n' : '';
    fs.writeFileSync(bin, `#!/bin/sh\necho $$ >> "${spawns}"\n${announce}exec sleep 30\n`, { mode: 0o755 });
    const pids = () => (fs.existsSync(spawns)
        ? fs.readFileSync(spawns, 'utf8').split('\n').filter(Boolean).map(Number)
        : []);
    const cleanup = () => {
        for (const pid of pids()) { try { process.kill(pid, 'SIGKILL'); } catch {} }
        fs.rmSync(dir, { recursive: true });
    };
    return { bin, pids, cleanup };
}

function alive(pid) {
    try {
        process.kill(pid, 0);
        return true;
    } catch (err) {
        if (err.code === 'ESRCH') return false;
        throw err;
    }
}

function startOpts(bin, extra) {
    return { lpy: bin, cwd: process.cwd(), onStdout() {}, onStderr() {}, connect: async () => {}, ...extra };
}

test('NreplServerSupervisor: overlapping starts share one spawn, and stop kills it', async () => {
    const fake = fakeServer(true);
    const sup = new core.NreplServerSupervisor();
    try {
        const ports = [];
        const opts = startOpts(fake.bin, { connect: async (port) => { ports.push(port); } });
        const a = sup.start(opts);
        const b = sup.start(opts);
        await Promise.all([a, b]);
        assert.strictEqual(fake.pids().length, 1, 'a second start while one is in flight must not spawn');
        assert.deepStrictEqual(ports, [1234]);
        assert.strictEqual(sup.process.pid, fake.pids()[0]);
        await sup.stop();
        assert.strictEqual(sup.process, undefined);
        for (const pid of fake.pids()) assert.ok(!alive(pid), `server ${pid} must be dead after stop`);
    } finally {
        await sup.stop();
        fake.cleanup();
    }
});

test('NreplServerSupervisor: a start that times out kills its server', async () => {
    const fake = fakeServer(false);
    const sup = new core.NreplServerSupervisor();
    try {
        const started = sup.start(startOpts(fake.bin, { startTimeoutMs: 1000 }));
        await new Promise((r) => setImmediate(r));
        // From the supervisor, not the fake's pid file, which the shell may not have written yet.
        const pid = sup.process?.pid;
        assert.ok(pid, 'the server must have been spawned');
        await assert.rejects(started, /did not start in time/);
        assert.strictEqual(sup.process, undefined);
        assert.strictEqual(sup.starting, false);
        assert.ok(!alive(pid), 'the timed-out server must be dead');
    } finally {
        await sup.stop();
        fake.cleanup();
    }
});

test('NreplServerSupervisor: a start whose connect fails kills its server', async () => {
    const fake = fakeServer(true);
    const sup = new core.NreplServerSupervisor();
    try {
        const refused = new Error('connect refused');
        const opts = startOpts(fake.bin, { connect: async () => { throw refused; } });
        await assert.rejects(sup.start(opts), (err) => err === refused);
        assert.strictEqual(fake.pids().length, 1, 'the fake server must have been launched');
        assert.strictEqual(sup.process, undefined);
        assert.ok(!alive(fake.pids()[0]), 'the server whose connect failed must be dead');
    } finally {
        await sup.stop();
        fake.cleanup();
    }
});

test('NreplServerSupervisor: a superseded connect keeps the server', async () => {
    const fake = fakeServer(true);
    const sup = new core.NreplServerSupervisor();
    try {
        const opts = startOpts(fake.bin, { connect: async () => { throw new core.NreplConnectCancelled(); } });
        await assert.rejects(sup.start(opts), core.NreplConnectCancelled);
        assert.strictEqual(fake.pids().length, 1);
        assert.strictEqual(sup.process?.pid, fake.pids()[0], 'the server must stay tracked');
        assert.ok(alive(fake.pids()[0]), 'a superseded connect must not kill the server');
        await sup.stop();
        assert.ok(!alive(fake.pids()[0]), 'stop must still kill it');
    } finally {
        await sup.stop();
        fake.cleanup();
    }
});

test('NreplServerSupervisor: stop after the spawn cancels the start without connecting', async () => {
    const fake = fakeServer(true);
    const sup = new core.NreplServerSupervisor();
    try {
        let connects = 0;
        const started = sup.start(startOpts(fake.bin, { connect: async () => { connects++; } }));
        await new Promise((r) => setImmediate(r));
        const pid = sup.process?.pid;
        assert.ok(pid, 'the server must have been spawned');
        await sup.stop();
        await assert.rejects(started, core.NreplConnectCancelled);
        assert.strictEqual(connects, 0, 'a stopped server must not be connected to');
        assert.strictEqual(sup.process, undefined);
        assert.ok(!alive(pid), 'the stopped server must be dead');
    } finally {
        await sup.stop();
        fake.cleanup();
    }
});

test('NreplServerSupervisor: stop as the banner arrives skips the connect', async () => {
    const fake = fakeServer(true);
    const sup = new core.NreplServerSupervisor();
    try {
        let connects = 0;
        let pid;
        let stopping;
        // onStdout runs before the port resolves, so this stop lands between banner and connect.
        const onStdout = () => { pid = sup.process?.pid; stopping = sup.stop(); };
        const started = sup.start(startOpts(fake.bin, { onStdout, connect: async () => { connects++; } }));
        await assert.rejects(started, core.NreplConnectCancelled);
        assert.ok(pid, 'the banner must have been seen');
        assert.strictEqual(connects, 0, 'a stopped server must not be connected to');
        await stopping;
        assert.ok(!alive(pid), 'the stopped server must be dead');
    } finally {
        await sup.stop();
        fake.cleanup();
    }
});

test('NreplServerSupervisor: stop while a restart awaits the old server cancels the spawn', async () => {
    const fake = fakeServer(true);
    const sup = new core.NreplServerSupervisor();
    try {
        await sup.start(startOpts(fake.bin));
        const restart = sup.start(startOpts(fake.bin));
        const stopped = sup.stop();
        await assert.rejects(restart, core.NreplConnectCancelled);
        await stopped;
        assert.strictEqual(sup.process, undefined);
        assert.strictEqual(fake.pids().length, 1, 'the cancelled restart must not spawn');
        for (const pid of fake.pids()) assert.ok(!alive(pid), `server ${pid} must be dead`);
    } finally {
        await sup.stop();
        fake.cleanup();
    }
});

function fakeDoc(text, fsPath, extra = {}) {
    return {
        uri: { fsPath, scheme: 'file' },
        isUntitled: false,
        positionAt(offset) {
            const before = text.slice(0, offset).split('\n');
            return { line: before.length - 1, character: before[before.length - 1].length };
        },
        lineAt: (line) => ({ text: text.split('\n')[line] ?? '' }),
        ...extra,
    };
}

test('evalMessage: file/line/column are 1-based at the first character sent', () => {
    const text = '(a)\n  (b 1)\n(c)\n';
    const doc = fakeDoc(text, '/w/x.lpy');
    const at = (needle) => {
        const start = text.indexOf(needle);
        return core.evalMessage(needle, doc, [start, start + needle.length]);
    };
    assert.deepStrictEqual(core.evalMessage(text, doc, [0, text.length]),
        { op: 'eval', code: text, file: '/w/x.lpy', line: 1, column: 1 }, 'whole document');
    assert.deepStrictEqual(at('(b 1)'),
        { op: 'eval', code: '(b 1)', file: '/w/x.lpy', line: 2, column: 3 }, 'mid-line start');
    assert.deepStrictEqual(at('(c)'),
        { op: 'eval', code: '(c)', file: '/w/x.lpy', line: 3, column: 1 }, 'top-level form');
});

test('evalMessage: column counts code points, not UTF-16 units', () => {
    const text = '(a "😀😀") (b)';
    const start = text.indexOf('(b)');
    assert.strictEqual(start, 11, 'two emoji are 4 UTF-16 units but 2 code points: column 10, not 12');
    assert.deepStrictEqual(core.evalMessage('(b)', fakeDoc(text, '/w/e.lpy'), [start, start + 3]),
        { op: 'eval', code: '(b)', file: '/w/e.lpy', line: 1, column: 10 });
});

test('evalMessage: position past the server caps is dropped with the file', () => {
    const lastCol = core.NREPL_MAX_EVAL_COLUMN;
    const wide = 'x'.repeat(lastCol + 5);
    const doc = fakeDoc(wide, '/w/min.lpy');
    const col = (c) => core.evalMessage('x', doc, [c - 1, c]);
    assert.deepStrictEqual(col(lastCol),
        { op: 'eval', code: 'x', file: '/w/min.lpy', line: 1, column: lastCol }, 'at the column cap');
    assert.deepStrictEqual(col(lastCol + 1),
        { op: 'eval', code: 'x' }, 'past the column cap');

    const lastLine = core.NREPL_MAX_EVAL_LINE;
    const tall = fakeDoc('', '/w/tall.lpy', { positionAt: (o) => ({ line: o, character: 0 }) });
    const line = (l) => core.evalMessage('x', tall, [l - 1, l]);
    assert.deepStrictEqual(line(lastLine),
        { op: 'eval', code: 'x', file: '/w/tall.lpy', line: lastLine, column: 1 }, 'at the line cap');
    assert.deepStrictEqual(line(lastLine + 1),
        { op: 'eval', code: 'x' }, 'past the line cap');
});

test('evalMessage: untitled and non-file documents send no file or position', () => {
    assert.deepStrictEqual(
        core.evalMessage('(a)', fakeDoc('(a)', 'Untitled-1', { isUntitled: true, uri: { fsPath: 'Untitled-1', scheme: 'untitled' } }), [0, 3]),
        { op: 'eval', code: '(a)' }, 'untitled');
    assert.deepStrictEqual(
        core.evalMessage('(a)', fakeDoc('(a)', '/x.lpy', { uri: { fsPath: '/x.lpy', scheme: 'vscode-remote' } }), [0, 3]),
        { op: 'eval', code: '(a)' }, 'non-file scheme');
});

test('completeMessage/macroexpandMessage use the eval file path, and none for untitled or non-file docs', () => {
    const doc = fakeDoc('(a)', '/w/f.lpy');
    const file = core.evalMessage('(a)', doc, [0, 3]).file;
    assert.strictEqual(file, '/w/f.lpy');
    assert.deepStrictEqual(core.completeMessage('ab', doc), { op: 'complete', prefix: 'ab', file });
    assert.deepStrictEqual(core.macroexpandMessage('(a)', doc), { op: 'macroexpand', code: '(a)', file });
    const untitled = fakeDoc('(a)', 'Untitled-1', { isUntitled: true, uri: { fsPath: 'Untitled-1', scheme: 'untitled' } });
    const remote = fakeDoc('(a)', '/x.lpy', { uri: { fsPath: '/x.lpy', scheme: 'vscode-remote' } });
    for (const d of [untitled, remote]) {
        assert.deepStrictEqual(core.completeMessage('ab', d), { op: 'complete', prefix: 'ab' });
        assert.deepStrictEqual(core.macroexpandMessage('(a)', d), { op: 'macroexpand', code: '(a)' });
    }
});

test('panelEvalMessage/panelCompleteMessage: __main__ sends no ns key, a file scope sends ns only', () => {
    const wire = (m) => JSON.parse(JSON.stringify(m));
    const mainEval = core.panelEvalMessage('(a)', undefined);
    const mainComplete = core.panelCompleteMessage('ab', undefined);
    // An `ns: undefined` key is dropped by JSON but would be a null-valued field to other encoders.
    assert.ok(!('ns' in mainEval) && !('ns' in mainComplete));
    assert.deepStrictEqual(wire(mainEval), { op: 'eval', code: '(a)' });
    assert.deepStrictEqual(wire(mainComplete), { op: 'complete', prefix: 'ab' });
    // Never file/line/column: the panel's code is not file text, so the eval must stay <nrepl-N>.
    assert.deepStrictEqual(wire(core.panelEvalMessage('(a)', '/w/f.lpy')), { op: 'eval', code: '(a)', ns: '/w/f.lpy' });
    assert.deepStrictEqual(wire(core.panelCompleteMessage('ab', '/w/f.lpy')), { op: 'complete', prefix: 'ab', ns: '/w/f.lpy' });
});

test('replScopeLabel: __main__, relative inside the deepest workspace root, absolute outside', () => {
    const roots = ['/w', '/w/pkg', '/other'];
    assert.strictEqual(core.replScopeLabel(undefined, roots), '__main__');
    assert.strictEqual(core.replScopeLabel('/w/src/foo.lpy', roots), path.join('src', 'foo.lpy'));
    assert.strictEqual(core.replScopeLabel('/w/pkg/m.lpy', roots), 'm.lpy');
    assert.strictEqual(core.replScopeLabel('/wx/foo.lpy', roots), '/wx/foo.lpy', 'a sibling with a shared name prefix is outside');
    assert.strictEqual(core.replScopeLabel('/tmp/foo.lpy', []), '/tmp/foo.lpy');
});

test('replScopeChoices: __main__ first, open and in-nREPL merged, current flagged, sorted', () => {
    const roots = ['/w'];
    const view = (cs) => cs.map((c) => [c.label, c.description, c.scope, c.current]);
    assert.deepStrictEqual(
        view(core.replScopeChoices(['/w/b.lpy', '/w/a.lpy'], ['/w/b.lpy', '/x/c.lpy'], '/w/a.lpy', roots)),
        [
            ['__main__', 'session scope', undefined, false],
            ['/x/c.lpy', 'in nREPL', '/x/c.lpy', false],
            ['a.lpy', 'open', '/w/a.lpy', true],
            ['b.lpy', 'open, in nREPL', '/w/b.lpy', false],
        ]);
    // The current scope stays listed after its file is closed; __main__ is current by default.
    assert.deepStrictEqual(core.replScopeChoices([], [], '/w/gone.lpy', roots).map((c) => c.scope), [undefined, '/w/gone.lpy']);
    assert.strictEqual(core.replScopeChoices([], [], undefined, roots)[0].current, true);
});

test('ReplScopeState: answered requests mark their file scoped; a new connection clears them, not the scope', () => {
    const s = new core.ReplScopeState();
    s.noteAnswered({ op: 'eval', code: '1', file: '/w/a.lpy', line: 1, column: 1 });
    s.noteAnswered({ op: 'load-file', path: '/w/b.lpy' });
    s.noteAnswered({ op: 'complete', prefix: 'x', ns: '/w/c.lpy' });
    s.noteAnswered({ op: 'eval', code: '1' });
    assert.deepStrictEqual([...s.scoped].sort(), ['/w/a.lpy', '/w/b.lpy', '/w/c.lpy']);
    s.current = '/w/a.lpy';
    s.connectionReset();
    assert.deepStrictEqual([...s.scoped], []);
    assert.strictEqual(s.current, '/w/a.lpy');
});

test('symbolPrefixStart stops at brackets, quotes and whitespace; longestCommonPrefix keeps code points whole', () => {
    const text = '(foo [bar.baz-q "s';
    const prefixAt = (caret) => text.slice(core.symbolPrefixStart(text, caret), caret);
    assert.strictEqual(prefixAt(4), 'foo');
    assert.strictEqual(prefixAt(15), 'bar.baz-q');
    assert.strictEqual(prefixAt(18), 's');
    assert.strictEqual(prefixAt(16), '');
    assert.strictEqual(core.longestCommonPrefix(['a\u{1F600}', 'a\u{1F601}']), 'a');
});

/** lispython's tools.lpy: $LISPYTHON_DIR, else a `lispython` checkout beside the repo or any ancestor (worktrees nest). */
function findLispythonTools() {
    const rel = path.join('src', 'lispy', 'tools.lpy');
    if (process.env.LISPYTHON_DIR) return path.join(process.env.LISPYTHON_DIR, rel);
    let dir = path.resolve(__dirname, '..');
    while (path.dirname(dir) !== dir) {
        const candidate = path.join(path.dirname(dir), 'lispython', rel);
        if (fs.existsSync(candidate)) return candidate;
        dir = path.dirname(dir);
    }
    return undefined;
}

test('NREPL_NS_FIELD mirrors the field lispython reads for eval/complete/macroexpand scope', (t) => {
    const tools = findLispythonTools();
    if (tools === undefined || !fs.existsSync(tools)) return t.skip('lispython checkout not found; set LISPYTHON_DIR');
    const src = fs.readFileSync(tools, 'utf8');
    // tools.lpy spells the field once, as NREPL-NS-FIELD; its value must equal ours.
    const constant = /^\(= NREPL-NS-FIELD "([^"]*)"\)/m.exec(src);
    assert.ok(constant, '(= NREPL-NS-FIELD "...") not found in tools.lpy');
    assert.strictEqual(constant[1], core.NREPL_NS_FIELD);
    // ...and every scope read and dispatch must go through that constant, not a literal.
    const def = /\(def _nrepl-request-scope [\s\S]*?(?=\n\(def )/.exec(src);
    assert.ok(def, '_nrepl-request-scope not found');
    assert.ok(def[0].includes('(.get msg NREPL-NS-FIELD)'), def[0]);
    assert.ok(def[0].includes('(in NREPL-NS-FIELD fields)'), def[0]);
    assert.ok(!def[0].includes(`"${core.NREPL_NS_FIELD}"`), def[0]);
    for (const op of ['eval', 'complete', 'macroexpand']) {
        const at = src.indexOf(`(== op "${op}")`);
        assert.ok(at >= 0, `dispatch for ${op} not found`);
        const fields = /_nrepl-request-scope msg (\[[^\]]*\])/.exec(src.slice(at))[1];
        assert.strictEqual(fields, '["file" NREPL-NS-FIELD]', op);
    }
});

test('lspCompletionKindName maps 1-based LSP kinds to VS Code names', () => {
    const k = core.lspCompletionKindName;
    assert.strictEqual(k(1), 'Text');
    assert.strictEqual(k(3), 'Function');
    assert.strictEqual(k(6), 'Variable');
    assert.strictEqual(k(7), 'Class');
    assert.strictEqual(k(14), 'Keyword');
    assert.strictEqual(k(25), 'TypeParameter');
    for (const bad of [0, 26, -1, 3.5, NaN, '3', undefined, null]) {
        assert.strictEqual(k(bad), 'Function', String(bad));
    }
});

test('nreplCompletionItems prefers items and skips malformed entries', () => {
    const got = core.nreplCompletionItems({
        items: [
            { label: 'if', kind: 14, labelDetails: { description: 'special form' } },
            { label: 'defmacro', kind: 3, labelDetails: { description: 'macro' } },
            { label: 'print', kind: 3 },
            { label: 'x', kind: 99 },
            { label: 'empty', kind: 3, labelDetails: { description: '' } },
            { kind: 3 },
        ],
    });
    assert.deepStrictEqual(got, [
        { label: 'if', kind: 'Keyword', description: 'special form' },
        { label: 'defmacro', kind: 'Function', description: 'macro' },
        { label: 'print', kind: 'Function' },
        { label: 'x', kind: 'Function' },
        { label: 'empty', kind: 'Function' },
    ]);
});

test('lspCompletionKindName agrees with the LSP CompletionItemKind enum', () => {
    const entries = Object.entries(require('vscode-languageserver-types').CompletionItemKind);
    assert.strictEqual(entries.length, 25);
    for (const [name, value] of entries) {
        assert.strictEqual(core.lspCompletionKindName(value), name, `kind ${value}`);
    }
});

test('nreplCompletionItems reads only items', () => {
    assert.deepStrictEqual(
        core.nreplCompletionItems({ completions: ['A', 'f'], kinds: { A: 'class', f: 'macro' } }), []);
    assert.deepStrictEqual(core.nreplCompletionItems({}), []);
});

test('runtimeOnlyItems: LSP wins a clash, nREPL-only names pass through in order', () => {
    const nrepl = [
        { label: 'def', kind: 'Keyword' },
        { label: 'my_var', kind: 'Variable' },
        { label: 'print', kind: 'Function' },
        { label: 'zed', kind: 'Variable' },
    ];
    const out = core.runtimeOnlyItems(['def', { label: 'print' }], nrepl);
    assert.deepStrictEqual(out.map((i) => i.label), ['my_var', 'zed']);
    assert.deepStrictEqual(core.runtimeOnlyItems([], nrepl), nrepl);
});

const mergeNrepl = [
    { label: 'def', kind: 'Keyword' },
    { label: 'my_var', kind: 'Variable' },
];
const mergeConvert = (i) => ({ label: i.label, from: 'nrepl' });
const lspItem = (label) => ({ label, from: 'lsp' });
const settled = (value) => ({ status: 'fulfilled', value });

test('mergeCompletionResults: array result becomes an incomplete list, LSP first, clash dropped', () => {
    const out = core.mergeCompletionResults(settled([lspItem('print'), lspItem('def')]), mergeNrepl, mergeConvert);
    assert.deepStrictEqual(out, {
        items: [lspItem('print'), lspItem('def'), { label: 'my_var', from: 'nrepl' }],
        isIncomplete: true,
    });
});

test('mergeCompletionResults: a complete list turns incomplete once nREPL items join', () => {
    const out = core.mergeCompletionResults(
        settled({ items: [lspItem('print')], isIncomplete: false }), mergeNrepl, mergeConvert);
    assert.strictEqual(out.isIncomplete, true);
    assert.deepStrictEqual(out.items.map((i) => i.label), ['print', 'def', 'my_var']);
});

test('mergeCompletionResults: null LSP yields the nREPL items as an incomplete list', () => {
    const out = core.mergeCompletionResults(settled(null), mergeNrepl, mergeConvert);
    assert.deepStrictEqual(out.items.map((i) => i.label), ['def', 'my_var']);
    assert.strictEqual(out.isIncomplete, true);
});

test('mergeCompletionResults: rejected LSP falls back to nREPL, rethrows when nREPL is empty', () => {
    const rejected = { status: 'rejected', reason: new Error('lsp down') };
    const out = core.mergeCompletionResults(rejected, mergeNrepl, mergeConvert);
    assert.deepStrictEqual(out.items.map((i) => i.label), ['def', 'my_var']);
    assert.strictEqual(out.isIncomplete, true);
    assert.throws(() => core.mergeCompletionResults(rejected, [], mergeConvert), /lsp down/);
});

// Entries copied from lispython tests/test_dx.py (test_complete_items_wire_shape).
const FILTERED_MY_V = { op: 'complete-result', word: 'MY_v', items: [{ label: 'my-var', filterText: 'MY_var', kind: 6 }] };
const FILTERED_WHI = { op: 'complete-result', word: 'whi', items: [{ label: 'while', filterText: 'while', kind: 14, labelDetails: { description: 'special form' } }] };

test('nreplCompletionItems: filterText passes through, distinct from the label', () => {
    const [item] = core.nreplCompletionItems(FILTERED_MY_V);
    assert.strictEqual(item.label, 'my-var');
    assert.strictEqual(item.filterText, 'MY_var');
    const [tagged] = core.nreplCompletionItems(FILTERED_WHI);
    assert.strictEqual(tagged.filterText, 'while');
    assert.strictEqual(tagged.description, 'special form');
});

test('mergeCompletionResults: nothing to add returns the LSP result untouched', () => {
    const list = { items: [lspItem('def')], isIncomplete: true };
    assert.strictEqual(core.mergeCompletionResults(settled(list), [mergeNrepl[0]], mergeConvert), list);
});
