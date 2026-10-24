const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const core = require('../out/nreplCore');

test('kind mapping', () => {
    const kinds = { A: 'class', f: 'function', v: 'variable', m: 'macro', k: 'keyword', x: 'weird' };
    const got = Object.keys(kinds).map((n) => core.completionKindName(kinds, n));
    assert.deepStrictEqual(got, ['Class', 'Function', 'Variable', 'Function', 'Keyword', 'Function']);
    assert.strictEqual(core.completionKindName(undefined, 'a'), 'Function');
    assert.strictEqual(core.completionKindName({}, 'toString'), 'Function');
    assert.strictEqual(core.completionKindName({ a: 'constructor' }, 'a'), 'Function');
});

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

test('getReplHtml carries a nonce-based CSP on every script and style tag', () => {
    const html = core.getReplHtml('N0nce+/=', 'vscode-webview:');
    assert.ok(html.includes(`<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src vscode-webview: 'nonce-N0nce+/='; script-src 'nonce-N0nce+/=';">`));
    const tags = html.match(/<(script|style)\b[^>]*>/g);
    assert.strictEqual(tags.length, 2);
    for (const tag of tags) assert.ok(tag.includes('nonce="N0nce+/="'), tag);
    assert.ok(!/<script(?![^>]*nonce=)/.test(html));
    assert.notStrictEqual(core.makeNonce(), core.makeNonce());
});

test('injected escapeHtml/linkifyErrors work inside the webview script', () => {
    const vm = require('node:vm');
    const html = core.getReplHtml('n', 'c');
    const script = /<script nonce="n">([\s\S]*?)<\/script>/.exec(html)[1];
    let listener;
    const entries = { innerHTML: '', appendChild() {} };
    const ctx = {
        document: {
            getElementById: () => entries,
            addEventListener() {},
            createElement: () => ({ scrollIntoView() {} }),
        },
        window: { addEventListener: (type, fn) => { if (type === 'message') listener = fn; } },
        acquireVsCodeApi: () => ({ postMessage() {} }),
    };
    vm.createContext(ctx);
    vm.runInContext(script, ctx);
    const src = 'File "/a.py", line 3';
    assert.strictEqual(vm.runInContext(`linkifyErrors(${JSON.stringify(src)})`, ctx), core.linkifyErrors(src));
    assert.strictEqual(vm.runInContext(`escapeHtml('<b>')`, ctx), '&lt;b&gt;');

    let div;
    ctx.document.createElement = () => (div = { scrollIntoView() {} });
    listener({ data: { type: 'eval', code: '<b>', result: { error: src } } });
    assert.ok(div.innerHTML.includes('&lt;b&gt;'));
    assert.ok(div.innerHTML.includes('<span class="file-link" data-file="/a.py" data-line="3">'));
});
