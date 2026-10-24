const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const net = require('net');
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

test('nreplRequestOn resolves with the first response line and removes its listeners', async (t) => {
    const a = await listen(t);
    a.on('connection', (s) => s.on('data', () => s.write('{"value": "3"}\n{"value": "ignored"}\n')));
    const sock = await clientTo(t, a);
    const before = requestListeners(sock);
    assert.deepStrictEqual(await core.nreplRequestOn(sock, { op: 'eval', code: '(+ 1 2)' }), { value: '3' });
    assert.deepStrictEqual(requestListeners(sock), before);
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
    a.on('connection', (s) => s.once('data', () => {
        s.write('{"value": "sp');
        setTimeout(() => s.write('lit"}\n'), 50);
    }));
    const sock = await clientTo(t, a);
    let chunks = 0;
    sock.on('data', () => chunks++);
    assert.deepStrictEqual(await core.nreplRequestOn(sock, { op: 'eval', code: '1' }), { value: 'split' });
    assert.strictEqual(chunks, 2, 'the response must really arrive in two reads');
});

test('nreplRequestOn keeps a multi-byte character split across reads', async (t) => {
    const bytes = Buffer.from('{"value": "한국어"}\n');
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
