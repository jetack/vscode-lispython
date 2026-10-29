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

function fakeConn(sock) {
    return { socket: sock, disconnects: 0, disconnect() { this.disconnects++; this.socket = undefined; } };
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
    assert.strictEqual(conn.disconnects, 1);
});

test('checkServerIdEcho: no reply keeps the connection, logs, and shows no toast', async () => {
    const sock = fakeSocket();
    const conn = fakeConn(sock);
    const n = notes();
    assert.strictEqual(await core.checkServerIdEcho(conn, sock, n, 30), 'inconclusive');
    assert.deepStrictEqual(n.toasts, []);
    assert.deepStrictEqual(n.logs, [core.NREPL_ID_ECHO_INCONCLUSIVE_MESSAGE]);
    assert.strictEqual(conn.disconnects, 0);
    assert.strictEqual(conn.socket, sock);
});

test('checkServerIdEcho: an echoing server is left alone and silent', async () => {
    const sock = fakeSocket();
    const conn = fakeConn(sock);
    const n = notes();
    const p = core.checkServerIdEcho(conn, sock, n, 1000);
    sock.reply({ id: sock.sent[0].id, completions: [] });
    assert.strictEqual(await p, 'echoed');
    assert.deepStrictEqual([n.logs, n.toasts, conn.disconnects], [[], [], 0]);
});

test('checkServerIdEcho: a verdict for a swapped-out socket changes nothing', async () => {
    const sock = fakeSocket();
    const conn = fakeConn(fakeSocket());
    const n = notes();
    const p = core.checkServerIdEcho(conn, sock, n, 5000);
    sock.reply({ completions: [] });
    assert.strictEqual(await p, 'missing');
    assert.deepStrictEqual([n.logs, n.toasts, conn.disconnects], [[], [], 0]);
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
