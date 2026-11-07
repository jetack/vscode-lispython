const test = require('node:test');
const assert = require('node:assert');
const path = require('path');
const Module = require('module');

const bundlePath = process.env.LISPY_BUNDLE
    ? path.resolve(process.cwd(), process.env.LISPY_BUNDLE)
    : path.join(__dirname, '../out/extension.js');

function makeStub() {
    const fn = function () {};
    return new Proxy(fn, {
        get: (_t, prop) =>
            prop === Symbol.toPrimitive ? () => ''
            : prop === 'then' ? undefined
            : makeStub(),
        apply: () => makeStub(),
        construct: () => makeStub(),
    });
}

// Loads the bundle against a vscode mock whose command registry, active editor, terminals and
// error messages are observable, and returns the registered `lispython.runFile` handler.
async function loadRunFile(doc) {
    const commands = new Map();
    const terminals = [];
    const errors = [];
    const vscode = Object.create(makeStub());
    vscode.commands = Object.assign(Object.create(makeStub()), {
        registerCommand: (id, fn) => { commands.set(id, fn); return { dispose() {} }; },
    });
    vscode.window = Object.assign(Object.create(makeStub()), {
        activeTextEditor: { document: doc },
        createTerminal: (opts) => {
            const t = { opts, sent: [], show() {}, sendText(s) { t.sent.push(s); } };
            terminals.push(t);
            return t;
        },
        showErrorMessage: (m) => { errors.push(m); },
    });
    const originalLoad = Module._load;
    Module._load = function (request, ...rest) {
        if (request === 'vscode') return vscode;
        return originalLoad.call(this, request, ...rest);
    };
    try {
        delete require.cache[bundlePath];
        const ext = require(bundlePath);
        try {
            await ext.activate({ subscriptions: [], extensionPath: '/x' });
        } catch { /* commands register first, before anything that needs a real host */ }
    } finally {
        Module._load = originalLoad;
    }
    return { run: commands.get('lispython.runFile'), terminals, errors };
}

const doc = (over) => ({
    languageId: 'lispython', isDirty: false, isUntitled: false,
    uri: { scheme: 'file', fsPath: '/tmp/a.lpy' }, save: async () => true, ...over,
});

test('runFile refuses an untitled buffer instead of running lpy on its name', async () => {
    const { run, terminals, errors } = await loadRunFile(
        doc({ isUntitled: true, uri: { scheme: 'untitled', fsPath: 'Untitled-1' } }));
    assert.strictEqual(typeof run, 'function');
    await run();
    assert.deepStrictEqual(terminals, []);
    assert.strictEqual(errors.length, 1);
    assert.match(errors[0], /[Ss]ave/);
});

test('runFile runs lpy on a saved file', async () => {
    const { run, terminals, errors } = await loadRunFile(doc());
    await run();
    assert.deepStrictEqual(errors, []);
    assert.strictEqual(terminals.length, 1);
    assert.match(terminals[0].sent[0], /"\/tmp\/a\.lpy"$/);
});
