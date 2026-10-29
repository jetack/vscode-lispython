const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const Module = require('module');

const bundlePath = process.env.LISPY_BUNDLE
    ? path.resolve(process.cwd(), process.env.LISPY_BUNDLE)
    : path.join(__dirname, '../out/extension.js');

// A Proxy that answers any property access or call with another harmless stub.
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

test('bundle is self-contained', () => {
    const requests = [];
    const originalLoad = Module._load;
    Module._load = function (request, ...rest) {
        requests.push(request);
        // esbuild's import interop copies own props onto an object whose prototype is the
        // module's prototype, so the stub's catch-all lives on the prototype.
        if (request === 'vscode') return Object.create(makeStub());
        return originalLoad.call(this, request, ...rest);
    };
    let ext;
    try {
        ext = require(bundlePath);
    } finally {
        Module._load = originalLoad;
    }

    assert.strictEqual(typeof ext.activate, 'function');
    assert.strictEqual(typeof ext.deactivate, 'function');

    const unexpected = requests.filter(
        (r) => r !== 'vscode' && !Module.isBuiltin(r) && r !== bundlePath,
    );
    assert.deepStrictEqual(unexpected, []);

    const source = fs.readFileSync(bundlePath, 'utf8');
    assert.ok(!source.includes('require("vscode-languageclient'));
    if (process.env.LISPY_EXPECT_PRODUCTION === '1') {
        assert.ok(!source.includes('sourceMappingURL'));
    }
});
