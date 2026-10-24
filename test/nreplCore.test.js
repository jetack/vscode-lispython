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
