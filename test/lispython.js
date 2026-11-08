// Locating a lispython checkout for the tests that pin vsc-lispy against it.
const fs = require('fs');
const path = require('path');

/**
 * The lispython checkout holding `rel`: $LISPYTHON_DIR, else a `lispython` checkout beside
 * the repo or any ancestor (worktrees nest). Undefined when none has it.
 */
function findLispython(rel) {
    if (process.env.LISPYTHON_DIR) {
        return fs.existsSync(path.join(process.env.LISPYTHON_DIR, rel)) ? process.env.LISPYTHON_DIR : undefined;
    }
    let dir = path.resolve(__dirname, '..');
    while (path.dirname(dir) !== dir) {
        const candidate = path.join(path.dirname(dir), 'lispython');
        if (fs.existsSync(path.join(candidate, rel))) return candidate;
        dir = path.dirname(dir);
    }
    return undefined;
}

/** The Python to run `checkout`'s sources with: its .venv's when present, else python3. */
function lispythonPython(checkout) {
    const venv = path.join(checkout, '.venv', 'bin', 'python');
    return fs.existsSync(venv) ? venv : 'python3';
}

/** The environment that puts `checkout`'s src first on sys.path. */
function lispythonEnv(checkout) {
    const src = path.join(checkout, 'src');
    const rest = process.env.PYTHONPATH;
    return { ...process.env, PYTHONPATH: rest ? `${src}${path.delimiter}${rest}` : src };
}

module.exports = { findLispython, lispythonPython, lispythonEnv };
