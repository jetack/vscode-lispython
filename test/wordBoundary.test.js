// vsc-lispy's completion word boundary against lispython's published spec
// (test/word-boundary.json, `python -m lispy.word_boundary`): completionWordSpan and
// symbolPrefixStart directly, and the language-configuration.json wordPattern through
// VS Code's own word lookup.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { findLispython, lispythonPython, lispythonEnv } = require('./lispython');

const ROOT = path.resolve(__dirname, '..');
const SPEC_TEXT = fs.readFileSync(path.join(__dirname, 'word-boundary.json'), 'utf8');
const SPEC = JSON.parse(SPEC_TEXT);
const WORD_BOUNDARY_PY = path.join('src', 'lispy', 'word_boundary.py');

// The production bundle (minified, as shipped) when `bun run test` built it, and the tsc output.
const builds = [['tsc', () => require('../out-test/nreplCore')], ['minified bundle', () => (
    process.env.LISPY_CORE_BUNDLE ? require(path.resolve(ROOT, process.env.LISPY_CORE_BUNDLE)) : undefined)]];

function buildTest(name, fn) {
    for (const [build, load] of builds) {
        test(`${name} [${build}]`, (t) => {
            const core = load();
            if (core === undefined) return t.skip('LISPY_CORE_BUNDLE not set; run `bun run test`');
            return fn(core, t);
        });
    }
}

/** UTF-16 offset of code point column `cp` in `text` (the spec counts code points). */
const utf16 = (text, cp) => [...text].slice(0, cp).join('').length;
/** Code point column of UTF-16 offset `u` in `text`. */
const codePoints = (text, u) => [...text.slice(0, u)].length;

/** The spec case in UTF-16 offsets. */
function utf16Case(c) {
    const at = (cp) => utf16(c.text, cp);
    return { text: c.text, caret: at(c.caret), start: at(c.start), replaceStart: at(c.replaceStart), end: at(c.end) };
}

const show = (text, caret) => JSON.stringify(text.slice(0, caret) + '|' + text.slice(caret));

// --- VS Code's word lookup -----------------------------------------------------------

// Copied from VS Code 1.82.0, src/vs/editor/common/core/wordHelper.ts (ensureValidWordDefinition,
// getWordAtText, _findRegexMatchEnclosingPosition), types and the config list dropped.
function ensureValidWordDefinition(wordDefinition) {
    let result = /(-?\d*\.\d\w*)|([^`~!@#$%^&*()\-=+[{\]}\\|;:'",.<>/?\s]+)/g;
    if (wordDefinition && (wordDefinition instanceof RegExp)) {
        if (!wordDefinition.global) {
            let flags = 'g';
            if (wordDefinition.ignoreCase) {
                flags += 'i';
            }
            if (wordDefinition.multiline) {
                flags += 'm';
            }
            if (wordDefinition.unicode) {
                flags += 'u';
            }
            result = new RegExp(wordDefinition.source, flags);
        } else {
            result = wordDefinition;
        }
    }
    result.lastIndex = 0;
    return result;
}

const DEFAULT_CONFIG = { maxLen: 1000, windowSize: 15, timeBudget: 150 };

function getWordAtText(column, wordDefinition, text, textOffset, config) {
    if (!config) {
        config = DEFAULT_CONFIG;
    }
    if (text.length > config.maxLen) {
        let start = column - config.maxLen / 2;
        if (start < 0) {
            start = 0;
        } else {
            textOffset += start;
        }
        text = text.substring(start, column + config.maxLen / 2);
        return getWordAtText(column, wordDefinition, text, textOffset, config);
    }
    const t1 = Date.now();
    const pos = column - 1 - textOffset;
    let prevRegexIndex = -1;
    let match = null;
    for (let i = 1; ; i++) {
        if (Date.now() - t1 >= config.timeBudget) {
            break;
        }
        const regexIndex = pos - config.windowSize * i;
        wordDefinition.lastIndex = Math.max(0, regexIndex);
        const thisMatch = _findRegexMatchEnclosingPosition(wordDefinition, text, pos, prevRegexIndex);
        if (!thisMatch && match) {
            break;
        }
        match = thisMatch;
        if (regexIndex <= 0) {
            break;
        }
        prevRegexIndex = regexIndex;
    }
    if (match) {
        const result = {
            word: match[0],
            startColumn: textOffset + 1 + match.index,
            endColumn: textOffset + 1 + match.index + match[0].length,
        };
        wordDefinition.lastIndex = 0;
        return result;
    }
    return null;
}

function _findRegexMatchEnclosingPosition(wordDefinition, text, pos, stopPos) {
    let match;
    while ((match = wordDefinition.exec(text))) {
        const matchIndex = match.index || 0;
        if (matchIndex <= pos && wordDefinition.lastIndex >= pos) {
            return match;
        } else if (stopPos > 0 && matchIndex > stopPos) {
            return null;
        }
    }
    return null;
}

/** language-configuration.json, read as VS Code reads it (JSONC: whole-line // comments). */
function languageConfiguration() {
    const text = fs.readFileSync(path.join(ROOT, 'language-configuration.json'), 'utf8');
    return JSON.parse(text.replace(/^[ \t]*\/\/.*$/gm, ''));
}

// As VS Code 1.82's _parseRegex (src/vs/workbench/contrib/codeEditor/browser/
// languageConfigurationExtensionPoint.ts) builds a wordPattern, string or object form.
const WORD_PATTERN = languageConfiguration().wordPattern;
const parsedWordPattern = typeof WORD_PATTERN === 'string'
    ? new RegExp(WORD_PATTERN, '')
    : new RegExp(WORD_PATTERN.pattern, WORD_PATTERN.flags);
// As ResolvedLanguageConfiguration.getWordDefinition hands it to getWordAtText.
const wordDefinition = () => ensureValidWordDefinition(parsedWordPattern);

/** VS Code's word at UTF-16 offset `caret` of the one-line `text`, as [start, end], or null. */
function vscodeWord(text, caret) {
    const word = getWordAtText(caret + 1, wordDefinition(), text, 0);
    return word === null ? null : [word.startColumn - 1, word.endColumn - 1];
}

/** Every match of the wordPattern over `text`, failing on an empty one (VS Code would hang). */
function assertNoEmptyMatch(text) {
    const re = wordDefinition();
    let m;
    while ((m = re.exec(text)) !== null) {
        assert.notStrictEqual(m[0], '', `the wordPattern matched empty at ${m.index} of ${JSON.stringify(text)}`);
    }
}

/**
 * The word VS Code must find for completion span `span` at `caret`: [replaceStart, end],
 * or null when that is empty. With the caret inside the peeled front (`-|-x`) the spec's
 * span starts at the token, which depends on the caret and so is no regex match: null.
 */
function expectedWord(span, caret) {
    if (caret < span.leafStart) return null;
    return span.replaceStart === span.end ? null : [span.replaceStart, span.end];
}

// --- spec cases ----------------------------------------------------------------------

buildTest('completionWordSpan and symbolPrefixStart give every spec case', (core) => {
    const rule = core.wordBoundaryRule();
    assert.ok(SPEC.cases.length > 50, 'the spec has its cases');
    const wrong = [];
    for (const c of SPEC.cases.map(utf16Case)) {
        const span = core.completionWordSpan(c.text, c.caret, rule);
        const got = [span.start, span.replaceStart, span.end];
        const want = [c.start, c.replaceStart, c.end];
        if (JSON.stringify(got) !== JSON.stringify(want)) wrong.push(`completionWordSpan ${show(c.text, c.caret)}: ${got}, want ${want}`);
        const prefixStart = core.symbolPrefixStart(c.text, c.caret, rule);
        if (prefixStart !== c.start) wrong.push(`symbolPrefixStart ${show(c.text, c.caret)}: ${prefixStart}, want ${c.start}`);
    }
    assert.deepStrictEqual(wrong, []);
});

buildTest('VS Code finds the spec word of every spec case with the wordPattern', (core) => {
    const rule = core.wordBoundaryRule();
    const exempt = [];
    const wrong = [];
    for (const c of SPEC.cases.map(utf16Case)) {
        const span = { ...c, leafStart: core.completionWordSpan(c.text, c.caret, rule).leafStart };
        if (c.caret < span.leafStart) exempt.push(show(c.text, c.caret));
        const [got, want] = [vscodeWord(c.text, c.caret), expectedWord(span, c.caret)];
        if (JSON.stringify(got) !== JSON.stringify(want)) wrong.push(`${show(c.text, c.caret)}: ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
        assertNoEmptyMatch(c.text);
    }
    assert.deepStrictEqual(wrong, []);
    // Only a caret inside a peeled front is exempt from the spec's span.
    assert.deepStrictEqual(exempt, ['"-|-x"', '"*|*x"']);
});

test('the wordPattern finds a word whose run starts more than 15 columns back', () => {
    // getWordAtText execs from 15-column windows, so the match must not start mid-word.
    const line = '(foo (:abcdefghijklmnopqrstuvw.xyz-0123456789abcdef ok?)';
    const leaf = line.indexOf('abc');
    for (let caret = leaf; caret <= line.indexOf('.'); caret++) {
        assert.deepStrictEqual(vscodeWord(line, caret), [leaf, line.indexOf('.')], show(line, caret));
    }
    const tail = line.indexOf('.') + 1;
    for (let caret = tail; caret <= line.indexOf(' ok?'); caret++) {
        assert.deepStrictEqual(vscodeWord(line, caret), [tail, line.indexOf(' ok?')], show(line, caret));
    }
});

// --- fuzz ----------------------------------------------------------------------------

/** A seeded PRNG (mulberry32), so the corpus is the same on every run. */
function prng(seed) {
    let a = seed >>> 0;
    return () => {
        a = (a + 0x6d2b79f5) >>> 0;
        let t = a;
        t = Math.imul(t ^ (t >>> 15), t | 1);
        t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}

// The characters the rule treats specially, plus plain letters, digits and spacing.
const ALPHABET = [
    '+', '-', '*', '^', ':', "'", '`', '~', '@', '(', ')', '[', ']', '{', '}', '"', '.', '0', '1', '9',
    'j', 'e', 'x', 'a', 'b', '한', '국', '\u{1D465}', '́', '²', ' ', ' ', '\u001c', '_', '=', '?',
    '!', '>', '#', ';', ',', '/', '<', '%',
];

// The lexer's and the peel's own characters, every line of up to 3 of them, so each
// short interaction (`^*~`, `:-1`, `*[`, `~@x`) is in the corpus, not just likely.
const LEXER_ALPHABET = ['^', '*', '~', '@', "'", '`', '(', '[', '"', '-', '+', ':', '.', '1', 'j', 'x', '_', '=', ' ', ')'];

function exhaustiveLines(alphabet, maxLength) {
    let level = [''];
    const out = [];
    for (let n = 1; n <= maxLength; n++) {
        level = level.flatMap((s) => alphabet.map((c) => s + c));
        out.push(...level);
    }
    return out;
}

/** The corpus: every short lexer line, random short lines, lines over 15 columns, and a few multi-line texts. */
function corpus() {
    const random = prng(1184);
    const line = (min, max) => {
        const n = min + Math.floor(random() * (max - min + 1));
        let s = '';
        for (let i = 0; i < n; i++) s += ALPHABET[Math.floor(random() * ALPHABET.length)];
        return s;
    };
    const lines = exhaustiveLines(LEXER_ALPHABET, 3);
    for (let i = 0; i < 1500; i++) lines.push(line(1, 7));
    for (let i = 0; i < 300; i++) lines.push(line(16, 40));
    const texts = [];
    for (let i = 0; i < 100; i++) texts.push(`${line(1, 6)}\n${line(1, 6)}`);
    return { lines, texts };
}

/** Every code point boundary of `text`, as UTF-16 offsets. */
function carets(text) {
    const out = [0];
    for (const c of text) out.push(out.at(-1) + c.length);
    return out;
}

buildTest('fuzz: the wordPattern word is the completionWordSpan word, and never empty', (core) => {
    const rule = core.wordBoundaryRule();
    const { lines } = corpus();
    let checked = 0;
    for (const text of lines) {
        assertNoEmptyMatch(text);
        for (const caret of carets(text)) {
            const span = core.completionWordSpan(text, caret, rule);
            assert.deepStrictEqual(vscodeWord(text, caret), expectedWord(span, caret), show(text, caret));
            checked++;
        }
    }
    assert.ok(checked > 10000, `${checked} carets`);
});

buildTest('fuzz: a newline bounds the word as the line end does', (core) => {
    const rule = core.wordBoundaryRule();
    for (const text of corpus().texts) {
        const nl = text.indexOf('\n');
        const parts = [[text.slice(0, nl), 0], [text.slice(nl + 1), nl + 1]];
        for (const [part, offset] of parts) {
            for (const caret of carets(part)) {
                const alone = core.completionWordSpan(part, caret, rule);
                const shifted = Object.fromEntries(Object.entries(alone).map(([k, v]) => [k, v + offset]));
                assert.deepStrictEqual(core.completionWordSpan(text, caret + offset, rule), shifted, show(text, caret + offset));
            }
        }
    }
});

// The live lispython, for the tests that compare against it rather than the pinned spec.
const checkout = findLispython(WORD_BOUNDARY_PY);
const lispythonMissing = 'lispython checkout not found; set LISPYTHON_DIR';

function runPython(args, input) {
    const r = spawnSync(lispythonPython(checkout), args, {
        cwd: checkout, env: lispythonEnv(checkout), input, maxBuffer: 64 * 1024 * 1024,
    });
    assert.strictEqual(r.status, 0, `${args.join(' ')}: ${r.stderr}`);
    return r.stdout;
}

// Reads [[text, caret], ...] (code points) as JSON on stdin; writes the spans as JSON.
const SPANS_PY = `
import json, sys
from lispy.core.parser import completion_word_span
queries = json.loads(sys.stdin.buffer.read().decode("utf-8"))
json.dump([completion_word_span(text, caret) for text, caret in queries], sys.stdout)
`;

buildTest('fuzz: completionWordSpan is lispython\'s completion_word_span', (core, t) => {
    if (checkout === undefined) return t.skip(lispythonMissing);
    const rule = core.wordBoundaryRule();
    const { lines, texts } = corpus();
    const queries = [];
    for (const text of [...lines, ...texts]) for (const caret of carets(text)) queries.push([text, caret]);
    const spans = JSON.parse(runPython(['-c', SPANS_PY], JSON.stringify(queries.map(([text, caret]) => [text, codePoints(text, caret)]))));
    assert.strictEqual(spans.length, queries.length);
    queries.forEach(([text, caret], i) => {
        const span = core.completionWordSpan(text, caret, rule);
        const expected = spans[i].map((cp) => utf16(text, cp));
        assert.deepStrictEqual([span.start, span.replaceStart, span.end], expected, show(text, caret));
    });
});

// --- the pinned spec and its derived files are current --------------------------------

test('test/word-boundary.json is what lispython publishes', (t) => {
    if (checkout === undefined) return t.skip(lispythonMissing);
    const live = runPython(['-m', 'lispy.word_boundary']);
    assert.ok(live.equals(Buffer.from(SPEC_TEXT, 'utf8')),
        're-run `python -m lispy.word_boundary > test/word-boundary.json`, then `bun run word-boundary`');
});

test('src/wordBoundaryRule.ts and the wordPattern are generated from test/word-boundary.json', () => {
    const r = spawnSync(process.execPath, [path.join(ROOT, 'scripts', 'word-boundary.mjs'), '--check'], { encoding: 'utf8' });
    assert.strictEqual(r.status, 0, r.stderr);
});
