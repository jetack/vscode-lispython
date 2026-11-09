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
        if (c.caret < span.leafStart) exempt.push({ c, leafStart: span.leafStart });
        const [got, want] = [vscodeWord(c.text, c.caret), expectedWord(span, c.caret)];
        if (JSON.stringify(got) !== JSON.stringify(want)) wrong.push(`${show(c.text, c.caret)}: ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
        assertNoEmptyMatch(c.text);
    }
    assert.deepStrictEqual(wrong, []);
    // Only a caret inside a real peeled front is exempt: the spec's word there is not empty
    // and starts before the leaf, and only peelable entries lie between, so a leafStart
    // past the true one cannot hide a regex word.
    const peelable = new RegExp(`^(?:${[...rule.tokenPrefixes, ...rule.signs]
        .map((e) => e.replace(/[\\^$.*+?()[\]{}|/]/g, '\\$&')).join('|')})+$`, 'u');
    for (const { c, leafStart } of exempt) {
        const at = show(c.text, c.caret);
        assert.ok(c.replaceStart < c.end, `${at}: the spec word is empty`);
        assert.ok(c.start < leafStart, `${at}: the spec word starts at the leaf`);
        assert.match(c.text.slice(c.start, leafStart), peelable, `${at}: the front holds more than peels`);
    }
});

test('the wordPattern stays fast on long runs of signs, leads and dots', () => {
    // A column inside a signs run or a `....` stretch must fail the regex's cheap checks:
    // re-lexing the run there is quadratic per exec, and VS Code's 150 ms budget is
    // checked only between them.
    const core = require('../out-test/nreplCore');
    const rule = core.wordBoundaryRule();
    const lines = [
        'x' + '-'.repeat(999), '-'.repeat(500) + '1'.repeat(499), '+-'.repeat(200) + 'ab', '-'.repeat(400),
        '('.repeat(500) + '.'.repeat(499), "'".repeat(500) + '.'.repeat(499), '*'.repeat(500) + '.'.repeat(499),
    ];
    for (const line of lines) {
        for (const caret of [line.length, line.length >> 1, line.length - 20, 1]) {
            const t0 = Date.now();
            const word = vscodeWord(line, caret);
            const ms = Date.now() - t0;
            const label = `${JSON.stringify(line.slice(0, 12))}… (${line.length}) at ${caret}`;
            assert.ok(ms < 50, `${label}: ${ms} ms`);
            assert.deepStrictEqual(word, expectedWord(core.completionWordSpan(line, caret, rule), caret), label);
        }
    }
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

// --- VS Code's word selection (editor.wordSeparators) --------------------------------

// Ported from VS Code 1.82.0: src/vs/editor/common/core/wordCharacterClassifier.ts
// (WordCharacterClassifier) and src/vs/editor/common/cursor/cursorWordOperations.ts
// (WordOperations.word, not in selection mode, with _doFindPreviousWordOnLine,
// _doFindNextWordOnLine, _findEndOfWord, _findStartOfWord), types and the model dropped.
const WordCharacterClass = { Regular: 0, Whitespace: 1, WordSeparator: 2 };
const WordType = { None: 0, Regular: 1, Separator: 2 };

function wordCharacterClassifier(wordSeparators) {
    const classes = new Map();
    for (let i = 0, len = wordSeparators.length; i < len; i++) {
        classes.set(wordSeparators.charCodeAt(i), WordCharacterClass.WordSeparator);
    }
    classes.set(32, WordCharacterClass.Whitespace);
    classes.set(9, WordCharacterClass.Whitespace);
    return { get: (chCode) => classes.get(chCode) ?? WordCharacterClass.Regular };
}

function findEndOfWord(lineContent, wordSeparators, wordType, startIndex) {
    const len = lineContent.length;
    for (let chIndex = startIndex; chIndex < len; chIndex++) {
        const chClass = wordSeparators.get(lineContent.charCodeAt(chIndex));
        if (chClass === WordCharacterClass.Whitespace) return chIndex;
        if (wordType === WordType.Regular && chClass === WordCharacterClass.WordSeparator) return chIndex;
        if (wordType === WordType.Separator && chClass === WordCharacterClass.Regular) return chIndex;
    }
    return len;
}

function findStartOfWord(lineContent, wordSeparators, wordType, startIndex) {
    for (let chIndex = startIndex; chIndex >= 0; chIndex--) {
        const chClass = wordSeparators.get(lineContent.charCodeAt(chIndex));
        if (chClass === WordCharacterClass.Whitespace) return chIndex + 1;
        if (wordType === WordType.Regular && chClass === WordCharacterClass.WordSeparator) return chIndex + 1;
        if (wordType === WordType.Separator && chClass === WordCharacterClass.Regular) return chIndex + 1;
    }
    return 0;
}

function findPreviousWordOnLine(lineContent, wordSeparators, column) {
    let wordType = WordType.None;
    for (let chIndex = column - 2; chIndex >= 0; chIndex--) {
        const chClass = wordSeparators.get(lineContent.charCodeAt(chIndex));
        if (chClass === WordCharacterClass.Regular) {
            if (wordType === WordType.Separator) {
                return { wordType, start: chIndex + 1, end: findEndOfWord(lineContent, wordSeparators, wordType, chIndex + 1) };
            }
            wordType = WordType.Regular;
        } else if (chClass === WordCharacterClass.WordSeparator) {
            if (wordType === WordType.Regular) {
                return { wordType, start: chIndex + 1, end: findEndOfWord(lineContent, wordSeparators, wordType, chIndex + 1) };
            }
            wordType = WordType.Separator;
        } else if (chClass === WordCharacterClass.Whitespace) {
            if (wordType !== WordType.None) {
                return { wordType, start: chIndex + 1, end: findEndOfWord(lineContent, wordSeparators, wordType, chIndex + 1) };
            }
        }
    }
    if (wordType !== WordType.None) {
        return { wordType, start: 0, end: findEndOfWord(lineContent, wordSeparators, wordType, 0) };
    }
    return null;
}

function findNextWordOnLine(lineContent, wordSeparators, column) {
    let wordType = WordType.None;
    const len = lineContent.length;
    for (let chIndex = column - 1; chIndex < len; chIndex++) {
        const chClass = wordSeparators.get(lineContent.charCodeAt(chIndex));
        if (chClass === WordCharacterClass.Regular) {
            if (wordType === WordType.Separator) {
                return { wordType, start: findStartOfWord(lineContent, wordSeparators, wordType, chIndex - 1), end: chIndex };
            }
            wordType = WordType.Regular;
        } else if (chClass === WordCharacterClass.WordSeparator) {
            if (wordType === WordType.Regular) {
                return { wordType, start: findStartOfWord(lineContent, wordSeparators, wordType, chIndex - 1), end: chIndex };
            }
            wordType = WordType.Separator;
        } else if (chClass === WordCharacterClass.Whitespace) {
            if (wordType !== WordType.None) {
                return { wordType, start: findStartOfWord(lineContent, wordSeparators, wordType, chIndex - 1), end: chIndex };
            }
        }
    }
    if (wordType !== WordType.None) {
        return { wordType, start: findStartOfWord(lineContent, wordSeparators, wordType, len - 1), end: len };
    }
    return null;
}

/** The selection double-click makes at UTF-16 offset `caret` of the one-line `text`, as [start, end]. */
function doubleClickWord(text, caret, separators) {
    const wordSeparators = wordCharacterClassifier(separators);
    const column = caret + 1;
    const prevWord = findPreviousWordOnLine(text, wordSeparators, column);
    const nextWord = findNextWordOnLine(text, wordSeparators, column);
    let startColumn;
    let endColumn;
    if (prevWord && prevWord.wordType === WordType.Regular && prevWord.start <= column - 1 && column - 1 <= prevWord.end) {
        startColumn = prevWord.start + 1;
        endColumn = prevWord.end + 1;
    } else if (nextWord && nextWord.wordType === WordType.Regular && nextWord.start <= column - 1 && column - 1 <= nextWord.end) {
        startColumn = nextWord.start + 1;
        endColumn = nextWord.end + 1;
    } else {
        startColumn = prevWord ? prevWord.end + 1 : 1;
        endColumn = nextWord ? nextWord.start + 1 : text.length + 1;
    }
    // new Range(...) swaps a start past the end (`*|*x`: prev ends after next starts).
    return [Math.min(startColumn, endColumn) - 1, Math.max(startColumn, endColumn) - 1];
}

// VS Code 1.82's default (src/vs/editor/common/core/wordHelper.ts USUAL_WORD_SEPARATORS),
// what the editor uses when package.json sets none: the test then fails on content.
const VSCODE_DEFAULT_SEPARATORS = '`~!@#$%^&*()-=+[{]}\\|;:\'",.<>/?';

function lispythonWordSeparators() {
    const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
    return pkg.contributes.configurationDefaults['[lispython]']['editor.wordSeparators'] ?? VSCODE_DEFAULT_SEPARATORS;
}

// Spec cases (text, caret in code points) where one class per character cannot give the
// spec's start..end, by reason. Anything else that differs fails the test, and so does an
// entry here that stops differing.
const SEPARATOR_DIVERGENCES = {
    // A sign is no separator (or `->f` and `swap!` would split), so the front stays on.
    'a peeled sign front is outside the spec word': [
        ['-fo', 3], ['--x', 3], ['(- -fo', 6], ['-.x', 3], ['-_a', 3], ['-²', 2], ['-1', 2], ['-foo', 1],
    ],
    // Intended: double-click selects the attribute, as in Python.
    'a dotted path is one completion word, but `.` separates for selection': [
        ['heapq.nl', 8], ['heapq.', 6], ['a.b.c', 3], ['a..b', 4], ['x.', 2], ['.method', 7], ['(.c', 3],
        ['x .c', 4], ['1.5', 3],
    ],
    'the spec word at a string stops at its quote, which no class can say': [
        ['a"', 1], ['f"x"', 1], ['한"', 1], ['"ab"', 3],
    ],
    'a meta char, prefix or opening is a separator even where it is part of a symbol': [
        ['a~b', 3], ['a@b', 3], ['a(b', 3], ["a'b", 3], ['except*', 7], ['**x', 1], ['**=', 3], ['^**=', 4],
        ['~', 1], ['~@x', 1], ['*[1 2]', 1],
    ],
};

test('double-click with the [lispython] editor.wordSeparators selects the spec word', () => {
    const separators = lispythonWordSeparators();
    const known = new Map();
    for (const [reason, cases] of Object.entries(SEPARATOR_DIVERGENCES)) {
        for (const [text, caret] of cases) known.set(JSON.stringify([text, caret]), reason);
    }
    const wrong = [];
    const diverged = new Set();
    for (const spec of SPEC.cases) {
        const c = utf16Case(spec);
        const got = doubleClickWord(c.text, c.caret, separators);
        const id = JSON.stringify([spec.text, spec.caret]);
        const differs = JSON.stringify(got) !== JSON.stringify([c.start, c.end]);
        if (differs && known.has(id)) diverged.add(id);
        else if (differs) wrong.push(`${show(c.text, c.caret)}: ${JSON.stringify(got)}, want ${JSON.stringify([c.start, c.end])}`);
    }
    const stale = [...known.keys()].filter((id) => !diverged.has(id));
    assert.deepStrictEqual(wrong, [], 'spec cases the separators get wrong');
    assert.deepStrictEqual(stale, [], 'listed divergences that are not (or no longer) spec cases that differ');
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
