// vsc-lispy's completion word boundary against lispython's published spec
// (test/word-boundary.json, `python -m lispy.word_boundary`): caretContext,
// completionWordSpan and symbolPrefixStart directly, and the language-configuration.json
// wordPattern through VS Code's own word lookup.
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

/** The spec case in UTF-16 offsets of its line (null bounds stay null). */
function utf16Case(c) {
    const at = (cp) => (cp === null ? null : utf16(c.text, cp));
    return { before: c.before, text: c.text, caret: at(c.caret), start: at(c.start), replaceStart: at(c.replaceStart), end: at(c.end) };
}

/** The spec case's input as completionWordSpan takes it: the text before the line and the line, and the caret in it. */
const caseInput = (c) => ({ text: c.before + c.text, caret: c.before.length + c.caret });

/** A spec case's key in a divergence table: [text, caret], or [before, text, caret] after other lines. */
const caseKey = (spec) => JSON.stringify(spec.before ? [spec.before, spec.text, spec.caret] : [spec.text, spec.caret]);

/** The span of the spec case `c` (UTF-16) by `core`, in offsets of the case's line, or null. */
function lineSpan(core, c, rule) {
    const input = caseInput(c);
    const span = core.completionWordSpan(input.text, input.caret, rule);
    return span === null ? null : Object.fromEntries(Object.entries(span).map(([k, v]) => [k, v - c.before.length]));
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
    assert.ok(SPEC.cases.some((c) => c.start === null) && SPEC.cases.some((c) => c.before !== ''), 'the spec has text and multi-line cases');
    for (const c of SPEC.cases.map(utf16Case)) {
        const span = lineSpan(core, c, rule);
        const got = span === null ? [null, null, null] : [span.start, span.replaceStart, span.end];
        const want = [c.start, c.replaceStart, c.end];
        const at = show(c.before + c.text, c.before.length + c.caret);
        if (JSON.stringify(got) !== JSON.stringify(want)) wrong.push(`completionWordSpan ${at}: ${got}, want ${want}`);
        const input = caseInput(c);
        const prefixStart = core.symbolPrefixStart(input.text, input.caret, rule);
        const wantStart = c.start === null ? null : c.before.length + c.start;
        if (prefixStart !== wantStart) wrong.push(`symbolPrefixStart ${at}: ${prefixStart}, want ${wantStart}`);
    }
    assert.deepStrictEqual(wrong, []);
});

buildTest('completionWordSpan reads the context from earlier CRLF lines', (core) => {
    const rule = core.wordBoundaryRule();
    const span = (text, caret) => {
        const s = core.completionWordSpan(text, caret, rule);
        return s === null ? null : [s.start, s.replaceStart, s.end];
    };
    // Line 2 (from 6) is in the f"""-literal opened on line 1; its field's { is at 8, so `ab`
    // (9-11) is a field's word, ending at the spec's `:` (11); read alone as code, the line
    // would make `ab:>4` one symbol (9-14).
    const text = 'f"""\r\n  {ab:>4}\r\n"""';
    assert.deepStrictEqual(span(text, 11), [9, 9, 11]);
    assert.deepStrictEqual(span(text, 10), [9, 9, 11]);
    // Between `>` (12) and `4` (13): the format spec's text, no word.
    assert.strictEqual(span(text, 13), null);
});

// VS Code's word lookups see the caret's line alone, so they cannot read the caret's context
// from the text's start. The spec cases whose word differs for that reason, by reason
// ([text, caret] in code points, [before, text, caret] after other lines). Anything else
// that differs fails the test, and so does an entry here that stops differing.
const WORD_PATTERN_DIVERGENCES = {
    "no completion word in a string's or comment's text, while VS Code's word lookup still finds one": [
        ['; ab', 4], ['f"{x!r}"', 5], ['f"{x:>10}"', 7], ['f"\\N{ab}"', 7], ['f"""\n', '  text {x}', 4], ['(print "ab', 10],
        ['a;b', 3],
    ],
    "an f/t-string field is a context read from the text's start (its delimiters and debug = end the word); the wordPattern lexes the line as code": [
        ['f"{x!r}"', 4], ['f"{x:>10}"', 4], ['f"{x=}"', 4], ['f"{x =}"', 6], ['f"{x=}"', 5], ['t"{x!r}"', 4], ['f"""{x:>4}"""', 6],
        ['f"""\n', '  Name: {name!r}', 13], ['f"""\n', '  Width: {w:>10}', 11], ['f"""\n', '  {x=}', 4], ['f"{x!r}', 4],
        ['f"{a\\=}"', 5], ['f"""{x=', 6], ['f"""{x= ', 6], ['f"{x=\\', 4], ['f"}"\n', 'f"{x!r}"', 4], ['f"{\\=}"', 4], ['f"{\\=}"', 5],
    ],
    'a word starts after a string closed earlier in its run, which the wordPattern lexes from a stop character': [
        ["(f '''a-b''')", 12], ["'''x:'''", 8], ["f\"{'''a-b'''}\"", 12],
    ],
};

/** The divergence table `table` as a map from caseKey to reason. */
function divergences(table) {
    const known = new Map();
    for (const [reason, cases] of Object.entries(table)) {
        for (const key of cases) known.set(JSON.stringify(key), reason);
    }
    return known;
}

buildTest('VS Code finds the spec word of every spec case with the wordPattern', (core) => {
    const rule = core.wordBoundaryRule();
    const known = divergences(WORD_PATTERN_DIVERGENCES);
    const diverged = new Set();
    const exempt = [];
    const wrong = [];
    for (const spec of SPEC.cases) {
        const c = utf16Case(spec);
        const span = lineSpan(core, c, rule);
        if (span !== null && c.caret < span.leafStart) exempt.push({ c, leafStart: span.leafStart });
        const [got, want] = [vscodeWord(c.text, c.caret), span === null ? null : expectedWord(span, c.caret)];
        const differs = JSON.stringify(got) !== JSON.stringify(want);
        if (differs && known.has(caseKey(spec))) diverged.add(caseKey(spec));
        else if (differs) wrong.push(`${caseKey(spec)}: ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
        assertNoEmptyMatch(c.text);
    }
    assert.deepStrictEqual(wrong, [], 'spec cases the wordPattern gets wrong');
    const stale = [...known.keys()].filter((id) => !diverged.has(id));
    assert.deepStrictEqual(stale, [], 'listed divergences that are not (or no longer) spec cases that differ');
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

// Spec cases ([text, caret] in code points, [before, text, caret] after other lines) where
// one class per character on the caret's line cannot give the spec's start..end, by reason. Anything else that differs fails the test, and so does an
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
        ['~', 1], ['~@x', 1], ['*[1 2]', 1], ['f"{(a:b)}"', 5], ['{x:y}', 2],
    ],
    // Intended: double-click selects a word wherever it stands.
    "a string's or comment's text holds no completion word, but double-click selects one there": [
        ["x'''", 4], ['; ab', 4], ['f"{{ab"', 6], ['f"{x!r}"', 5], ['f"{x:>10}"', 7], ['f"\\N{ab}"', 7], ["'''ab'''", 5],
        ['f"""\n', '  text {x}', 4], ['"a\n', 'b"', 1], ['(print "ab', 10], ['f"{x!}"\n', '(print "abc")', 10],
        ['x = f"{a:\n}"\n', '"abc"', 4], ['a;b', 3],
    ],
    "an f/t-string field is a context read from the text's start (its delimiters and debug = end the word); a class per character knows none": [
        ['f"{x!r}"', 4], ['f"{x=}"', 4], ['f"{x =}"', 6], ['f"{x=}"', 5], ['t"{x!r}"', 4], ['f"""\n', '  Name: {name!r}', 13],
        ['f"""\n', '  {x=}', 4], ['f"{x!r}', 4], ['f"{a\\=}"', 5], ['f"""{x=', 6], ['f"""{x= ', 6], ['f"{x=\\', 4],
        ['f"}"\n', 'f"{x!r}"', 4], ['f"{\\=}"', 4], ['f"{\\=}"', 5],
    ],
    'a word starts after a string closed earlier in its run': [
        ["(f '''a-b''')", 12], ["f\"{'''a-b'''}\"", 12],
    ],
};

test('double-click with the [lispython] editor.wordSeparators selects the spec word', () => {
    const separators = lispythonWordSeparators();
    const known = divergences(SEPARATOR_DIVERGENCES);
    const wrong = [];
    const diverged = new Set();
    for (const spec of SPEC.cases) {
        const c = utf16Case(spec);
        const got = doubleClickWord(c.text, c.caret, separators);
        const id = caseKey(spec);
        const want = c.start === null ? null : [c.start, c.end];
        const differs = JSON.stringify(got) !== JSON.stringify(want);
        if (differs && known.has(id)) diverged.add(id);
        else if (differs) wrong.push(`${id}: ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
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

// The f/t-string reader's own characters: fields, conversions, specs, the debug =, escapes
// (\, \N{), quotes, comments and line breaks, so its modes and refusals are exercised.
const FIELD_ALPHABET = [
    'f', 'r', 't', 'N', 'x', 'a', '1', '{', '}', ':', '!', '=', '\\', '"', "'", '(', ')', '[', ']', ';',
    ' ', '\n', '\r', '>', '-', '.', '_', '\u{1D465}',
];
// How a field text starts: in a literal, in its field, raw, a t-string, a triple-quoted
// string, at a conversion (a valid name, so what follows it is what the reader checks).
const FIELD_STARTS = ['f"', 'f"{', 'rf"', 't"', "'''", "f'''{", 'f"{x', '"', '', 'f"{x!r', 'f"{x!a'];

/**
 * The corpus: every short lexer line, random short lines, lines over 15 columns, and texts
 * of several lines: random lines, and random field texts from FIELD_STARTS.
 */
function corpus() {
    const random = prng(1184);
    const pick = (xs) => xs[Math.floor(random() * xs.length)];
    const chars = (alphabet, min, max) => {
        const n = min + Math.floor(random() * (max - min + 1));
        let s = '';
        for (let i = 0; i < n; i++) s += pick(alphabet);
        return s;
    };
    const line = (min, max) => chars(ALPHABET, min, max);
    const lines = exhaustiveLines(LEXER_ALPHABET, 3);
    for (let i = 0; i < 1500; i++) lines.push(line(1, 7));
    for (let i = 0; i < 300; i++) lines.push(line(16, 40));
    const texts = [];
    for (let i = 0; i < 100; i++) texts.push(`${line(1, 6)}\n${line(1, 6)}`);
    for (let i = 0; i < 3000; i++) texts.push(pick(FIELD_STARTS) + chars(FIELD_ALPHABET, 1, 12));
    return { lines, texts };
}

/** Every code point boundary of `text`, as UTF-16 offsets. */
function carets(text) {
    const out = [0];
    for (const c of text) out.push(out.at(-1) + c.length);
    return out;
}

/**
 * Whether the wordPattern can know the completion word at `caret` of the one-line `text`: the
 * caret is in top-level code with no string closed before it on the line (the regex re-lexes
 * the run from its start as code), or the reader refuses the text (the line is then code).
 */
function wordPatternReads(core, text, caret, rule) {
    const found = core.caretContext(text, caret, rule);
    return found === null || (found.context === 'code' && found.resume === 0);
}

buildTest('fuzz: the wordPattern word is the completionWordSpan word, and never empty', (core) => {
    const rule = core.wordBoundaryRule();
    const { lines } = corpus();
    let checked = 0;
    for (const text of lines) {
        assertNoEmptyMatch(text);
        for (const caret of carets(text)) {
            if (!wordPatternReads(core, text, caret, rule)) continue;
            const span = core.completionWordSpan(text, caret, rule);
            assert.deepStrictEqual(vscodeWord(text, caret), expectedWord(span, caret), show(text, caret));
            checked++;
        }
    }
    assert.ok(checked > 10000, `${checked} carets`);
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

// Reads [[before, line, col], ...] (code points) as JSON on stdin; writes, per query, its
// caret_context's context and resume (where the reader refuses the text before the caret,
// null and the error's message, null when it stopped with none) and the completion_word_span.
const SPANS_PY = `
import json, sys
from lispy.core.parser import caret_context, caret_tokens, completion_word_span
queries = json.loads(sys.stdin.buffer.read().decode("utf-8"))
out = []
for before, line, col in queries:
    found = caret_context(before + line, len(before) + col)
    span = completion_word_span(line, col, before)
    if found is None:
        tokens = caret_tokens(line, col, before)
        out.append([None, tokens[-1][0] if tokens and tokens[-1][1] == "error" else None, span])
    else:
        out.append([found[0], found[1], span])
json.dump(out, sys.stdout)
`;

// The one refusal caretContext does not make: a conversion name other than s, r and a, a
// set the spec does not publish (a client need not agree on a text the reader refuses).
// lispython card TASK-1362 asks for the set to be published.
const UNPORTED_REFUSAL = "expected 's', 'r', or 'a'";

/** The lexer's line breaks (split_lines). */
const LINE_BREAK = /\r\n|\r|\n/g;

/** [before, line, caret in the line] of UTF-16 offset `caret` of `text`, or null inside a CRLF. */
function lineInput(text, caret) {
    let lineStart = 0;
    for (const m of text.matchAll(LINE_BREAK)) {
        const end = m.index + m[0].length;
        if (end > caret) {
            if (m.index < caret) return null;
            return [text.slice(0, lineStart), text.slice(lineStart, m.index), caret - lineStart];
        }
        lineStart = end;
    }
    return [text.slice(0, lineStart), text.slice(lineStart), caret - lineStart];
}

buildTest('fuzz: caretContext and completionWordSpan are lispython\'s caret_context and completion_word_span', (core, t) => {
    if (checkout === undefined) return t.skip(lispythonMissing);
    const rule = core.wordBoundaryRule();
    const { lines, texts } = corpus();
    const queries = [];
    for (const text of [...lines, ...texts]) {
        for (const caret of carets(text)) {
            const input = lineInput(text, caret);
            if (input !== null) queries.push({ text, caret, input });
        }
    }
    const asCodePoints = ({ input: [before, line, col] }) => [before, line, codePoints(line, col)];
    const answers = JSON.parse(runPython(['-c', SPANS_PY], JSON.stringify(queries.map(asCodePoints))));
    assert.strictEqual(answers.length, queries.length);
    const seen = { checked: 0, text: 0, field: 0, multiLine: 0, refused: 0, lineAlone: 0 };
    queries.forEach(({ text, caret, input: [before, line] }, i) => {
        const [context, resume, span] = answers[i];
        const scanned = before + line;
        const at = show(text, caret);
        if (context === null) {
            // The reader refuses the text before the caret: so does caretContext, but for
            // the unported check, where the word is no concern (the client need not agree).
            if (resume !== null && resume.includes(UNPORTED_REFUSAL)) return;
            assert.strictEqual(core.caretContext(scanned, caret, rule), null, `caretContext ${at}, refused: ${resume}`);
            seen.refused++;
            // Then the word is the line-alone reading's (_line_context), compared below.
            if (before !== '') seen.lineAlone++;
        } else {
            assert.deepStrictEqual(core.caretContext(scanned, caret, rule), { context, resume: utf16(scanned, resume) }, `caretContext ${at}`);
        }
        // The whole text: what follows the caret's line is never read.
        const got = core.completionWordSpan(text, caret, rule);
        const expected = span === null ? null : span.map((cp) => before.length + utf16(line, cp));
        assert.deepStrictEqual(got === null ? null : [got.start, got.replaceStart, got.end], expected, `completionWordSpan ${at}`);
        seen.checked++;
        if (context === 'text') seen.text++;
        if (context === 'field' || context === 'bracket') seen.field++;
        if (before !== '') seen.multiLine++;
    });
    // The corpus reaches what the field alphabet is there for.
    assert.ok(seen.checked > 50000 && seen.text > 10000 && seen.field > 4000 && seen.multiLine > 4000 && seen.refused > 1000
        && seen.lineAlone > 100, JSON.stringify(seen));
});

// --- the pinned spec and its derived files are current --------------------------------

test('test/word-boundary.json is what lispython publishes', (t) => {
    if (checkout === undefined) return t.skip(lispythonMissing);
    const live = runPython(['-m', 'lispy.word_boundary']);
    assert.ok(live.equals(Buffer.from(SPEC_TEXT, 'utf8')),
        're-run `python -m lispy.word_boundary > test/word-boundary.json`, then `bun run word-boundary`');
});

test('src/wordBoundaryRule.ts, the wordPattern and package.json\'s editor.wordSeparators are generated from test/word-boundary.json', () => {
    const r = spawnSync(process.execPath, [path.join(ROOT, 'scripts', 'word-boundary.mjs'), '--check'], { encoding: 'utf8' });
    assert.strictEqual(r.status, 0, r.stderr);
});
