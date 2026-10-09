import test from 'node:test';
import assert from 'node:assert/strict';
import { compare, resolve, pendingCount, units, structure, signals, hash } from '../lib/review.mjs';

// Proposed integration tests. The final regression group intentionally exposes
// current bugs; adjust the import to ../lib/review.mjs when moving this file.
const fixtures = [
  ['empty', '', ''],
  ['whole insertion', '', '\\section{Relief}\nThe claim is dismissed.\n'],
  ['whole deletion', 'The claim is dismissed.\n', ''],
  ['line endings, BOM and final newline', '\uFEFFFirst.\r\n\r\nSecond.', '\uFEFFFirst revised.\r\n\r\nSecond.\r\n'],
  ['whitespace only', 'A. \n\nB.\n', 'A.\n\n\nB.\n'],
  ['negation', 'The debtor shall not waive this defence.\n', 'The debtor shall waive this defence.\n'],
  ['Swiss amount and date', 'CHF 125’000 by 31.10.2026.\n', 'CHF 152’000 by 15.11.2026.\n'],
  ['Unicode normalization', 'Caf\u00e9\n', 'Cafe\u0301\n'],
  ['footnote across paragraphs', 'Claim.\\footnote{First ground.\n\nSecond ground.}\n', 'Claim.\\footnote{First ground.\n\nNew ground.}\n'],
  ['enumeration', '\\begin{enumerate}\n\\item Pay.\n\n\\item Notify.\n\\end{enumerate}\n', '\\begin{enumerate}\n\\item Pay promptly.\n\n\\item Notify.\n\\end{enumerate}\n'],
  ['repeated paragraphs reordered', 'Admitted.\n\nDenied.\n\nAdmitted.\n', 'Denied.\n\nAdmitted.\n\nAdmitted.\n'],
  ['reference target', '\\label{old}\n\nSee \\ref{old}.\n', '\\label{new}\n\nSee \\ref{new}.\n'],
  ['preamble', '\\documentclass{article}\n\\newcommand{\\Party}{Alpha}\n\\begin{document}\n\\Party\n\\end{document}', '\\documentclass{article}\n\\newcommand{\\Party}{Beta}\n\\begin{document}\n\\Party\n\\end{document}'],
  ['verbatim braces', '\\begin{verbatim}\n{ unmatched\n\n\\end{verbatim}\n\nClaim.\n', '\\begin{verbatim}\n} unmatched\n\n\\end{verbatim}\n\nClaim revised.\n'],
];

for (const [name, before, after] of fixtures) {
  test(`lossless endpoints: ${name}`, () => {
    assert.equal(units(before).join(''), before);
    assert.equal(units(after).join(''), after);
    const review = compare(before, after);
    const changes = review.segments.filter(s => s.type === 'change');
    const accepted = Object.fromEntries(changes.map(s => [s.id, 'accepted']));
    const rejected = Object.fromEntries(changes.map(s => [s.id, 'rejected']));
    assert.equal(resolve(review, accepted), after);
    assert.equal(resolve(review, rejected), before);
    assert.equal(resolve(review, {}, 'before'), before);
    assert.equal(resolve(review, {}, 'after'), after);
    assert.equal(review.beforeHash, hash(before));
    assert.equal(review.afterHash, hash(after));
    assert.equal(pendingCount(review, {}), changes.length);
    assert.equal(pendingCount(review, accepted), 0);
    assert.equal(pendingCount(review, rejected), 0);
  });
}

test('mixed decisions preserve an unchanged separator and exact source', () => {
  const review = compare('Old first.\n\nUntouched.\n\nOld last.\n', 'New first.\n\nUntouched.\n\nNew last.\n');
  const changes = review.segments.filter(s => s.type === 'change');
  assert.equal(changes.length, 2);
  assert.equal(resolve(review, {[changes[0].id]: 'accepted', [changes[1].id]: 'rejected'}), 'New first.\n\nUntouched.\n\nOld last.\n');
});

test('invalid decisions and source inputs are rejected', () => {
  const review = compare('Old.', 'New.');
  const id = review.segments.find(s => s.type === 'change').id;
  assert.throws(() => resolve(review, {[id]: 'accept'}), /Invalid review decision/);
  assert.throws(() => resolve(review, {'unknown-unit': 'accepted'}), /Invalid review decision/);
  assert.throws(() => compare(null, 'text'), /Two source documents/);
  assert.throws(() => compare('a'.repeat(300001), ''), /under 300 KB/);
});

test('balanced footnotes and lists remain indivisible', () => {
  assert.equal(units('Claim.\\footnote{First ground.\n\nSecond ground.}\n\nNext.\n').length, 2);
  assert.equal(units('\\begin{enumerate}\n\\item A\n\n\\item B\n\\end{enumerate}\n\nNext.\n').length, 2);
});

test('structure identifies broken braces and environment order', () => {
  assert.ok(structure('Text.\\footnote{Unclosed.').length);
  assert.ok(structure('\\begin{enumerate}\\begin{quote}\\end{enumerate}\\end{quote}').length);
  assert.deepEqual(structure('Escaped \\{ brace. % } \\end{ghost}\n'), []);
});

// Functional regression tests below fail against the initially inspected file.
test('adjacent revised paragraphs have independent decisions', () => {
  const review = compare('First.\n\nSecond.\n', 'First revised.\n\nSecond revised.\n');
  const changes = review.segments.filter(s => s.type === 'change');
  assert.equal(changes.length, 2);
  assert.equal(resolve(review, {[changes[0].id]: 'accepted', [changes[1].id]: 'rejected'}), 'First revised.\n\nSecond.\n');
});

test('verbatim literal braces do not swallow later paragraphs', () => {
  const source = '\\begin{verbatim}\n{ unmatched brace\n\n\\end{verbatim}\n\nNext paragraph.\n\nLast.\n';
  assert.deepEqual(structure(source), []);
  assert.equal(units(source).length, 3);
});

test('percent after an even backslash count starts a comment', () => {
  const source = 'A line break.\\\\% \\begin{invented} }\n\nNext paragraph.\n';
  assert.deepEqual(structure(source), []);
  assert.equal(units(source).length, 2);
});

test('environment changes raise a structure signal even when prose is identical', () => {
  assert.ok(signals('\\begin{enumerate}\\item Claim.\\end{enumerate}', '\\begin{itemize}\\item Claim.\\end{itemize}').some(s => s.label === 'Structure'));
});

test('numeric legal dates raise a timing signal', () => {
  assert.ok(signals('Due by 31.10.2026.', 'Due by 15.11.2026.').some(s => s.label === 'Timing'));
});

test('German negation changes raise an obligation signal', () => {
  assert.ok(signals('Es besteht kein Anspruch.', 'Es besteht ein Anspruch.').some(s => s.label === 'Obligation'));
});

test('parenthetical citations raise a reference signal', () => {
  assert.ok(signals('See \\parencite{old}.', 'See \\parencite{new}.').some(s => s.label === 'Reference'));
});

test('unchanged preamble does not absorb visible document units',()=>{
 const before='\\documentclass{article}\n\\begin{document}\n\n\\section{Parties}\n\nUnchanged body.\n\nOriginal.\n\n\\end{document}\n';
 const r=compare(before,before.replace('Original.','Proposed.'));
 assert.ok(r.segments.some(s=>!s.preamble&&s.heading==='Parties'));
 assert.ok(r.segments.some(s=>!s.preamble&&s.before.includes('Unchanged body.')));
});
test('format-only changes remain visible as source changes',()=>{
 const s=compare('A claim.','A \\textbf{claim}.').segments.find(s=>s.type==='change');
 assert.equal(s.sourceOnly,true);assert.ok(s.signals.length);
});
test('reading diff keeps a complete money value together',()=>{
 const s=compare('Pay CHF 125,000.','Pay CHF 150,000.').segments[0];
 assert.ok(s.words.some(w=>w.removed&&w.value==='125,000'));
 assert.ok(s.words.some(w=>w.added&&w.value==='150,000'));
});

test('blank lines in the preamble do not leak configuration into the reading view',()=>{
 const source=String.raw`\documentclass{article}

\newcommand{\Client}{Alpha}

\begin{document}

Body.

\end{document}
`;
 const r=compare(source,source.replace('Alpha','Beta'));
 assert.ok(r.segments.find(s=>s.before.includes('newcommand')).preamble);
 assert.ok(!r.segments.find(s=>s.before.includes('Body.')).preamble);
});
