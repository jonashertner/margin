import test, {before, after} from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp, readFile, writeFile, rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {compare, resolve} from '../lib/review.mjs';
import {reviseParagraph} from '../lib/authoring.mjs';

// All persistence and generated documents stay in an isolated, disposable directory.
const work = await mkdtemp(path.join(tmpdir(), 'folio-word-tests-'));
process.env.FOLIO_DATA_DIR = path.join(work, 'data');
process.env.MARGIN_DATA_DIR = process.env.FOLIO_DATA_DIR;
const {wordAvailable, importWord, exportWord, originalWord, restoreWordFiles,
  importText, mapComparisonComments, anchorComments, MAX_WORD_BYTES} = await import('../lib/word.mjs');
const fixtures = fileURLToPath(new URL('./fixtures/word/', import.meta.url));
const manifest = JSON.parse(await readFile(path.join(fixtures, 'expected-manifest.json'), 'utf8'));
const {wordPython:python} = await import('../lib/runtime.mjs');
const run = promisify(execFile);
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const norm = text => text.replace(/\s+/g, ' ').trim();
const notesOf = result => Object.values(result.comments).flat();
const openFixture = async name => importWord({name, data:(await readFile(path.join(fixtures, name))).toString('base64')});
let native;

before(async () => {
  assert.equal(await wordAvailable(), true, 'Word integration dependencies must be installed for this suite');
  native = await openFixture('03-native-tracked.docx');
});
after(async () => { await rm(work, {recursive:true, force:true}); });

async function inspectDocx(base64, filename) {
  const file = path.join(work, filename);
  await writeFile(file, Buffer.from(base64, 'base64'));
  const script = `import json,sys,zipfile,xml.etree.ElementTree as E
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
with zipfile.ZipFile(sys.argv[1]) as z:
 for name in z.namelist():
  if name.endswith(('.xml','.rels')):E.fromstring(z.read(name))
 d=E.fromstring(z.read('word/document.xml'))
 text=lambda el: ''.join(x.text or '' for x in el.iter() if x.tag in (W+'t',W+'delText'))
 comments=[]
 if 'word/comments.xml' in z.namelist():
  for c in E.fromstring(z.read('word/comments.xml')).findall(W+'comment'):
   comments.append({'author':c.get(W+'author'),'at':c.get(W+'date'),'text':text(c)})
 footnotes=[]
 if 'word/footnotes.xml' in z.namelist():
  footnotes=[text(f) for f in E.fromstring(z.read('word/footnotes.xml')).findall(W+'footnote') if int(f.get(W+'id','0'))>0]
 core=E.fromstring(z.read('docProps/core.xml')) if 'docProps/core.xml' in z.namelist() else []
 fields=('creator','lastModifiedBy','title','subject','description','keywords','category','contentStatus')
 metadata={e.tag.rsplit('}',1)[-1]:e.text or '' for e in core if e.tag.rsplit('}',1)[-1] in fields}
 print(json.dumps({'insertions':len(list(d.iter(W+'ins'))),'deletions':len(list(d.iter(W+'del'))),'revisionAuthors':list({e.get(W+'author') for e in d.iter() if e.tag in (W+'ins',W+'del')}),'numberedParagraphs':sum(p.find('./'+W+'pPr/'+W+'numPr') is not None for p in d.iter(W+'p')),'paragraphs':[text(p) for p in d.iter(W+'p')],'comments':comments,'footnotes':footnotes,'metadata':metadata}))`;
  const result = await run(python, ['-c', script, file]);
  return JSON.parse(result.stdout);
}

function reviewedNative() {
  const session = {...native, decisions:{}, comments:structuredClone(native.comments)};
  const review = compare(session.before, session.after);
  for (const segment of review.segments.filter(s => s.type === 'change')) {
    // Negotiate the payment and clause replacement, but retain confidentiality.
    session.decisions[segment.id] = segment.before.includes('including the existence') ? 'rejected' : 'accepted';
  }
  return session;
}

test('native Word revisions import both wording endpoints and independently reviewable clauses', () => {
  const expected = manifest.files['03-native-tracked.docx'];
  for (const clause of expected.beforeClauseParagraphs) assert.ok(norm(native.beforeConvertedText).includes(norm(clause)), clause);
  for (const clause of expected.afterClauseParagraphs) assert.ok(norm(native.convertedText).includes(norm(clause)), clause);
  assert.ok(native.before.includes('48,000') && native.before.includes('20 days'));
  assert.ok(native.after.includes('52,000') && native.after.includes('30 days'));
  assert.ok(!native.after.includes('may terminate'));
  assert.ok(!native.before.includes('publish an announcement'));
  assert.ok(native.before.includes('\\footnote{'));
  assert.ok(native.before.includes(manifest.sharedFootnotes[0].text));
  const review = compare(native.before, native.after);
  assert.ok(review.segments.filter(s => s.type === 'change').length >= 3, 'Numbered clauses cannot collapse into one review choice');
  assert.deepEqual(review.issues, {before:[], after:[]});
  assert.equal((native.before.match(/\\item\b/g) || []).length, 6);
  assert.equal((native.after.match(/\\item\b/g) || []).length, 6);
  assert.equal(native.originalRevisions.length, expected.trackedElementCount);
  assert.ok(native.originalRevisions.every(r => r.author === manifest.revisionAuthor && r.at === manifest.revisionDate));
});

test('Word comments retain author, date, wording, and unchanged or revised clause context', () => {
  assert.deepEqual(native.unmatchedComments, []);
  const review = compare(native.before, native.after);
  for (const expected of manifest.sharedComments) {
    const entry = Object.entries(native.comments).find(([, notes]) => notes.some(n => n.wordId === expected.id));
    assert.ok(entry, 'Every received comment must remain reachable');
    const note = entry[1].find(n => n.wordId === expected.id);
    assert.equal(note.text, expected.text);
    assert.equal(note.author, expected.author);
    assert.equal(note.at, expected.date);
    const segment = review.segments.find(s => s.id === entry[0]);
    assert.ok(segment.after.includes(expected.anchor || expected.anchorAfter));
    assert.equal(segment.type, expected.paragraphKey === 'scope' ? 'equal' : 'change');
  }
});

test('original DOCX bytes survive preservation and portable review restoration', async () => {
  const bytes = await readFile(path.join(fixtures, '03-native-tracked.docx'));
  assert.equal(native.attachment.hash, manifest.files['03-native-tracked.docx'].sha256);
  const preserved = await originalWord(native.attachment.hash);
  assert.deepEqual(Buffer.from(preserved.data, 'base64'), bytes);
  const portable = [{...native.attachment, role:'source', data:preserved.data}];
  const restored = await restoreWordFiles(JSON.parse(JSON.stringify(portable)));
  assert.equal(restored[0].hash, hash(bytes));
  assert.equal(restored[0].size, bytes.length);
  assert.equal(restored[0].data, undefined, 'Stored session metadata must not duplicate the full original');
  assert.deepEqual(Buffer.from((await originalWord(restored[0].hash)).data, 'base64'), bytes);
  await assert.rejects(restoreWordFiles([{...portable[0], hash:'0'.repeat(64)}]), /hash/i);
  await assert.rejects(restoreWordFiles([{...portable[0], size:1}]), /size/i);
  await assert.rejects(originalWord('../not-a-hash'), /reference/i);
});

test('mixed review exports genuine native Word changes and a clean counterpart with exact reviewed wording', async () => {
  const session = reviewedNative();
  const material = resolve(compare(session.before, session.after), session.decisions);
  assert.ok(material.includes('52,000') && material.includes('30 days'));
  assert.ok(material.includes('including the existence'));
  assert.ok(material.includes('publish an announcement'));
  assert.ok(!material.includes('may terminate'));
  for (const mode of ['clean', 'tracked']) {
    const exported = await exportWord(session, mode);
    const xml = await inspectDocx(exported.data, mode+'.docx');
    assert.equal(exported.sourceHash, hash(material));
    if (mode === 'clean') assert.equal(xml.insertions + xml.deletions, 0);
    else {
      assert.ok(xml.insertions > 0 && xml.deletions > 0, 'A redline must contain native Word revisions');
      assert.deepEqual(xml.revisionAuthors, ['Margin'], 'A newly generated comparison cannot impersonate the incoming reviewer');
    }
    assert.ok(xml.numberedParagraphs >= 6, 'Numbering must use w:numPr');
    assert.deepEqual(xml.footnotes.map(norm), manifest.sharedFootnotes.map(n => n.text));
    assert.equal(xml.comments.length, 2);
    for (const expected of manifest.sharedComments) {
      const note = xml.comments.find(n => n.text === expected.text);
      assert.ok(note);
      assert.equal(note.author, expected.author);
      assert.equal(note.at, expected.date);
    }
    assert.ok(Object.values(xml.metadata).every(value => value === ''), 'Outgoing document properties must be blank');
    const roundTrip = await importWord({data:exported.data, name:mode+'.docx'});
    assert.ok(norm(roundTrip.convertedText).includes('CHF 52,000 within 30 days'));
    assert.ok(norm(roundTrip.convertedText).includes('including the existence of this agreement'));
    assert.ok(norm(roundTrip.convertedText).includes('publish an announcement'));
    assert.ok(!norm(roundTrip.convertedText).includes('may terminate'));
    if (mode === 'tracked') {
      assert.ok(norm(roundTrip.beforeConvertedText).includes('CHF 48,000 within 20 days'));
      assert.ok(norm(roundTrip.beforeConvertedText).includes('may terminate'));
      assert.ok(!norm(roundTrip.beforeConvertedText).includes('publish an announcement'));
    } else assert.equal(roundTrip.before, roundTrip.after);
  }
});

test('Word export blocks unresolved decisions and unsafe comment loss', async () => {
  await assert.rejects(exportWord({...native, decisions:{}}, 'clean'), /Review every change/);
  const session = reviewedNative();
  await assert.rejects(exportWord({...session, comments:{unknown:[{text:'Must not disappear'}]}}, 'clean'), /comment.*match/i);
  await assert.rejects(exportWord({...session, conversion:{reports:[{unmatchedComments:[{text:'Retain me'}]}]}}, 'clean'), /safe passage anchor/i);
  const beforeText = 'Keep this clause.\n\nDelete this clause.\n\n';
  const afterText = 'Keep this clause.\n\n';
  const review = compare(beforeText, afterText);
  const deleted = review.segments.find(s => s.type === 'change');
  await assert.rejects(exportWord({before:beforeText, after:afterText, decisions:{[deleted.id]:'accepted'}, comments:{[deleted.id]:[{text:'A comment on the deleted clause', at:'2026-10-09T10:00:00Z'}]}}, 'clean'), /safely place|anchor|comment/i);
});

test('comparing two clean Word documents maps both received comment sets without losing unchanged comments', async () => {
  const baseline = await openFixture('01-baseline.docx');
  const proposed = await openFixture('02-clean-proposed.docx');
  assert.equal(baseline.before, baseline.after);
  assert.equal(proposed.before, proposed.after);
  const result = mapComparisonComments(baseline.after, proposed.after, [
    {side:'before', source:baseline.after, comments:baseline.comments},
    {side:'after', source:proposed.after, comments:proposed.comments}
  ]);
  assert.equal(notesOf(result).length, 2, 'Identical imported notes should be represented once');
  const review = compare(baseline.after, proposed.after);
  assert.ok(Object.keys(result.comments).some(id => review.segments.find(s => s.id === id)?.type === 'equal'));
  await assert.rejects(async () => mapComparisonComments(baseline.after, proposed.after, [{side:'before', source:'unrelated', comments:baseline.comments}]), /source.*match/);
});

test('paragraph-mark revisions preserve split and joined paragraph boundaries', async () => {
  const result = await openFixture('04-paragraph-boundaries.docx');
  const expected = manifest.files['04-paragraph-boundaries.docx'];
  const paragraphs = side => result.blocks[side].map(b => norm(b.text)).filter(Boolean);
  for (const side of ['before', 'after']) {
    for (const paragraph of expected[side+'Paragraphs']) assert.ok(paragraphs(side).includes(norm(paragraph)), `${side}: ${paragraph}`);
  }
  assert.notEqual(result.before, result.after, 'Paragraph-mark revisions cannot become a no-op');
});

test('complex Word constructs are preserved or explicitly reported before review', async () => {
  const result = await openFixture('05-special-constructs.docx');
  const codes = new Set(result.warnings.map(w => w.code));
  for (const code of ['headers_footers', 'content_controls', 'fields', 'tracked_table_rows', 'text_boxes']) assert.ok(codes.has(code), code);
  assert.ok(result.counts.tables >= 1);
  assert.ok(result.after.includes('north loading bay'));
  assert.ok(result.after.includes('15 October 2026'));
  assert.equal(result.attachment.hash, manifest.files['05-special-constructs.docx'].sha256);
});

test('literal pasted text cannot inject LaTeX commands and retains special characters in Word', async () => {
  const literal = 'Pay 100% & CHF 5_000 #1 $2 {net}.\n\nLiteral \\input{/private/secret} ~ ^ text.';
  const result = importText(literal, 'Literal input');
  assert.equal(result.before, result.after);
  assert.ok(result.after.includes('100\\% \\& CHF 5\\_000'));
  assert.ok(result.after.includes('\\textbackslash{}input\\{/private/secret\\}'));
  assert.ok(!result.after.includes('\\input{'));
  const exported = await exportWord({...result, comments:{}, decisions:{}}, 'clean');
  const xml = await inspectDocx(exported.data, 'literal.docx');
  assert.ok(xml.paragraphs.includes('Pay 100% & CHF 5_000 #1 $2 {net}.'));
  assert.ok(xml.paragraphs.includes('Literal \\input{/private/secret} ~ ^ text.'));
  assert.throws(() => importText('  '), /Paste some/);
});

test('comment coordinates disambiguate duplicate paragraphs following astral Unicode', () => {
  const source = 'Heading 😀.\n\nSame paragraph.\n\nSame paragraph.\n\n';
  const fragment = 'Same paragraph.\n\n';
  const start = source.lastIndexOf(fragment);
  const review = compare(source, source);
  const result = anchorComments(review, [{id:'7', text:'Second occurrence', author:'Reviewer', afterAnchor:{start, end:start+fragment.length, source:fragment}}]);
  assert.deepEqual(result.unmatched, []);
  assert.equal(Object.keys(result.comments)[0], review.segments.at(-1).id);
});

test('comments on unchanged passages survive a prose edit elsewhere', () => {
  const source = 'Unchanged paragraph.\n\nEditable paragraph.\n\n';
  const review = compare(source, source);
  const session = {before:source, after:source, decisions:{}, comments:{[review.segments[0].id]:[{text:'Keep this note', author:'Original reviewer', origin:'word'}]}};
  const next = reviseParagraph(review, session, review.segments[1].id, 'Edited paragraph.', compare);
  const target = compare(next.before, next.after).segments.find(s => s.after.includes('Unchanged'));
  assert.deepEqual(next.comments[target.id], session.comments[review.segments[0].id]);
});

test('maximum-size portable bytes are handled without stack overflow and tampering is detected', async () => {
  const bytes = Buffer.alloc(MAX_WORD_BYTES, 65);
  bytes.write('PK', 0, 'ascii');
  const record = {hash:hash(bytes), size:bytes.length, name:'transport-limit.docx', data:bytes.toString('base64')};
  const restored = await restoreWordFiles([record]);
  assert.equal(restored[0].size, MAX_WORD_BYTES);
  assert.deepEqual(Buffer.from((await originalWord(record.hash)).data, 'base64'), bytes);
  await writeFile(path.join(process.env.FOLIO_DATA_DIR, 'originals', record.hash+'.docx'), Buffer.from('tampered'));
  await assert.rejects(originalWord(record.hash), /integrity check/);
  await assert.rejects(restoreWordFiles([record]), /integrity check/);
  await assert.rejects(restoreWordFiles([{...record, data:Buffer.concat([bytes, Buffer.from('x')]).toString('base64')}]), /4 MB|supported Word/);
});
