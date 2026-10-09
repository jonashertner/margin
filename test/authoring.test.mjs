import test from 'node:test';
import assert from 'node:assert/strict';
import {editableParagraph,reviseParagraph} from '../lib/authoring.mjs';
import {compare,resolve} from '../lib/review.mjs';
const simple=raw=>({before:raw,after:raw,type:'equal'});
const open=(before,after=before)=>({before,after,decisions:{},comments:{},audit:[],snapshots:[{label:'Preserve me'}]});
const revise=(session,id,text,fn=compare)=>reviseParagraph(compare(session.before,session.after),session,id,text,fn);

test('single-document editing preserves wrappers, CRLF, BOM, and untouched bytes',()=>{
  const source='\uFEFF\\documentclass{article}\r\n\\begin{document}\r\n\r\n\\noindent  Payment is due. \r\n\r\nAnother paragraph.\r\n\\end{document}';
  const session=open(source),review=compare(source,source),target=review.segments.find(s=>s.after.includes('Payment'));
  const next=reviseParagraph(review,session,target.id,'Payment is due in 30 days.',compare);
  assert.equal(next.before,source);
  assert.equal(next.after,source.replace('Payment is due.','Payment is due in 30 days.'));
  assert.equal(next.authoring.originalAfter,source);
  assert.deepEqual(next.snapshots,session.snapshots);
  assert.equal(resolve(compare(next.before,next.after),next.decisions),next.after);
  assert.equal(session.after,source);
});
test('literal TeX specials are escaped and remain editable',()=>{
  const session=open('Pay now.\n\n'); const target=compare(session.before,session.after).segments[0];
  const text='100% & CHF 5_000 #1 $2 {net} \\ ~ ^';
  const next=reviseParagraph(compare(session.before,session.after),session,target.id,text,compare);
  assert.equal(editableParagraph(compare(next.before,next.after).segments[0]).text,text);
  assert.ok(next.after.includes('100\\% \\& CHF 5\\_000'));
});
test('formatted macros, citations, comments, maths, and multiple paragraphs stay read-only',()=>{
  for(const raw of ['\\textbf{Pay now.}\n','See \\ref{claim}.\n','Amount $x$.\n','Text. % comment\n','First.\n\nSecond.\n','Due~today.\n'])assert.equal(editableParagraph(simple(raw)).editable,false,raw);
});
test('no-op editing is byte-identical with no audit event',()=>{
  const session=open('\\noindent Pay 10\\%.\r\n\r\n'),review=compare(session.before,session.after);
  const result=reviseParagraph(review,session,review.segments[0].id,'Pay 10%.',compare);
  assert.deepEqual(result,session);
});
test('unrelated decisions and all notes survive editing a selected change',()=>{
  const session=open('First old.\n\nMiddle unchanged.\n\nLast old.\n','First new.\n\nMiddle unchanged.\n\nLast new.\n');
  const review=compare(session.before,session.after),changes=review.segments.filter(s=>s.type==='change');
  session.decisions[changes[0].id]='rejected'; session.decisions[changes[1].id]='accepted';
  session.comments[changes[0].id]=[{text:'Preserve first note',at:'2026-10-09'}];
  session.comments[changes[1].id]=[{text:'Preserve last note',at:'2026-10-09'}];
  const original=structuredClone(session);
  const next=reviseParagraph(review,session,changes[0].id,'First negotiated.',compare);
  const newChanges=compare(next.before,next.after).segments.filter(s=>s.type==='change');
  assert.equal(next.decisions[newChanges[0].id],'accepted');
  assert.equal(next.decisions[newChanges[1].id],'accepted');
  assert.deepEqual(next.comments[newChanges[0].id],session.comments[changes[0].id]);
  assert.deepEqual(next.comments[newChanges[1].id],session.comments[changes[1].id]);
  assert.equal(next.audit.at(-1).previousDecision,'rejected');
  assert.deepEqual(session,original);
});
test('editing an unchanged paragraph preserves other decisions',()=>{
  const session=open('First old.\n\nMiddle unchanged.\n\nLast old.\n','First new.\n\nMiddle unchanged.\n\nLast new.\n');
  const review=compare(session.before,session.after),changes=review.segments.filter(s=>s.type==='change');
  session.decisions[changes[0].id]='rejected'; session.decisions[changes[1].id]='accepted';
  const target=review.segments.find(s=>s.after.includes('Middle'));
  const next=reviseParagraph(review,session,target.id,'Middle revised.',compare);
  assert.equal(resolve(compare(next.before,next.after),next.decisions),'First old.\n\nMiddle revised.\n\nLast new.\n');
});
test('ambiguous duplicate remaps reject before mutating session',()=>{
  const session=open('Old.\n','New.\n'),review=compare(session.before,session.after),original=structuredClone(session);
  const ambiguous=(before,after)=>{const r=compare(before,after);r.segments[0].before='';r.segments[0].after='';r.segments.push({...r.segments[0]});return r;};
  assert.throws(()=>reviseParagraph(review,session,review.segments[0].id,'Edited.',ambiguous),/regroup/);
  assert.deepEqual(session,original);
});
test('duplicate sibling notes follow source coordinates when digest occurrences shift',()=>{
  const session=open('Same.\n\nSame.\n\nTail.\n','New.\n\nNew.\n\nTail.\n');
  const review=compare(session.before,session.after),changes=review.segments.filter(s=>s.type==='change');
  session.decisions[changes[1].id]='rejected';
  session.comments[changes[1].id]=[{text:'Note for the second occurrence',at:'2026-10-09'}];
  const next=reviseParagraph(review,session,changes[0].id,'Edited.',compare);
  const nextChanges=compare(next.before,next.after).segments.filter(s=>s.type==='change');
  assert.equal(next.decisions[nextChanges[1].id],'rejected');
  assert.equal(next.comments[nextChanges[1].id][0].text,'Note for the second occurrence');
  assert.equal(next.comments[nextChanges[0].id],undefined);
});
test('a edit that makes duplicate paragraphs realign is refused without mutation',()=>{
  const session=open('First.\n\nSecond.\n\n'),review=compare(session.before,session.after),original=structuredClone(session);
  assert.throws(()=>reviseParagraph(review,session,review.segments[0].id,'Second.',compare),/regroup/);
  assert.deepEqual(session,original);
});
test('new single line breaks use the source paragraph CRLF convention',()=>{
  const session=open('Pay now.\r\n\r\n'),review=compare(session.before,session.after);
  const next=reviseParagraph(review,session,review.segments[0].id,'Pay\nnow.',compare);
  assert.equal(next.after,'Pay\r\nnow.\r\n\r\n');
});
test('stale sources, blank text, paragraph breaks, and controls are rejected',()=>{
  const session=open('Pay now.\n'),review=compare(session.before,session.after),id=review.segments[0].id;
  for(const text of ['','A\n\nB','A\u0000B'])assert.throws(()=>reviseParagraph(review,session,id,text,compare));
  assert.throws(()=>reviseParagraph(review,{...session,after:'Changed.'},id,'New.',compare),/document changed/);
});
test('repeated edits retain the incoming proposal without duplicating full paragraphs in history',()=>{
  const source='a'.repeat(249990)+'.\n\n';let session=open(source);
  for(let i=0;i<3;i++){const review=compare(session.before,session.after);session=reviseParagraph(review,session,review.segments[0].id,'a'.repeat(249990)+String(i)+'.',compare);}
  assert.equal(session.authoring.originalAfter,source);
  assert.ok(Buffer.byteLength(JSON.stringify(session))<800000);
  assert.equal(session.authoring.edits.length,3);
  assert.ok(session.authoring.edits.every(e=>e.previousHash.length===64&&e.replacementHash.length===64));
});
