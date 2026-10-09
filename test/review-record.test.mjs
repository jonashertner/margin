import test from 'node:test';
import assert from 'node:assert/strict';
import {compare} from '../lib/review.mjs';
import {buildReviewRecord,reviewContext} from '../public/review-record.js';
test('machine record keeps exact Unicode source, pending decisions and source positions',()=>{
  const before='🖋️ Pay 10 days.\n\nKeep this.',after='🖋️ Pay 30 days.\n\nKeep this.';
  const review=compare(before,after),change=review.segments.find(s=>s.type==='change');
  const session={title:'Agreement',before,after,decisions:{},comments:{[change.id]:[{text:'Who can confirm the timing?',at:'2026-10-09'}]},context:{openQuestions:'Timing is unconfirmed.'}};
  const result=buildReviewRecord(session,review);
  assert.equal(result.status.remaining,1);assert.equal(result.status.legalApproval,'not_recorded');
  assert.equal(result.context.openQuestions,'Timing is unconfirmed.');
  for(const p of result.passages){assert.equal(before.slice(p.source.beforeRange.start,p.source.beforeRange.end),p.source.before);assert.equal(after.slice(p.source.afterRange.start,p.source.afterRange.end),p.source.proposed);}
  assert.equal(result.passages.find(p=>p.id===change.id).comments[0].text,'Who can confirm the timing?');
  session.decisions[change.id]='accepted';
  const decided=buildReviewRecord(session,review);assert.equal(decided.status.wordingReview,'decided');assert.equal(decided.status.legalApproval,'not_recorded');
});
test('review context is bounded and does not promote arbitrary fields to authority',()=>{
  assert.deepEqual(reviewContext({approved:true}),{purpose:'',openQuestions:'',nextStep:''});
  assert.throws(()=>reviewContext({purpose:42}),/could not be read/);
  assert.throws(()=>reviewContext({nextStep:'x'.repeat(4001)}),/4,000/);
});
