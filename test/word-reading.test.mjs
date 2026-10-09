import test from 'node:test';
import assert from 'node:assert/strict';
import {compare,numberedParagraph,plain,resolve} from '../lib/review.mjs';
import {editableParagraph,reviseParagraph} from '../lib/authoring.mjs';
const clause=(number,text,{eol='\n',legacyOrder=false,tight=false}={})=>{
 const counter=number===1?'':`\\setcounter{enumi}{${number-1}}${eol}`;
 const label=`\\def\\labelenumi{\\arabic{enumi}.}${eol}`;
 return `\\begin{enumerate}${eol}${legacyOrder?label+counter:counter+label}${tight?`\\tightlist${eol}`:''}\\item${eol}  ${text}${eol}\\end{enumerate}${eol}${eol}`;
};
const session=(before,after=before)=>({before,after,decisions:{},comments:{},audit:[],snapshots:[{label:'Incoming Word'}]});

test('imported numbered clauses display real numbers without TeX machinery',()=>{
 for(const legacyOrder of [false,true])for(const eol of ['\n','\r\n'])for(const number of [1,2,19]){
  const raw=clause(number,'Pay CHF 52,000 within 30 days.',{legacyOrder,eol,tight:true});
  assert.equal(plain(raw),`${number}. Pay CHF 52,000 within 30 days.`);
  const parsed=numberedParagraph(raw);assert.equal(parsed.number,number);
  assert.equal(parsed.prefix+parsed.body+parsed.suffix,raw);
 }
 assert.equal(plain(clause(2,'Payment.')+clause(3,'Interest.')),'2. Payment.\n\n3. Interest.');
});
test('numbered display keeps footnote wording visible and source unchanged',()=>{
 const raw=clause(3,'Interest at 4\\%.\\footnote{An explanatory note.}');
 assert.equal(plain(raw),'3. Interest at 4%. [Footnote: An explanatory note.]');
 assert.equal(editableParagraph({before:raw,after:raw}).editable,false);
});
test('ordinary clause changes keep legal signals without spurious structure attention',()=>{
 const change=compare(clause(2,'Pay CHF 48,000 within 20 days.'),clause(2,'Pay CHF 52,000 within 30 days.')).segments[0];
 assert.deepEqual(change.signals.map(s=>s.label),['Amount','Timing']);
 assert.equal(change.words.map(w=>w.value).join('').includes('setcounter'),false);
 const renumbered=compare(clause(2,'Pay now.'),clause(3,'Pay now.')).segments[0];
 assert.ok(renumbered.signals.some(s=>s.label==='Structure'));
 assert.ok(renumbered.words.some(w=>w.added&&w.value.includes('3')));
});
test('safe numbered editing retains exact wrapper, line endings and snapshots',()=>{
 const raw=clause(4,'Pay 10\\% now.',{legacyOrder:true,eol:'\r\n',tight:true});
 const state=session(raw),review=compare(raw,raw),target=review.segments[0];
 const editor=editableParagraph(target);assert.equal(editor.editable,true);assert.equal(editor.text,'Pay 10% now.');
 assert.equal(editor.prefix+'Pay 10\\% now.'+editor.suffix,raw);
 const next=reviseParagraph(review,state,target.id,'Pay 15% & CHF 52_000 now.',compare);
 assert.equal(next.after,raw.replace('Pay 10\\% now.','Pay 15\\% \\& CHF 52\\_000 now.'));
 assert.equal(next.before,raw);assert.equal(next.authoring.originalAfter,raw);assert.deepEqual(next.snapshots,state.snapshots);
 assert.equal(resolve(compare(next.before,next.after),next.decisions),next.after);
 assert.deepEqual(reviseParagraph(review,state,target.id,'Pay 10% now.',compare),state);
});
test('numbered editing retains unrelated decisions and notes including edited clause note',()=>{
 const before=clause(1,'Scope old.')+clause(2,'Payment old.')+clause(3,'Unchanged.');
 const after=clause(1,'Scope proposed.')+clause(2,'Payment proposed.')+clause(3,'Unchanged.');
 const state=session(before,after),review=compare(before,after),[scope,payment,unchanged]=review.segments;
 state.decisions[payment.id]='rejected';state.comments[scope.id]=[{text:'Keep this note',author:'Reviewer'}];state.comments[unchanged.id]=[{text:'Unchanged clause note'}];
 const next=reviseParagraph(review,state,scope.id,'Scope negotiated.',compare),nextReview=compare(next.before,next.after);
 assert.equal(resolve(nextReview,next.decisions),clause(1,'Scope negotiated.')+clause(2,'Payment old.')+clause(3,'Unchanged.'));
 assert.equal(next.comments[nextReview.segments[0].id][0].text,'Keep this note');
 assert.equal(next.comments[unchanged.id][0].text,'Unchanged clause note');
});
test('unsupported numbering wrappers stay read-only instead of being guessed',()=>{
 const base=clause(2,'Payment.');
 for(const raw of [base.replace('\\arabic','\\roman'),base.replace('\\item\n','\\item[Custom]\n'),base.replace('\\item\n','\\input{untrusted}\n\\item\n'),base.replace('Payment.','First.\n\\item Second.'),base.replace('Payment.',clause(1,'Nested.')),base.replace('{enumi}{1}','{enumi}{-1}'),base.replace('\\def\\labelenumi{\\arabic{enumi}.}','')]){
  assert.equal(numberedParagraph(raw),null,raw);assert.equal(editableParagraph({before:raw,after:raw}).editable,false,raw);
 }
});
test('numbered editor refuses references, macros, comments, maths and multiple paragraphs',()=>{
 for(const body of ['See \\ref{claim}.','See \\cite{case}.','\\textbf{Important.}','A. % retain comment','Amount $x$.','First.\n\nSecond.']){
  const raw=clause(7,body);assert.equal(editableParagraph({before:raw,after:raw}).editable,false,body);
 }
});
test('editing a rejected numbered clause starts from its original exact wrapper',()=>{
 const before=clause(2,'Original wording.',{legacyOrder:true}),after=clause(2,'Proposed wording.');
 const state=session(before,after),review=compare(before,after),target=review.segments[0];state.decisions[target.id]='rejected';
 const next=reviseParagraph(review,state,target.id,'Negotiated wording.',compare);
 assert.equal(next.after,before.replace('Original wording.','Negotiated wording.'));
 assert.equal(next.audit.at(-1).previousDecision,'rejected');
});
