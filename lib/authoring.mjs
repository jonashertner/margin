import { createHash } from 'node:crypto';
import { numberedParagraph } from './review.mjs';
const sha = text => createHash('sha256').update(text).digest('hex');
const copy = value => structuredClone(value);
const escapeMap = {'\\':'\\textbackslash{}','%':'\\%','&':'\\&','#':'\\#','$':'\\$','_':'\\_','{':'\\{','}':'\\}','^':'\\textasciicircum{}','~':'\\textasciitilde{}'};
const fail = reason => ({ editable:false, reason });

// Deliberately conservative: prose, escaped literal characters, and one optional
// leading \\noindent. Never call the lossy display/plain() conversion here.
export function editableParagraph(segment, which='after') {
  if (!segment || !['before','after'].includes(which)) return fail('Choose a paragraph.');
  if (segment.preamble || segment.end || segment.heading) return fail('This is document structure rather than a prose paragraph.');
  const raw=segment[which];
  if (typeof raw!=='string' || !raw.trim()) return fail('This version has no paragraph to edit.');
  const clause=numberedParagraph(raw), prose=clause?.body??raw;
  const prefixMatch=prose.match(/^[ \t\r\n]*(?:\\noindent(?=[ \t\r\n]|$)[ \t\r\n]*)?/);
  const prefix=(clause?.prefix??'')+prefixMatch[0], remaining=prose.slice(prefixMatch[0].length);
  const trailing=remaining.match(/[ \t\r\n]*$/)[0], body=remaining.slice(0,remaining.length-trailing.length);
  const suffix=trailing+(clause?.suffix??'');
  if (!body || /\r?\n[ \t]*\r?\n/.test(body)) return fail('This selection contains more than one paragraph.');
  if (/\r(?!\n)/.test(body) || /\r\n/.test(body) && /(?<!\r)\n/.test(body)) return fail('This paragraph mixes line endings; edit its source to preserve them.');
  let text='';
  for (let i=0;i<body.length;) {
    const c=body[i];
    if(c==='\\') {
      const token=body.slice(i).match(/^\\(?:[%&#_${}]|textbackslash\{\}|textasciitilde\{\}|textasciicircum\{\})/);
      if(!token)return fail('This paragraph contains LaTeX formatting or references. Its source remains available.');
      const value=token[0];
      text+=value==='\\textbackslash{}'?'\\':value==='\\textasciitilde{}'?'~':value==='\\textasciicircum{}'?'^':value.slice(1);
      i+=value.length;
    } else {
      if(/[%&#_${}^~]/.test(c))return fail('This paragraph contains LaTeX structure. Its source remains available.');
      text+=c;i++;
    }
  }
  const eol=raw.includes('\r\n')?'\r\n':'\n';
  return {editable:true,text:text.replace(/\r\n/g,'\n'),prefix,suffix,eol,raw,which};
}

function coordinates(review) {
  let oldOffset=0,newOffset=0;
  return review.segments.map(segment=>{const item={segment,oldOffset,newOffset};oldOffset+=segment.before.length;newOffset+=segment.after.length;return item;});
}

// compareFn is the application's existing compare(before, after). Recomputing is
// required to detect regrouping; IDs alone are insufficient for safe remapping.
export function reviseParagraph(review,session,id,newText,compareFn) {
  if(typeof compareFn!=='function')throw Error('A comparison function is required to preserve review decisions.');
  if(typeof newText!=='string')throw Error('Enter the revised wording.');
  if(!review?.segments || review.segments.map(s=>s.before).join('')!==session.before || review.segments.map(s=>s.after).join('')!==session.after)throw Error('The document changed. Reopen this paragraph before editing it.');
  const old=coordinates(review), targets=old.filter(x=>x.segment.id===id);
  if(targets.length!==1)throw Error('The paragraph could not be identified uniquely.');
  const target=targets[0], segment=target.segment;
  const which=session.decisions?.[id]==='rejected'?'before':'after';
  const editor=editableParagraph(segment,which);
  if(!editor.editable)throw Error(editor.reason);
  const text=newText.replace(/\r\n/g,'\n');
  if(text===editor.text)return copy(session);
  if(!text.trim())throw Error('Use Keep original or source editing to remove a paragraph.');
  if(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f\r]/.test(text)||text.toWellFormed()!==text)throw Error('The wording contains an unsupported character.');
  if(/\n[ \t]*\n/.test(text))throw Error('Edit one paragraph at a time.');
  const body=[...text].map(c=>escapeMap[c]||c).join('').replace(/\n/g,editor.eol);
  const replacement=editor.prefix+body+editor.suffix;
  if(segment.type==='change'&&replacement===segment.before)throw Error('This matches the earlier version. Choose Keep original to retain this review and its notes.');
  const revisedAfter=session.after.slice(0,target.newOffset)+replacement+session.after.slice(target.newOffset+segment.after.length);
  const nextReview=compareFn(session.before,revisedAfter), next=coordinates(nextReview);
  const delta=replacement.length-segment.after.length;
  function destination(item,edited=false) {
    const expectedAfter=edited?replacement:item.segment.after;
    const expectedOffset=item.newOffset+(item.newOffset>=target.newOffset+segment.after.length&&item!==target?delta:0);
    const matches=next.filter(x=>x.oldOffset===item.oldOffset && x.newOffset===(edited?target.newOffset:expectedOffset) && x.segment.before===item.segment.before && x.segment.after===expectedAfter);
    if(matches.length!==1)throw Error('This edit would regroup existing changes. The review has been preserved; edit this passage in its source instead.');
    return matches[0].segment;
  }
  const revisedSegment=destination(target,true);
  if(revisedSegment.type!=='change')throw Error('This edit no longer creates a distinct reviewable change.');
  const ids=new Map(old.map(x=>[x.segment.id,x]));
  if(ids.size!==old.length)throw Error('Duplicate paragraph identifiers prevent a safe edit.');
  const decisions={},comments={};
  for(const [oldId,decision] of Object.entries(session.decisions||{})) {
    if(!['accepted','rejected'].includes(decision)||!ids.has(oldId)||ids.get(oldId).segment.type!=='change')throw Error('An existing review decision could not be preserved.');
    if(oldId===id)continue;
    const mapped=destination(ids.get(oldId));
    if(mapped.type!=='change')throw Error('An existing review decision could not be preserved.');
    decisions[mapped.id]=decision;
  }
  for(const [oldId,notes] of Object.entries(session.comments||{})) {
    if(!ids.has(oldId)||!Array.isArray(notes))throw Error('An existing review note could not be preserved.');
    const mapped=oldId===id?revisedSegment:destination(ids.get(oldId));
    comments[mapped.id]=copy(notes);
  }
  // Saving a direct wording edit is an explicit choice of that wording. Preserve
  // any previous target decision in the audit; unrelated decisions are retained.
  decisions[revisedSegment.id]='accepted';
  const at=new Date().toISOString();
  const entry={action:'edited',id:revisedSegment.id,previousId:id,previousDecision:session.decisions?.[id]||'pending',at,actor:'Local reviewer',beforeHash:sha(session.after),afterHash:sha(revisedAfter)};
  const nextSession={...copy(session),after:revisedAfter,beforeHash:nextReview.beforeHash,afterHash:nextReview.afterHash,decisions,comments,audit:[...copy(session.audit||[]),entry]};
  nextSession.authoring={...copy(session.authoring||{}),originalAfter:session.authoring?.originalAfter??session.after,edits:[...copy(session.authoring?.edits||[]),{at,offset:target.newOffset,previousHash:sha(segment.after),replacementHash:sha(replacement),previousDecision:entry.previousDecision}]};
  return nextSession;
}
