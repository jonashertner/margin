export function reviewContext(value={}) {
  const result={};
  for(const key of ['purpose','openQuestions','nextStep']) {
    if(value?.[key]!==undefined&&typeof value[key]!=='string')throw Error('The review note could not be read.');
    const text=value?.[key]||'';
    if(text.length>4000)throw Error('Keep each part of the review note under 4,000 characters.');
    result[key]=text;
  }
  return result;
}

// Human and machine readers receive the same wording, decisions and uncertainties.
// Source positions are UTF-16 code units, matching JavaScript string slicing.
export function buildReviewRecord(session,review,{at=new Date().toISOString()}={}) {
  let beforeOffset=0,afterOffset=0;
  const passages=review.segments.map(segment=>{
    const beforeRange={start:beforeOffset,end:beforeOffset+segment.before.length};
    const afterRange={start:afterOffset,end:afterOffset+segment.after.length};
    beforeOffset=beforeRange.end;afterOffset=afterRange.end;
    const decision=segment.type==='equal'?'unchanged':session.decisions?.[segment.id]||'pending';
    return {
      id:segment.id,kind:segment.type,decision,
      source:{before:segment.before,proposed:segment.after,beforeRange,afterRange},
      reading:{before:(segment.words||[]).filter(w=>!w.added).map(w=>w.value).join(''),proposed:(segment.words||[]).filter(w=>!w.removed).map(w=>w.value).join('')},
      comments:session.comments?.[segment.id]||[],
      prompts:segment.signals||[],
    };
  });
  const remaining=passages.filter(p=>p.decision==='pending').length;
  return {
    format:'margin-review-record',version:1,generatedAt:at,
    title:session.title,workspaceId:session.workspaceId,
    context:reviewContext(session.context),
    status:{remaining,wordingReview:remaining?'in_progress':'decided',legalApproval:'not_recorded',identities:'not_authenticated'},
    source:{format:'latex',offsetUnit:'utf16_code_units',before:session.before,proposed:session.after,beforeHash:review.beforeHash,proposedHash:review.afterHash,incomingProposal:session.authoring?.originalAfter??session.after},
    originals:(session.attachments||[]).map(({hash,name,size,role})=>({sha256:hash,name,size,role})),
    conversion:session.conversion||null,
    passages,
    history:{authenticated:false,events:session.audit||[]},
    limits:['Reading text simplifies LaTeX; exact source is controlling for this comparison.','Prompts are heuristics, not legal findings.','A wording decision is not approval, signature or authority to act.','Missing comments or questions do not establish that there are no concerns.','A Word snapshot is separate from the live Word document.'],
  };
}
