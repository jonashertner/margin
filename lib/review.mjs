import { diffArrays, diffWordsWithSpace } from 'diff';
import { createHash } from 'node:crypto';
export const hash = value => createHash('sha256').update(value).digest('hex');
function wordDiff(before,after){const tokens=s=>s.match(/\s+|\d+(?:[.,’']\d+)*%?|[\p{L}\p{M}_]+|[^\s]/gu)||[];return (diffArrays(tokens(before),tokens(after),{timeout:1200})||[{removed:true,value:tokens(before)},{added:true,value:tokens(after)}]).map(p=>({...p,value:p.value.join('')}));}
const escaped = (s,i) => { let n=0; while(i>0 && s[--i]==='\\') n++; return n%2===1; };
function noComments(source) {
 let out='',comment=false;
 for(let i=0;i<source.length;i++){const c=source[i];if(c==='\n')comment=false;if(c==='%'&&!escaped(source,i))comment=true;if(!comment)out+=c;}
 return out;
}
function opaque(source) {
 return source.replace(/\\begin\{(verbatim\*?|lstlisting|minted)\}[\s\S]*?\\end\{\1\}/g,m=>m.replace(/[^\r\n]/g,' ')).replace(/\\verb\*?([^\w\s])[^\n]*?\1/g,m=>' '.repeat(m.length));
}
export function structure(source) {
 let depth=0,stack=[],issues=[]; const clean=noComments(opaque(source));
 for(let i=0;i<clean.length;i++){if(escaped(clean,i))continue;if(clean[i]==='{')depth++;if(clean[i]==='}')depth--;if(depth<0){issues.push('An opening brace is missing.');depth=0;}}
 if(depth)issues.push('An opening brace has no matching closing brace.');
 for(const m of clean.matchAll(/\\(begin|end)\{([^}]+)\}/g)){if(m[1]==='begin')stack.push(m[2]);else if(stack.pop()!==m[2])issues.push(`Unpaired ${m[2]} environment.`);}
 if(stack.length)issues.push(`Unclosed environment: ${stack.join(', ')}.`);
 return [...new Set(issues)];
}
// Preserve all source bytes after UTF-8 decoding; split only at balanced boundaries.
export function units(source) {
 const lines=source.match(/[^\n]*\n|[^\n]+$/g)||[];let blocks=[],buffer='',depth=0,env=[],literal=null;
 for(const line of lines){
  buffer+=line;let clean=noComments(line);
  if(literal){if(clean.includes('\\end{'+literal+'}'))literal=null;continue;}
  const begin=clean.match(/\\begin\{(verbatim\*?|lstlisting|minted)\}/);
  if(begin&&!clean.includes('\\end{'+begin[1]+'}')){literal=begin[1];continue;}
  clean=opaque(clean);
  for(let i=0;i<clean.length;i++){if(!escaped(clean,i)){if(clean[i]==='{')depth++;if(clean[i]==='}')depth--;}}
  for(const m of clean.matchAll(/\\(begin|end)\{([^}]+)\}/g)){if(m[2]==='document')continue;if(m[1]==='begin')env.push(m[2]);else env.pop();}
  if((/^\s*$/.test(line)&&depth===0&&env.length===0)||(/\\(?:begin|end)\{document\}/.test(clean)&&depth===0&&env.length===0)){blocks.push(buffer);buffer='';}
 }
 if(buffer)blocks.push(buffer);return blocks;
}
// Recognize only the bounded one-item decimal wrapper emitted by our Word
// importer. Keep its bytes so authoring can replace the prose slice alone.
export function numberedParagraph(source) {
  if(typeof source!=='string')return null;
  const opening=source.match(/^[ \t\r\n]*\\begin\{enumerate\}[ \t\r\n]*/);
  const closing=source.match(/[ \t\r\n]*\\end\{enumerate\}[ \t\r\n]*$/);
  if(!opening||!closing||opening[0].length>=closing.index)return null;
  let offset=opening[0].length,counter=0,hasCounter=false,hasLabel=false,hasTightlist=false;
  while(offset<closing.index){
    const rest=source.slice(offset,closing.index);let token;
    if((token=rest.match(/^\\setcounter\{enumi\}\{(\d+)\}[ \t\r\n]*/))){
      if(hasCounter||!Number.isSafeInteger(Number(token[1]))||Number(token[1])>2147483646)return null;
      hasCounter=true;counter=Number(token[1]);
    }else if((token=rest.match(/^\\def\\labelenumi\{\\arabic\{enumi\}\.\}[ \t\r\n]*/))){
      if(hasLabel)return null;hasLabel=true;
    }else if((token=rest.match(/^\\tightlist\b[ \t\r\n]*/))){
      if(hasTightlist)return null;hasTightlist=true;
    }else break;
    offset+=token[0].length;
  }
  const item=source.slice(offset,closing.index).match(/^\\item(?=[ \t\r\n])[ \t\r\n]*/);
  if(!hasLabel||!item)return null;
  offset+=item[0].length;
  const body=source.slice(offset,closing.index);
  if(!body.trim()||[...body.matchAll(/\\(?:begin|end|item)\b/g)].some(m=>!escaped(body,m.index)))return null;
  return {body,prefix:source.slice(0,offset),suffix:source.slice(closing.index),number:counter+1,label:`${counter+1}.`};
}
export function plain(source) {
  source=source.replace(/\\begin\{enumerate\}[\s\S]*?\\end\{enumerate\}/g,raw=>{
    const clause=numberedParagraph(raw);return clause?`${clause.label} ${clause.body}`:raw;
  });
  return noComments(source).replace(/\\(?:noindent|maketitle)\b/g,'')
    .replace(/\\(?:documentclass|usepackage|geometry|hypersetup)(?:\[[^\]]*\])?\{[^}]*\}/g,'')
    .replace(/\\(?:begin|end)\{[^}]*\}(?:\[[^\]]*\])?/g,'')
    .replace(/\\(?:label)\{[^}]*\}/g,'')
    .replace(/\\(?:section|subsection|subsubsection|title|textbf|textit|emph|underline|textrm)\*?(?:\[[^\]]*\])?\{([^{}]*)\}/g,'$1')
    .replace(/\\footnote\{([^{}]*)\}/g,' [Footnote: $1]')
    .replace(/\\(?:ref|eqref|autoref|cite|parencite)\{([^}]+)\}/g,'[ref: $1]')
    .replace(/\\item(?:\[([^\]]+)\])?/g,'• $1 ')
    .replace(/\\(?:vspace|hspace)\*?\{[^}]*\}/g,' ').replace(/\\hrulefill/g,'________________').replace(/\\hfill/g,'     ').replace(/\\(?:maketitle|newpage|clearpage|noindent|medskip|bigskip|smallskip|centering|today)\b/g,'')
    .replace(/\\([%&#_$])/g,'$1').replace(/\\\\/g,'\n').replace(/~/g,' ')
    .replace(/\{([^{}]*)\}/g,'$1').trim();
}
export function signals(before,after) {
 const delta=diffWordsWithSpace(plain(before),plain(after)).filter(p=>p.added||p.removed).map(p=>p.value).join(' ');
 const all=plain(before+' '+after); const rawDelta=diffWordsWithSpace(before,after).filter(p=>p.added||p.removed).map(p=>p.value).join(' '); let items=[];
 const add=(label,detail)=>items.push({label,detail});
 if(/(?:CHF|EUR|USD|GBP|€|\$|£)|\b\d[\d'’.,]*\s*%/.test(all)&&/\d|CHF|EUR|USD|GBP|€|\$|£/.test(delta))add('Amount','Check the amount, currency and any related calculations.');
 if((/\b\d{1,4}[./-]\d{1,2}[./-]\d{1,4}\b/.test(all)&&/\d/.test(delta))||/\b(?:days?|months?|years?|weeks?|January|February|March|April|May|June|July|August|September|October|November|December|Tage?|Frist|Monate?)\b/i.test(all)&&/\d|days?|months?|years?|weeks?|Tage?|Frist/i.test(delta))add('Timing','Check the time period, trigger and consequences of delay.');
 if(/\b(?:shall|must|may|not|unless|except|waiv\w*|releas\w*|liabil\w*|unconditional\w*|irrevocab\w*|verzicht\w*|haft\w*|nicht|kein\w*|ohne|ausschliess\w*|fraud|wilful|misconduct)\b/i.test(delta))add('Obligation','Check whether this changes an obligation, exception or release.');
 if(JSON.stringify(before.match(/\\(?:ref|cite|parencite|autocite|footcite|label|BeilageNr|Beweisofferte)(?:\[[^\]]*\])?\{[^}]*\}|\b(?:clause|section|article|exhibit|annex|Art\.)\s+[\w.]+/gi))!==JSON.stringify(after.match(/\\(?:ref|cite|parencite|autocite|footcite|label|BeilageNr|Beweisofferte)(?:\[[^\]]*\])?\{[^}]*\}|\b(?:clause|section|article|exhibit|annex|Art\.)\s+[\w.]+/gi)))add('Reference','Verify the referenced provision or exhibit in context.');
 const oldClause=numberedParagraph(before),newClause=numberedParagraph(after);
 const knownWrapper=oldClause&&newClause&&oldClause.number===newClause.number;
 if(!knownWrapper&&(JSON.stringify(before.match(/\\(?:begin|end)\{[^}]+\}/g))!==JSON.stringify(after.match(/\\(?:begin|end)\{[^}]+\}/g))||/\\(?:def|newcommand|renewcommand|documentclass|usepackage|input|include|begin|end|section)/.test(rawDelta)||/\\(?:def|newcommand|renewcommand|documentclass|usepackage)/.test(before+after))||knownWrapper&&/\\(?:def|newcommand|renewcommand|documentclass|usepackage|input|include|begin|end|section)/.test(oldClause.body+newClause.body))add('Structure','This unit includes document structure or configuration. Review the source and compiled PDF.');
 if(/\b(?:signed|signature|electronically|counterpart|unterzeich\w*)\b/i.test(delta))add('Execution','Check the intended signing method and any applicable form requirements.');
 if(before!==after&&plain(before)===plain(after)&&!items.some(x=>x.label==='Structure'))add('Source','The wording is unchanged. Check formatting, commands and comments in the source.');
 return items;
}
export function compare(before,after) {
 if(typeof before!=='string'||typeof after!=='string')throw Error('Two source documents are required.');
 if(Buffer.byteLength(before,'utf8')>300000||Buffer.byteLength(after,'utf8')>300000)throw Error('Please use documents under 300 KB for this prototype.');
 const parts=diffArrays(units(before),units(after),{timeout:3000}); if(!parts)throw Error('The comparison is too large. Compare a smaller document.');
 let segments=[],oldLine=1,newLine=1,oldOffset=0,newOffset=0; const oldStart=before.indexOf('\\begin{document}'),newStart=after.indexOf('\\begin{document}'); const occurrences=new Map();
 const add=(b,a,change)=>{
  const isPreamble=!!((b&&oldStart>=0&&oldOffset<=oldStart)||(a&&newStart>=0&&newOffset<=newStart));
  const isEnd=/^\s*\\end\{document\}\s*$/.test(a||b);
  const heading=(a||b).match(/\\(?:sub)*section\*?\{([^}]+)\}/)?.[1];
  const digest=hash(b+'\0'+a).slice(0,16);const occurrence=occurrences.get(digest)||0;occurrences.set(digest,occurrence+1);
  segments.push({id:`unit-${digest}-${occurrence}`,type:change?'change':'equal',before:b,after:a,oldLine,newLine,preamble:isPreamble,end:isEnd,heading,words:wordDiff(plain(b),plain(a)),sourceOnly:change&&plain(b)===plain(a),signals:change?signals(b,a):[]});
  oldLine+=(b.match(/\n/g)||[]).length;newLine+=(a.match(/\n/g)||[]).length;oldOffset+=b.length;newOffset+=a.length;
 };
 for(let i=0;i<parts.length;i++){
  if(!parts[i].added&&!parts[i].removed){for(const unit of parts[i].value)add(unit,unit,false);continue;}
  const old=[],next=[];
  while(i<parts.length&&(parts[i].added||parts[i].removed)){(parts[i].removed?old:next).push(...parts[i].value);i++;}i--;
  // Pair equal-sized replacements only when each unit is independently balanced.
  if(old.length===next.length&&old.length>1&&[...old,...next].every(u=>structure(u).length===0))old.forEach((u,n)=>add(u,next[n],true));
  else add(old.join(''),next.join(''),true);
 }
 return {id:hash(before+'\0'+after),beforeHash:hash(before),afterHash:hash(after),segments,issues:{before:structure(before),after:structure(after)}};
}
export function resolve(review, decisions={}, pending='after') {
 for(const [id,decision] of Object.entries(decisions)){if(!review.segments.some(s=>s.id===id&&s.type==='change')||!['accepted','rejected'].includes(decision))throw Error('Invalid review decision.');}
 return review.segments.map(s=>s.type==='equal'?s.before:decisions[s.id]==='rejected'?s.before:decisions[s.id]==='accepted'?s.after:pending==='before'?s.before:s.after).join('');
}
export function pendingCount(review, decisions={}){return review.segments.filter(s=>s.type==='change'&&!decisions[s.id]).length;}
