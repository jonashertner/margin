import { mkdtemp, writeFile, readFile, rm, realpath, access, readdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import {xelatex as engine,texRoot,pdfPython as python,pdftoppm as renderer,sandbox,pythonRuntime,dependencyRoots,sandboxProfile} from './runtime.mjs';
const run=promisify(execFile);
export {texRoot};
const vendor=new URL('./vendor/latexdiff-so',import.meta.url).pathname;
export async function available(){return !!(engine&&sandbox&&python&&(await pythonRuntime(python)).modules.pypdf);}
export async function compile(source,original){
 if(!await available())throw Error('Local PDF compilation is unavailable. Export the LaTeX source to compile with your own setup.');
 const job=await realpath(await mkdtemp(path.join(tmpdir(),'margin-pdf-')));
 const profile=sandboxProfile(job,[texRoot,...dependencyRoots(engine)]);
 try{
  let target=source;
  if(original!==undefined){
   await writeFile(path.join(job,'original.tex'),original);await writeFile(path.join(job,'proposed.tex'),source);
   const result=await run('/usr/bin/perl',[vendor,'--encoding=utf8','--type=UNDERLINE','--config=MAXCHANGESLETTER=0',path.join(job,'original.tex'),path.join(job,'proposed.tex')],{timeout:20000,maxBuffer:4000000,cwd:job});target=result.stdout.replace('\\begin{document}',String.raw`\definecolor{FolioAdd}{rgb}{0.180,0.290,0.620}
\definecolor{FolioDelete}{rgb}{0.647,0.302,0.290}
\renewcommand{\DIFadd}[1]{{\protect\color{FolioAdd}\hspace{0.12em}\uline{#1}}}
\renewcommand{\DIFdel}[1]{{\protect\color{FolioDelete}\sout{#1}}}
\emergencystretch=3em
\raggedbottom
\begin{document}`);
  }
  await writeFile(path.join(job,'document.tex'),target);
  const env={PATH:path.dirname(engine)+':/usr/bin:/bin',HOME:job,TMPDIR:job,TEXMFHOME:path.join(job,'texmf'),TEXMFVAR:path.join(job,'var'),TEXMFCONFIG:path.join(job,'config'),openout_any:'p',shell_escape:'0'};
  let output='';
  for(let pass=0;pass<2;pass++){
   try{const r=await run(sandbox,['-p',profile,engine,'-no-shell-escape','-no-parse-first-line','-interaction=nonstopmode','-halt-on-error','-file-line-error','document.tex'],{cwd:job,env,timeout:25000,maxBuffer:3000000});output=r.stdout;}
   catch(e){const log=String(e.stdout||e.stderr||e.message);const lines=log.split('\n');const start=lines.findIndex(l=>/^!|^\.\/document.tex:\d/.test(l));throw Error('PDF could not be compiled. '+(start>=0?lines.slice(start,start+7).join('\n'):lines.slice(-10).join('\n')));}
  }
  const cleaner=new URL('./clean-pdf.py',import.meta.url).pathname;
  const cleanProfile=sandboxProfile(job,[...(await pythonRuntime(python)).roots,...dependencyRoots(python),path.dirname(cleaner)]);
  await run(sandbox,['-p',cleanProfile,python,'-I','-B',cleaner,path.join(job,'document.pdf'),path.join(job,'clean.pdf')],{timeout:10000,maxBuffer:200000,cwd:job,env});
  const renderProfile=sandboxProfile(job,dependencyRoots(renderer));
  let images=[],previewError=null;
  try{if(!renderer)throw Error('pdftoppm unavailable');await run(sandbox,['-p',renderProfile,renderer,'-png','-scale-to','1600','-f','1','-l','30','clean.pdf','page'],{cwd:job,env,timeout:25000,maxBuffer:200000});const pages=(await readdir(job)).filter(f=>/^page-\d+\.png$/.test(f)).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));images=await Promise.all(pages.map(async f=>(await readFile(path.join(job,f))).toString('base64')));}catch{previewError='The PDF compiled, but page images could not be rendered. Download the PDF to view it.';}
  return {pdf:await readFile(path.join(job,'clean.pdf')),images,previewError,warnings:output.split('\n').filter(l=>/Warning:|Overfull|Underfull/.test(l)).slice(0,15)};
 }finally{await rm(job,{recursive:true,force:true});}
}
