import http from 'node:http';
import https from 'node:https';
import {createHandoffStore} from './lib/word-handoff.mjs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { randomBytes } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { compare,resolve,pendingCount,structure,hash } from './lib/review.mjs';
import { editableParagraph,reviseParagraph } from './lib/authoring.mjs';
import {wordAvailable,importWord,exportWord,originalWord,restoreWordFiles,importText,mapComparisonComments} from './lib/word.mjs';
import { compile,available } from './lib/compile.mjs';
const root=path.dirname(fileURLToPath(import.meta.url));
const port=Number(process.env.PORT??4317),httpsPort=Number(process.env.MARGIN_HTTPS_PORT??4318);
const token=randomBytes(24).toString('hex'),handoffs=createHandoffStore();
const tlsCert=process.env.MARGIN_TLS_CERT,tlsKey=process.env.MARGIN_TLS_KEY;
if(!!tlsCert!==!!tlsKey)throw Error('Set both MARGIN_TLS_CERT and MARGIN_TLS_KEY to enable Word.');
let compiling=false;
const send=(res,status,data)=>{res.writeHead(status,{'Content-Type':'application/json'});res.end(JSON.stringify(data));};
async function body(req){const chunks=[];let size=0;for await(const chunk of req){size+=chunk.length;if(size>25000000)throw Error('Review file is too large.');chunks.push(chunk);}return JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(Buffer.concat(chunks)));}
const handle=async(req,res)=>{
 const secure=!!req.socket.encrypted,localPort=req.socket.localPort;
 const hosts=secure?[`localhost:${localPort}`,`127.0.0.1:${localPort}`]:[`127.0.0.1:${localPort}`];
 const origin=`${secure?'https':'http'}://${req.headers.host}`;
 res.setHeader('X-Content-Type-Options','nosniff');res.setHeader('Referrer-Policy','no-referrer');res.setHeader('Cache-Control','no-store');
 res.setHeader('Content-Security-Policy',"default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; font-src 'self'; connect-src 'self'; frame-src 'self' blob:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'");
 if(!hosts.includes(req.headers.host)||req.headers.origin&&req.headers.origin!==origin){send(res,403,{error:'Only this local workspace can make requests.'});return;}
 try{
  const url=new URL(req.url,origin);
  if(url.pathname==='/word'||url.pathname==='/word/')res.setHeader('Content-Security-Policy',"default-src 'self'; script-src 'self' https://appsforoffice.microsoft.com; style-src 'self'; img-src 'self' data:; font-src 'self'; connect-src 'self' https://appsforoffice.microsoft.com; frame-src https://appsforoffice.microsoft.com; object-src 'none'; base-uri 'none'; frame-ancestors 'self' https://*.officeapps.live.com https://*.office.com https://*.microsoft365.com https://*.cloud.microsoft");
  if(req.method==='GET'&&url.pathname==='/api/config'){send(res,200,{token,webOrigin:`http://127.0.0.1:${server.address().port}`,pdf:await available(),word:await wordAvailable()});return;}
  if(req.method==='POST'&&url.pathname.startsWith('/api/')){
   if(req.headers['x-folio-token']!==token){send(res,403,{error:'Reload the workspace to reconnect.'});return;}
   const data=await body(req);
   if(url.pathname==='/api/word/handoff'){const transfer=handoffs.create(data);const httpPort=server.address().port;send(res,200,{url:`http://127.0.0.1:${httpPort}/#word=${transfer.id}`,expiresAt:transfer.expiresAt});return;}
   if(url.pathname==='/api/word/handoff/read'){send(res,200,handoffs.take(data.ticket));return;}
   if(url.pathname==='/api/word/map-comments'){send(res,200,mapComparisonComments(data.before,data.after,data.sources));return;}
   if(url.pathname==='/api/word/import'){send(res,200,await importWord(data));return;}
   if(url.pathname==='/api/word/export'){send(res,200,await exportWord(data.session,data.mode));return;}
   if(url.pathname==='/api/word/original'){send(res,200,await originalWord(data.hash));return;}
   if(url.pathname==='/api/word/restore'){send(res,200,{files:await restoreWordFiles(data.files)});return;}
   if(url.pathname==='/api/text/import'){send(res,200,importText(data.text,data.title));return;}
   if(url.pathname==='/api/compare'){const result=compare(data.before,data.after);for(const s of result.segments){s.editorAfter=editableParagraph(s);s.editorBefore=editableParagraph(s,'before');}send(res,200,result);return;}
   if(url.pathname==='/api/revise'){const result=compare(data.session.before,data.session.after);send(res,200,reviseParagraph(result,data.session,data.id,data.text,compare));return;}
   if(url.pathname==='/api/export'||url.pathname==='/api/pdf'){
    const review=compare(data.before,data.after),decisions=data.decisions||{};
    const material=resolve(review,decisions),pending=pendingCount(review,decisions);
    if(!['clean','preview','comparison','proposal'].includes(data.mode))throw Error('Unknown export mode.');
    if((url.pathname==='/api/export'||['clean','comparison'].includes(data.mode))&&pending)throw Error(`Review the ${pending} remaining changes before clean export.`);
    if(url.pathname==='/api/export'){const issues=structure(material);if(issues.length)throw Error(issues.join(' '));send(res,200,{source:material,hash:hash(material)});return;}
    if(compiling){send(res,409,{error:'Another PDF is compiling. Please try again shortly.'});return;}
    compiling=true;
    try{const source=data.mode==='proposal'?(data.authoring?.originalAfter??data.after):material;if(data.mode==='proposal')compare(data.before,source);const result=await compile(source,['comparison','proposal'].includes(data.mode)?data.before:undefined);send(res,200,{pdf:result.pdf.toString('base64'),images:result.images,previewError:result.previewError,sourceHash:hash(source),warnings:result.warnings,pending,mode:data.mode});}finally{compiling=false;}return;
   }
   send(res,404,{error:'Unknown action.'});return;
  }
  const assets={'/review-storage.js':'review-storage.js','/review-record.js':'review-record.js','/word':'word.html','/word/':'word.html','/word/setup':'word-setup.html','/word.js':'word.js','/word-host.js':'word-host.js','/word.css':'word.css','/':'index.html','/app.js':'app.js','/styles.css':'styles.css','/demo.js':'demo.js','/favicon.svg':'favicon.svg','/fonts/PaperMono.woff2':'fonts/PaperMono.woff2','/fonts/Timeless.ttf':'fonts/Timeless.ttf','/fonts/Timeless-Bold.ttf':'fonts/Timeless-Bold.ttf'};
  const officeAsset=/^\/office\/(?:manifest\.xml|icon-(?:16|32|64|80|128)\.png)$/.test(url.pathname)?url.pathname.slice(1):null;
  const filename=officeAsset||assets[url.pathname];if(req.method!=='GET'||!filename){res.writeHead(404);res.end('Not found');return;}
  const ext=path.extname(filename);const mime={'.xml':'application/xml; charset=utf-8','.png':'image/png','.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.svg':'image/svg+xml','.woff2':'font/woff2','.ttf':'font/ttf'};
  let bytes;try{bytes=await readFile(path.join(root,officeAsset?'':'public',filename));}catch(e){if(e.code==='ENOENT'){res.writeHead(404);res.end('Not found');return;}throw e;}
  res.writeHead(200,{'Content-Type':mime[ext]});res.end(bytes);
 }catch(e){send(res,400,{error:e.message||'Something went wrong.'});}
};
const server=http.createServer(handle);
let tls;
const stopListeners=()=>{server.close();tls?.close();process.exitCode=1;};
server.listen(port,'127.0.0.1',()=>console.log(`Margin is ready at http://127.0.0.1:${server.address().port}`));
if(tlsCert){
 tls=https.createServer({cert:await readFile(tlsCert),key:await readFile(tlsKey),minVersion:'TLSv1.2'},handle);
 tls.listen(httpsPort,'127.0.0.1',()=>console.log(`Margin for Word is ready at https://localhost:${tls.address().port}/word`));
 tls.on('error',error=>{console.error(`Word listener could not start: ${error.message}`);stopListeners();});
}
server.on('error',error=>{console.error(`Margin could not start: ${error.message}`);stopListeners();});
