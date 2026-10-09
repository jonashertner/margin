import test,{before,after} from 'node:test';
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {once} from 'node:events';
let server,token,base;
before(async()=>{server=spawn(process.execPath,['server.mjs'],{env:{...process.env,PORT:'0',MARGIN_TLS_CERT:'',MARGIN_TLS_KEY:''},stdio:['ignore','pipe','pipe']});const [output]=await once(server.stdout,'data');base=String(output).match(/http:\/\/127\.0\.0\.1:\d+/)[0];const r=await fetch(base+'/api/config');token=(await r.json()).token;});
after(()=>{server?.kill();});
const post=(route,body,headers={})=>fetch(base+'/api/'+route,{method:'POST',headers:{'Content-Type':'application/json','X-Folio-Token':token,...headers},body:JSON.stringify(body)});
test('foreign origin is denied',async()=>assert.equal((await post('compare',{before:'a',after:'b'},{Origin:'https://example.com'})).status,403));
test('missing request token is denied',async()=>assert.equal((await post('compare',{before:'a',after:'b'},{'X-Folio-Token':''})).status,403));
test('private local paths cannot be requested',async()=>assert.equal((await fetch(base+'/server.mjs')).status,404));
test('pending changes block clean source export',async()=>{const r=await post('export',{before:'Old.',after:'New.',decisions:{},mode:'clean'});assert.equal(r.status,400);assert.match((await r.json()).error,/remaining changes/);});
test('mixed decisions export exact source, including BOM and CRLF',async()=>{const before='\uFEFFOne.\r\n\r\nTwo.',after='\uFEFFNew one.\r\n\r\nNew two.';const r=await(await post('compare',{before,after})).json();const c=r.segments.filter(s=>s.type==='change');const decisions={[c[0].id]:'accepted',[c[1].id]:'rejected'};const out=await post('export',{before,after,decisions,mode:'clean'});assert.equal(out.status,200);assert.equal((await out.json()).source,'\uFEFFNew one.\r\n\r\nTwo.');});
test('invalid structural mix does not export',async()=>{const before='Text.',after='Text. \\footnote{unclosed';const r=await(await post('compare',{before,after})).json();const decisions=Object.fromEntries(r.segments.filter(s=>s.type==='change').map(s=>[s.id,'accepted']));const out=await post('export',{before,after,decisions,mode:'clean'});assert.equal(out.status,400);assert.match((await out.json()).error,/brace/);});
test('reviewed redline blocks pending decisions before compilation',async()=>{const r=await post('pdf',{before:'Old.',after:'New.',decisions:{},mode:'comparison'});assert.equal(r.status,400);assert.match((await r.json()).error,/remaining changes/);});
test('a single document exposes a prose editor and can be revised safely',async()=>{const source='\\documentclass{article}\n\\begin{document}\n\nPay 30 days.\n\n\\end{document}\n';const review=await(await post('compare',{before:source,after:source})).json();const target=review.segments.find(s=>s.editorAfter?.editable);const response=await post('revise',{session:{before:source,after:source,decisions:{},comments:{},audit:[]},id:target.id,text:'Pay 45 days.'});assert.equal(response.status,200);const next=await response.json();assert.equal(next.before,source);assert.equal(next.authoring.originalAfter,source);assert.equal(next.after,source.replace('30','45'));assert.equal(Object.values(next.decisions)[0],'accepted');});
test('source text survives a UTF-8 character split across network chunks',async()=>{
  const {request}=await import('node:http');
  const source='Zürich — pièce 🖋️.';const bytes=Buffer.from(JSON.stringify({before:source,after:source}));
  const split=bytes.indexOf(Buffer.from('ü'))+1;
  const data=await new Promise((resolve,reject)=>{const req=request(base+'/api/compare',{method:'POST',headers:{'Content-Type':'application/json','X-Folio-Token':token}},res=>{let body='';res.on('data',chunk=>body+=chunk);res.on('end',()=>resolve({status:res.statusCode,body:JSON.parse(body)}));});req.on('error',reject);req.write(bytes.subarray(0,split));setTimeout(()=>req.end(bytes.subarray(split)),10);});
  assert.equal(data.status,200);assert.equal(data.body.segments.map(s=>s.before).join(''),source);
});

test('Word handoff requires local token and can only be read once',async()=>{
 const snapshot={data:Buffer.from('PK fictional docx bytes').toString('base64'),name:'Agreement.docx'};
 assert.equal((await post('word/handoff',snapshot,{'X-Folio-Token':''})).status,403);
 const r=await post('word/handoff',snapshot);assert.equal(r.status,200);const {url}=await r.json();
 assert.ok(url.startsWith(base+'/#word='));assert.ok(!url.includes(snapshot.data));
 const ticket=new URL(url).hash.slice(6);
 const read=await post('word/handoff/read',{ticket});assert.equal(read.status,200);assert.deepEqual(await read.json(),snapshot);
 assert.equal((await post('word/handoff/read',{ticket})).status,400);
});
test('Office runtime is permitted only on the Word pane',async()=>{
 const page=await fetch(base+'/');const pane=await fetch(base+'/word');assert.equal(pane.status,200);
 assert.doesNotMatch(page.headers.get('content-security-policy'),/appsforoffice/);
 assert.match(pane.headers.get('content-security-policy'),/script-src 'self' https:\/\/appsforoffice.microsoft.com/);
 assert.match(page.headers.get('content-security-policy'),/frame-ancestors 'none'/);
 assert.equal((await fetch(base+'/office/manifest.xml')).status,200);
 assert.equal((await fetch(base+'/office/../server.mjs')).status,404);
});
