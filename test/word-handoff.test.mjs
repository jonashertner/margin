import test from 'node:test';
import assert from 'node:assert/strict';
import {createHandoffStore} from '../lib/word-handoff.mjs';
const snapshot={data:Buffer.from('PK fictional test bytes').toString('base64'),name:'Agreement.docx'};
test('local snapshots retain bytes and can be claimed only once',()=>{
  const store=createHandoffStore(),{id}=store.create(snapshot);
  assert.match(id,/^[a-f0-9]{48}$/);
  assert.deepEqual(store.take(id),snapshot);
  assert.throws(()=>store.take(id),/expired or was already opened/);
});
test('expired snapshots cannot be read and no longer occupy capacity',()=>{
  let time=0;const store=createHandoffStore({now:()=>time,ttl:100,limit:1});
  const first=store.create(snapshot);assert.throws(()=>store.create(snapshot),/waiting to open/);
  time=100;assert.throws(()=>store.take(first.id),/expired/);
  const second=store.create(snapshot);assert.deepEqual(store.take(second.id),snapshot);
});
test('invalid bytes and filenames cannot escape the local snapshot boundary',()=>{
  const store=createHandoffStore();assert.throws(()=>store.create({data:'not base64'}),/Word/);
  assert.throws(()=>store.create({data:Buffer.from('not a docx').toString('base64')}),/supported Word/);
  assert.throws(()=>store.create({data:Buffer.alloc(4_000_001).toString('base64')}),/Word/);
  const transfer=store.create({...snapshot,name:'C:\\private\\Agreement.docx'});
  assert.equal(store.take(transfer.id).name,'Agreement.docx');
  assert.throws(()=>store.take('__proto__'),/expired/);
});
