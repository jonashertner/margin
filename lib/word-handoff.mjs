import {randomBytes} from 'node:crypto';
import {decodeWord} from './word.mjs';

// Short-lived, one-use local transfers. Document bytes never appear in URLs.
export function createHandoffStore({now=Date.now,ttl=5*60*1000,limit=3}={}) {
  const entries=new Map();
  function prune(){for(const [id,entry] of entries)if(entry.expires<=now())entries.delete(id);}
  return {
    create({data,name}) {
      decodeWord(data);
      prune();
      if(entries.size>=limit)throw Error('Three documents are waiting to open. Open one of them or wait five minutes before trying again.');
      const id=randomBytes(24).toString('hex'),expires=now()+ttl;
      const safeName=String(name||'Word document.docx').split(/[\\/]/).pop().replace(/[\u0000-\u001f]/g,'').slice(0,240)||'Word document.docx';
      entries.set(id,{data,name:safeName,expires});
      setTimeout(()=>entries.delete(id),ttl).unref();
      return {id,expiresAt:new Date(expires).toISOString()};
    },
    take(id) {
      prune();
      if(typeof id!=='string'||!/^[a-f0-9]{48}$/.test(id)||!entries.has(id))throw Error('This Word snapshot has expired or was already opened. Open it again from the Margin pane in Word.');
      const {data,name}=entries.get(id);
      entries.delete(id);
      return {data,name};
    }
  };
}
