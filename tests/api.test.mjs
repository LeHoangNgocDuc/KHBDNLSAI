import {test} from 'node:test';
import assert from 'node:assert/strict';
import handler from '../api/gemini.js';
function response(){const obj={statusCode:200,headers:{},body:null};obj.status=n=>(obj.statusCode=n,obj);obj.setHeader=(k,v)=>(obj.headers[k]=v,obj);obj.json=x=>(obj.body=x,obj);return obj;}
test('API rejects invalid BYOK key without calling provider',async()=>{
 const req={method:'POST',headers:{},body:{op:'models',apiKey:'short'}};const res=response();await handler(req,res);
 assert.equal(res.statusCode,400);assert.match(res.body.error,/API Key/i);
});
test('API filters supported model list, never echoes a key',async()=>{
 const prev=globalThis.fetch;
 globalThis.fetch=async(url,opts)=>{assert.match(url,/v1beta\/models/);assert.equal(opts.headers['x-goog-api-key'],'AN_EXAMPLE_TEST_KEY_123');return {ok:true,text:async()=>JSON.stringify({models:[{name:'models/gemini-flash',supportedGenerationMethods:['generateContent']},{name:'models/other',supportedGenerationMethods:['embedContent']}]})};};
 try{const req={method:'POST',headers:{},body:{op:'models',apiKey:'AN_EXAMPLE_TEST_KEY_123'}};const res=response();await handler(req,res);assert.equal(res.statusCode,200);assert.equal(res.body.models.length,1);assert.ok(!JSON.stringify(res.body).includes('AN_EXAMPLE_TEST_KEY_123'));}finally{globalThis.fetch=prev;}
});
