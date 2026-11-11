import {test} from 'node:test';
import assert from 'node:assert/strict';
import {DEFAULT_PRIMARY,DEFAULT_BACKUP,PREFERRED_MODELS,sortModelIds,isTextGenerationModel,chooseModelOrder,classifyGeminiError} from '../public/src/model-policy.js';
import handler from '../api/gemini.js';

function response(){const r={statusCode:200,headers:{},body:null};r.status=n=>(r.statusCode=n,r);r.setHeader=(k,v)=>(r.headers[k]=v,r);r.json=d=>(r.body=d,r);return r;}

test('the new default never starts with legacy 2.5 models',()=>{
  assert.equal(DEFAULT_PRIMARY,'gemini-3.8-flash');
  assert.equal(DEFAULT_BACKUP,'gemini-3.6-flash');
  assert.ok(PREFERRED_MODELS.every(x=>!x.startsWith('gemini-2.5-')));
  assert.deepEqual(chooseModelOrder({available:PREFERRED_MODELS}).slice(0,2),[DEFAULT_PRIMARY,DEFAULT_BACKUP]);
});

test('model discovery drops non-text models and prioritizes modern stable Flash',()=>{
  assert.equal(isTextGenerationModel('models/gemini-3.8-flash'),true);
  for(const x of ['gemini-3.8-live','gemini-3.8-flash-tts','gemini-3.1-flash-image','gemini-embedding-001','gemini-3.5-transcribe'])assert.equal(isTextGenerationModel(x),false);
  assert.deepEqual(sortModelIds(['models/gemini-2.5-pro','models/gemini-3.6-flash','models/gemini-3.8-flash','models/gemini-3.8-live']),['gemini-3.8-flash','gemini-3.6-flash','gemini-2.5-pro']);
});

test('automatic fallback uses live available list, manual mode uses one selection',()=>{
  const available=['gemini-3.5-flash-lite','gemini-3.6-flash'];
  assert.deepEqual(chooseModelOrder({available,primary:'gemini-2.5-flash',backup:'gemini-3.6-flash',mode:'smart'}),['gemini-3.6-flash','gemini-3.5-flash-lite']);
  assert.deepEqual(chooseModelOrder({available,primary:'gemini-2.5-flash',backup:'gemini-3.6-flash',mode:'manual'}),['gemini-2.5-flash']);
});

test('model unavailability is distinct from quota and invalid key',()=>{
  assert.equal(classifyGeminiError(404,'This model is no longer available to new users.'),'MODEL_UNAVAILABLE');
  assert.equal(classifyGeminiError(429,'RESOURCE_EXHAUSTED'),'QUOTA_EXCEEDED');
  assert.equal(classifyGeminiError(400,'API key not valid'),'INVALID_KEY');
});

test('models endpoint supports pagination and omits irrelevant modalities',async()=>{
  const original=globalThis.fetch;let called='';
  globalThis.fetch=async(url)=>{called=String(url);return {ok:true,text:async()=>JSON.stringify({models:[
    {name:'models/gemini-3.8-flash',supportedGenerationMethods:['generateContent']},
    {name:'models/gemini-3.8-live',supportedGenerationMethods:['generateContent']},
    {name:'models/gemini-3.1-flash-image',supportedGenerationMethods:['generateContent']}
  ],nextPageToken:'next-one'})};};
  try{
    const res=response();await handler({method:'POST',headers:{},body:{op:'models',apiKey:'AN_EXAMPLE_TEST_KEY_123',payload:{pageToken:'prev-one'}}},res);
    assert.equal(res.statusCode,200);
    assert.equal(res.body.models.length,1);
    assert.equal(res.body.models[0].name,'gemini-3.8-flash');
    assert.match(called,/pageToken=prev-one/);
    assert.equal(res.body.nextPageToken,'next-one');
  } finally{globalThis.fetch=original;}
});

test('proxy reports model unavailability without leaking API key',async()=>{
  const original=globalThis.fetch;
  globalThis.fetch=async()=>({ok:false,status:404,headers:{get:()=>null},text:async()=>JSON.stringify({error:{message:'models/gemini-2.5-flash is no longer available to new users'}})});
  try{
    const res=response();await handler({method:'POST',headers:{},body:{op:'generate',model:'gemini-2.5-flash',apiKey:'AN_EXAMPLE_TEST_KEY_123',payload:{contents:[{parts:[{text:'test'}]}]}}},res);
    assert.equal(res.statusCode,404);
    assert.equal(res.body.code,'MODEL_UNAVAILABLE');
    assert.ok(!JSON.stringify(res.body).includes('AN_EXAMPLE_TEST_KEY_123'));
  }finally{globalThis.fetch=original;}
});
