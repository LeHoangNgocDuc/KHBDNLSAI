// Vercel serverless proxy. API keys are supplied on each request and NOT retained/logged.
const API='https://generativelanguage.googleapis.com/v1beta';
const MODEL_RE=/^[a-zA-Z0-9_.-]{3,100}$/;
const keyOk=s=>typeof s==='string'&&s.length>=12&&s.length<=250&&!/[\s\r\n]/.test(s);
const send=(res,code,data)=>res.status(code).setHeader('Cache-Control','no-store').json(data);

export default async function handler(req,res){
  if(req.method!=='POST')return send(res,405,{error:'Chỉ hỗ trợ POST.'});
  const length=+(req.headers['content-length']||0);
  if(length>240000)return send(res,413,{error:'Nội dung vượt giới hạn an toàn 240 KB.'});
  const {op='generate',apiKey,model,payload}=req.body||{};
  if(!keyOk(apiKey))return send(res,400,{error:'API Key không hợp lệ.'});
  if(op!=='models' && op!=='generate')return send(res,400,{error:'Yêu cầu không hỗ trợ.'});
  if(op==='generate' && (!MODEL_RE.test(model||'') || !payload || typeof payload!=='object'))
    return send(res,400,{error:'Thiếu model hoặc dữ liệu phân tích.'});
  try {
    const url=op==='models'?`${API}/models?pageSize=100`:`${API}/models/${model}:generateContent`;
    const controller=new AbortController();
    const timeout=setTimeout(()=>controller.abort(),48000);
    let response;
    try {
      response=await fetch(url,{method:op==='models'?'GET':'POST',headers:{'x-goog-api-key':apiKey,'Content-Type':'application/json'},body:op==='generate'?JSON.stringify(payload):undefined,signal:controller.signal});
    } finally {clearTimeout(timeout);}
    const raw=await response.text();
    let data;try{data=JSON.parse(raw)}catch{data={error:{message:raw.slice(0,160)}}}
    if(!response.ok){
      const description=String(data?.error?.message||'Google Gemini API trả lỗi.').slice(0,450);
      const retryAfter=response.headers.get('retry-after');
      return send(res,response.status>=400&&response.status<600?response.status:502,{error:description,status:response.status,retryAfter});
    }
    if(op==='models'){
      return send(res,200,{models:(data.models||[]).filter(x=>(x.supportedGenerationMethods||[]).includes('generateContent')).map(x=>({name:x.name.replace(/^models\//,''),displayName:x.displayName||x.name,description:(x.description||'').slice(0,160)})),nextPageToken:data.nextPageToken||''});
    }
    const text=(data.candidates||[]).flatMap(c=>(c.content?.parts||[]).map(x=>x.text||'')).join('');
    return send(res,200,{text,finishReason:data.candidates?.[0]?.finishReason||'',usage:data.usageMetadata?{input:data.usageMetadata.promptTokenCount,output:data.usageMetadata.candidatesTokenCount}:null});
  } catch(e){
    return send(res,502,{error:e?.name==='AbortError'?'Gemini phản hồi quá thời gian cho phép.':'Không thể kết nối Gemini API.',status:502});
  }
}
