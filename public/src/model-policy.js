// Central model policy for Gemini text/JSON generation.
// Never assume that a listed model is callable by every API key.
export const PREFERRED_MODELS = Object.freeze([
  'gemini-3.8-flash',
  'gemini-3.7-flash',
  'gemini-3.6-flash',
  'gemini-3.5-flash',
  'gemini-3.5-flash-lite',
  'gemini-3.1-flash-lite'
]);

export const DEFAULT_PRIMARY = PREFERRED_MODELS[0];
export const DEFAULT_BACKUP = PREFERRED_MODELS[2];

export function isTextGenerationModel(value) {
  const name=String(value||'').replace(/^models\//,'').toLowerCase();
  return /^gemini-[\w.-]+$/.test(name)
    && !/(?:live|tts|image|audio|transcribe|embedding|robotics|omni|computer-use)/.test(name);
}

export function sortModelIds(names) {
  const unique=[...new Set((names||[]).map(x=>String(x||'').replace(/^models\//,'').trim()).filter(isTextGenerationModel))];
  return unique.sort((a,b)=>{
    const rank=s=>{
      const i=PREFERRED_MODELS.indexOf(s);
      if(i>=0)return i;
      if(s==='gemini-3.1-pro-preview')return 20;
      if(s.startsWith('gemini-3-'))return 25;
      if(s.startsWith('gemini-2.5-'))return 100;
      return 50;
    };
    return rank(a)-rank(b)||a.localeCompare(b);
  });
}

// The manual mode intentionally never overrides the teacher's selection.
export function chooseModelOrder({primary=DEFAULT_PRIMARY,backup=DEFAULT_BACKUP,mode='smart',available=PREFERRED_MODELS,cycle=0}={}) {
  if(mode==='manual')return isTextGenerationModel(primary)?[primary]:[];
  const first=mode==='round'&&cycle%2 ? [backup,primary] : [primary,backup];
  const pool=sortModelIds(available);
  const additions=mode==='smart'||mode==='round' ? [...PREFERRED_MODELS,...pool] : [];
  const models=[...new Set([...first,...additions].filter(isTextGenerationModel))];
  // Do not attempt models absent from the live models list, if a list exists.
  const filtered=pool.length?models.filter(x=>pool.includes(x)):models;
  return filtered.slice(0,5);
}

export function classifyGeminiError(status,message='') {
  const m=String(message).toLowerCase();
  if(/no longer available|not available to new users|not found for api version|not found|no longer supported|model.*does not exist/.test(m))return 'MODEL_UNAVAILABLE';
  if(status===429)return 'QUOTA_EXCEEDED';
  if(status===401||/api key not valid|api_key_invalid|api key expired/.test(m))return 'INVALID_KEY';
  if(status===403)return 'ACCESS_DENIED';
  if(status===400&&/responsemimetype|generationconfig|json schema/.test(m))return 'REQUEST_UNSUPPORTED';
  if(status===408||status===502||status===503||status===504)return 'TRANSIENT';
  return 'OTHER';
}
