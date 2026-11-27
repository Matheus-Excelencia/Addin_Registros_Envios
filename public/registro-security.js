/* Excel write boundary and workbook-bound, one-use duplication. */
const RegistroSecurity=(()=>{
 const KEY="registroDuplicado",TTL=120000;
 function cell(v){if(typeof v==="string"&&/^[\s\u0000-\u001f\u007f\u200b-\u200f\u202a-\u202e\u2060-\u206f\ufeff]*[=+\-@']/.test(v))return "'"+v;if(typeof v==="number"&&!Number.isFinite(v))throw Error("Valor numérico inválido.");return v;}
 function clear(){try{sessionStorage.removeItem(KEY)}catch{}}
 async function scope(){
  if(window.REGISTROS_DEMO===true)return "demo:v1";
  const doc=typeof Office!=="undefined"?Office.context?.document:null;
  if(!doc||typeof doc.getFilePropertiesAsync!=="function")return null;
  const url=await new Promise(resolve=>{const timer=setTimeout(()=>resolve(null),2500);try{doc.getFilePropertiesAsync(r=>{clearTimeout(timer);resolve(r?.status==="succeeded"&&typeof r.value?.url==="string"?r.value.url:null)})}catch{clearTimeout(timer);resolve(null)}});
  if(!url||!url.trim()||!globalThis.crypto?.subtle)return null;
  try{const bytes=await crypto.subtle.digest("SHA-256",new TextEncoder().encode(url));return "office:"+Array.from(new Uint8Array(bytes),n=>n.toString(16).padStart(2,"0")).join("")}catch{return null}
 }
 async function store(payload,source){clear();const workbook=await scope();if(!workbook||workbook!==source||!["email","sms"].includes(payload?.tipo))return false;try{sessionStorage.setItem(KEY,JSON.stringify({version:1,workbook,expires:Date.now()+TTL,payload}));return true}catch{return false}}
 async function take(type){let raw;try{raw=sessionStorage.getItem(KEY);sessionStorage.removeItem(KEY)}catch{return null}if(!raw)return null;try{const e=JSON.parse(raw);if(e.version!==1||!Number.isFinite(e.expires)||e.expires<=Date.now()||e.expires>Date.now()+TTL||e.payload?.tipo!==type)return null;const workbook=await scope();return workbook&&workbook===e.workbook&&e.expires>Date.now()?e.payload:null}catch{return null}}
 return Object.freeze({cell,scope,store,take,clear});
})();