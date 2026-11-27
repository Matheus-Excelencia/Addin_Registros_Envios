/* UI preferences only. No tokens, names, records or workbook URLs persisted. */
const PersonalPrefs=(()=>{
 let prefix=null,store=null,kind="memory",memory={},pending=null;
 const allowed=new Set(["sempre","opcional","oculto"]);
 const hash=async s=>Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256",new TextEncoder().encode(s))),b=>b.toString(16).padStart(2,"0")).join("");
 function claims(token){try{let p=String(token).split(".")[1].replace(/-/g,"+").replace(/_/g,"/");p=p.padEnd(Math.ceil(p.length/4)*4,"=");return JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(p),c=>c.charCodeAt(0))));}catch{return {};}}
 function identity(c){return typeof c.tid==="string"&&c.tid&&typeof c.oid==="string"&&c.oid?JSON.stringify([c.tid,c.oid]):typeof c.iss==="string"&&c.iss&&typeof c.sub==="string"&&c.sub?JSON.stringify([c.iss,c.sub]):null;}
 async function silent(){let timer;try{return await Promise.race([Promise.resolve().then(()=>Office.auth.getAccessToken({allowSignInPrompt:false,allowConsentPrompt:false})),new Promise(resolve=>timer=setTimeout(()=>resolve(null),2200))]);}catch{return null;}finally{clearTimeout(timer);}}
 async function resolve(){try{const workbook=await RegistroSecurity.scope();if(!workbook)return;
 if(window.REGISTROS_DEMO===true){prefix="registra.prefs.v2.demo";store=sessionStorage;kind="demo";return;}
 if(!window.REGISTROS_HOMOLOGACAO&&typeof Office?.auth?.getAccessToken==="function"){const id=identity(claims(await silent()));if(id){prefix="registra.prefs.v2."+await hash(JSON.stringify([workbook,id]));store=localStorage;kind="account";try{sessionStorage.removeItem("registra.prefs.v2.session."+workbook);}catch{}return;}}
 prefix="registra.prefs.v2.session."+workbook;store=sessionStorage;kind="session";
 }catch{prefix=null;store=null;kind="memory";}}
 function init(){if(!pending)pending=resolve();return pending;}
 function read(){if(!store||!prefix)return memory;try{const o=JSON.parse(store.getItem(prefix)||"{}");return o&&typeof o==="object"&&!Array.isArray(o)?o:{};}catch{return memory;}}
 function write(change){const next={...read(),...change};memory=next;if(!store||!prefix)return false;try{store.setItem(prefix,JSON.stringify(next));return true;}catch{store=null;kind="memory";return false;}}
 function getOrder(ch){const a=read().order?.[ch];return Array.isArray(a)?a.filter(x=>typeof x==="string").slice(0,300):[];}
 function setOrder(ch,keys){if(!["email","sms"].includes(ch)||!Array.isArray(keys)||keys.length>300||keys.some(k=>typeof k!=="string"||k.length>256))return false;return write({order:{...read().order,[ch]:[...new Set(keys)]}});}
 function mode(key,m,def){const v=read().modes?.[m]?.[key];return allowed.has(v)?v:def;}
 function setMode(key,m,v){if(!["historico","resumo"].includes(m)||!allowed.has(v))return false;const d=read().modes||{};return write({modes:{...d,[m]:{...d[m],[key]:v}}});}
 function enabled(key,m){return read().enabled?.[m]?.[key]===true;}
 function setEnabled(key,m,v){if(!["historico","resumo"].includes(m))return false;const d=read().enabled||{};return write({enabled:{...d,[m]:{...d[m],[key]:v===true}}});}
 function label(){return kind==="account"?"Preferências da sua conta Microsoft, neste navegador e arquivo.":kind==="demo"?"Preferências da demonstração, nesta sessão.":kind==="session"?"Conta Microsoft não identificada. Preferências temporárias nesta aba; feche-a ao trocar de usuário.": "Preferências apenas nesta tela; armazenamento ou identidade do arquivo indisponível.";}
 return Object.freeze({init,getOrder,setOrder,mode,setMode,enabled,setEnabled,label,kind:()=>kind});
})();