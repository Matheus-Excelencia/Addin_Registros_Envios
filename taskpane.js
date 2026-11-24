Office.onReady(function(){
    iniciarBoasVindas();
    const go = u => window.location.assign(u);
    const n = document.getElementById("btnNovo");
    const h = document.getElementById("btnHistorico");
    const r = document.getElementById("btnResumo");
    const c = document.getElementById("btnConfig");

    if (n) n.addEventListener("click", () => go("novo.html"));
    if (h) h.addEventListener("click", () => go("historico.html"));
    if (r) r.addEventListener("click", () => go("resumo.html"));
    if (c) c.addEventListener("click", () => go("config.html"));
});
function dadosBoasVindas(token){const partes=String(token).split(".");if(partes.length!==3)throw new Error("Token inválido");const payload=partes[1].replace(/-/g,"+").replace(/_/g,"/");return JSON.parse(decodeURIComponent(Array.from(atob(payload.padEnd(Math.ceil(payload.length/4)*4,"=")),c=>"%"+c.charCodeAt(0).toString(16).padStart(2,"0")).join("")));}
function textoBoasVindas(dados={}){const nome=String(dados.name||dados.preferred_username||dados.email||dados.upn||"").trim();const especial=nome.normalize("NFC").toLocaleLowerCase("pt-BR").replace(/\s+/g," ")==="kamylli hemer";return {nome,saudacao:especial?"Bem-vinda":"Bem-vindo",mensagem:especial?"Ou melhor, Kamylli Maionese 😂\nO Matheus Passos é mais belo, e você que teria sorte se fôssemos parentes kkk":"",duracao:especial?6500:2600};}
async function iniciarBoasVindas(){const tela=document.getElementById("boasVindas"),menu=document.getElementById("menuPrincipal");if(!tela||!menu)return;try{if(sessionStorage.getItem("splash106"))return;sessionStorage.setItem("splash106","1");}catch{}let dados={};let timeout;try{const token=await Promise.race([Office.auth.getAccessToken({allowSignInPrompt:false,allowConsentPrompt:false}),new Promise((_,reject)=>{timeout=setTimeout(()=>reject(new Error("SSO indisponível")),1800);})]);dados=dadosBoasVindas(token);}catch{}finally{clearTimeout(timeout);}const texto=textoBoasVindas(dados);document.getElementById("saudacao").textContent=texto.saudacao;document.getElementById("nomeBoasVindas").textContent=texto.nome;document.getElementById("mensagemBoasVindas").textContent=texto.mensagem;tela.style.setProperty("--duracao",texto.duracao+"ms");menu.inert=true;tela.hidden=false;let encerrada=false;let timer;const fechar=()=>{if(encerrada)return;encerrada=true;clearTimeout(timer);tela.classList.add("saindo");menu.inert=false;setTimeout(()=>{tela.hidden=true;},450);};document.getElementById("pularBoasVindas").onclick=fechar;timer=setTimeout(fechar,texto.duracao);}
