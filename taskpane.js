function iniciarMenu(){
    const go = u => window.location.assign(u);
    const n = document.getElementById("btnNovo");
    const h = document.getElementById("btnHistorico");
    const r = document.getElementById("btnResumo");
    const c = document.getElementById("btnConfig");

    if (n) n.addEventListener("click", () => go("novo.html"));
    if (h) h.addEventListener("click", () => go("historico.html"));
    if (r) r.addEventListener("click", () => go("resumo.html"));
    if (c) c.addEventListener("click", () => go("config.html"));
    iniciarBoasVindas();
}
if(typeof document!=="undefined"){if(typeof Office!=="undefined"&&Office.onReady){let iniciado=false;const iniciar=()=>{if(iniciado)return;iniciado=true;iniciarMenu();};const espera=setTimeout(iniciar,3000);Office.onReady(()=>{clearTimeout(espera);iniciar();});}else iniciarMenu();}
function dadosBoasVindas(token){const partes=String(token).split(".");if(partes.length!==3)throw new Error("Token inválido");const payload=partes[1].replace(/-/g,"+").replace(/_/g,"/");return JSON.parse(decodeURIComponent(Array.from(atob(payload.padEnd(Math.ceil(payload.length/4)*4,"=")),c=>"%"+c.charCodeAt(0).toString(16).padStart(2,"0")).join("")));}
function textoBoasVindas(dados={}){const nome=String(dados.name||dados.preferred_username||dados.email||dados.upn||"").trim();const especial=nome.normalize("NFC").toLocaleLowerCase("pt-BR").replace(/\s+/g," ")==="kamylli hemer";return {nome,saudacao:especial?"Bem-vinda":"Bem-vindo",mensagem:especial?"Ou melhor, Kamylli Maionese 😂\nO Matheus Passos é mais belo, e você que teria sorte se fôssemos parentes kkk":"",duracao:especial?6500:2600};}
function digitarBoasVindas(campos, duracao, reduzido){
    const partes=campos.map(([el,texto])=>({el,texto,letras:typeof Intl!=="undefined"&&Intl.Segmenter?Array.from(new Intl.Segmenter("pt-BR",{granularity:"grapheme"}).segment(texto),x=>x.segment):Array.from(texto)}));
    const total=partes.reduce((n,p)=>n+p.letras.length,0);
    let timer,indice=0,letra=0,parado=false;
    const concluir=()=>{parado=true;clearTimeout(timer);for(const p of partes){p.el.textContent=p.texto;p.el.classList?.remove("digitando");}};
    if(reduzido||!total){concluir();return concluir;}
    for(const p of partes)p.el.textContent="";
    const intervalo=Math.max(8,Math.min(55,Math.floor(duracao*.55/total)));
    function passo(){
        if(parado)return;
        while(indice<partes.length&&letra>=partes[indice].letras.length){partes[indice].el.classList?.remove("digitando");indice++;letra=0;}
        if(indice>=partes.length){concluir();return;}
        const p=partes[indice];p.el.classList?.add("digitando");p.el.textContent+=p.letras[letra++];timer=setTimeout(passo,intervalo);
    }
    passo();return concluir;
}
async function iniciarBoasVindas(){
    const tela=document.getElementById("boasVindas"),menu=document.getElementById("menuPrincipal");
    if(!menu)return;
    if(!tela){menu.hidden=false;return;}
    try{if(sessionStorage.getItem("splash106")){menu.hidden=false;return;}sessionStorage.setItem("splash106","1");}catch{}
    menu.hidden=true;menu.inert=true;
    let dados={},timeout;
    try{const token=await Promise.race([Office.auth.getAccessToken({allowSignInPrompt:false,allowConsentPrompt:false}),new Promise((_,reject)=>{timeout=setTimeout(()=>reject(new Error("SSO indisponível")),1800);})]);dados=dadosBoasVindas(token);}catch{}finally{clearTimeout(timeout);}
    const texto=textoBoasVindas(dados);
    const reduzido=typeof matchMedia==="function"&&matchMedia("(prefers-reduced-motion: reduce)").matches;
    tela.style.setProperty("--duracao",texto.duracao+"ms");
    tela.hidden=false;
    const parar=digitarBoasVindas([[document.getElementById("saudacao"),texto.saudacao],[document.getElementById("nomeBoasVindas"),texto.nome],[document.getElementById("mensagemBoasVindas"),texto.mensagem]],texto.duracao,reduzido);
    let encerrada=false,timer;
    const fechar=()=>{if(encerrada)return;encerrada=true;clearTimeout(timer);parar();menu.hidden=false;tela.classList.add("saindo");const finalizar=()=>{tela.hidden=true;menu.inert=false;};if(reduzido)finalizar();else setTimeout(finalizar,800);};
    document.getElementById("pularBoasVindas").onclick=fechar;
    timer=setTimeout(fechar,texto.duracao);
}
