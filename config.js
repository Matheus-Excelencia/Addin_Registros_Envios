const FORMATOS_COMPATIVEIS={text:["geral","texto"],int:["geral","inteiro","decimal","decimal4","contabil","moeda","percentual"],float:["geral","decimal","decimal4","contabil","moeda","percentual"],date:["geral","data"],time:["geral","hora"],datetime:["geral","datahora"],bool:["geral"]};
function atualizarFormatos(){const t=document.getElementById("perguntaTipo"),f=document.getElementById("perguntaFormato");const a=FORMATOS_COMPATIVEIS[t.value]||[];for(const o of f.options)o.disabled=!a.includes(o.value);if(!a.includes(f.value))f.value=a[0]||"geral";}
let cadastrandoPergunta=false;
let info=null;
const COLUNAS_PERGUNTAS=["PERGUNTA PERSONALIZADA","FORMULARIO","TIPO RESPOSTA","OBRIGATORIA","FORMATO EXCEL","COLUNA EXCEL"];
Office.onReady(async function(){
 document.getElementById("btnVoltar").onclick=()=>location.href="taskpane.html";
 document.getElementById("btnCadastro").onclick=adicionarCadastro;
 document.getElementById("btnMensagem").onclick=adicionarMensagem;
 document.getElementById("btnLimite").onclick=salvarLimite;
 document.getElementById("btnPergunta").onclick=adicionarPergunta;
 document.getElementById("perguntaTipo").onchange=atualizarFormatos; atualizarFormatos();
 await carregar();
});
function n(v){return String(v??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toUpperCase()}
function idx(m,a){for(const x of a){const k=n(x);if(m[k]!==undefined)return m[k]}return -1}
function head(v){for(let i=0;i<Math.min(v.length,30);i++){const m={};(v[i]||[]).forEach((x,j)=>{if(String(x??"").trim())m[n(x)]=j});if(idx(m,["EMPRESAS","EMPRESA"])!==-1&&idx(m,["SUPERVISORES","SUPERVISOR"])!==-1)return{linha:i,mapa:m}}return null}
function status(t,ok=true){const e=document.getElementById("status");e.textContent=t;e.className="status "+(ok?"sucesso":"erro")}
function emailValido(v){return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v||"").trim())}
async function carregar(){
 status("● Carregando configurações...",true);
 try{
  await Excel.run(async context=>{
   const sh=context.workbook.worksheets.getItem("Config"),u=sh.getUsedRangeOrNullObject(true);
   u.load(["values","isNullObject","rowIndex","columnIndex"]);await context.sync();
   if(u.isNullObject)throw new Error("A aba Config está vazia.");
   info=head(u.values);if(!info)throw new Error("Cabeçalho do Config não encontrado.");
   const il=idx(info.mapa,["LIMITE SMS","LIMITE_SMS"]);let lim=160;
   if(il!==-1)for(let i=info.linha+1;i<u.values.length;i++){const x=Number(u.values[i][il]);if(Number.isInteger(x)&&x>0){lim=x;break}}
   document.getElementById("limite").value=lim;
   const perguntas=lerPerguntas(u.values, u.rowIndex||0, info);
   renderizarPerguntas(perguntas);
  });
  status("✓ Configurações carregadas.",true);
 }catch(e){status(e.message||"Erro ao carregar.",false)}
}
function indicesPerguntas(mapa){return {
 pergunta:idx(mapa,["PERGUNTA PERSONALIZADA"]),
 formulario:idx(mapa,["FORMULARIO"]),
 tipo:idx(mapa,["TIPO RESPOSTA"]),
 obrigatoria:idx(mapa,["OBRIGATORIA"]),
 formato:idx(mapa,["FORMATO EXCEL"]),
 coluna:idx(mapa,["COLUNA EXCEL"])
}}
function lerPerguntas(valores,base,info){
 const ix=indicesPerguntas(info.mapa);if([ix.pergunta,ix.formulario,ix.tipo,ix.obrigatoria].some(v=>v<0))return [];
 const resultado=[];
 for(let i=info.linha+1;i<valores.length;i++){
  const r=valores[i]||[],nome=String(r[ix.pergunta]??"").trim();
  if(!nome)continue;
  resultado.push({nome,formulario:String(r[ix.formulario]||"ambos").toLowerCase(),tipo:String(r[ix.tipo]||"text").toLowerCase(),obrigatoria:n(r[ix.obrigatoria])==="SIM"||r[ix.obrigatoria]===true,linha:base+i,formato:ix.formato<0?"geral":String(r[ix.formato]||"geral"),coluna:ix.coluna<0?nome:String(r[ix.coluna]||nome)});
 }
 return resultado;
}
function renderizarPerguntas(perguntas){
 const area=document.getElementById("listaPerguntas");area.innerHTML="";
 perguntas=[{nome:"Qual valor total dos processos?",formulario:"ambos",tipo:"float",formato:"contabil",coluna:"Valor",obrigatoria:false},...perguntas.filter(q=>n(q.coluna)!=="VALOR" && !/^QUAL VALOR TOTAL DOS PROCESSOS\??$/.test(n(q.nome)))];
 perguntas.forEach(q=>{
  const item=document.createElement("div");item.className="pergunta-item";
  const nome=document.createElement("strong");nome.textContent=q.nome;
  const detalhe=document.createElement("small");detalhe.textContent=(q.formulario==="ambos"?"E-mail e SMS":q.formulario==="email"?"Somente E-mail":"Somente SMS")+" · "+q.tipo+" · "+q.formato+" · Coluna: "+q.coluna+" · "+(q.obrigatoria?"Obrigatória":"Opcional");
  item.appendChild(nome);item.appendChild(detalhe);area.appendChild(item);
 });
}
async function obterEstruturaConfig(context){
 const sh=context.workbook.worksheets.getItem("Config"),u=sh.getUsedRangeOrNullObject(true);
 u.load(["values","isNullObject","rowIndex","columnIndex"]);await context.sync();
 if(u.isNullObject)throw new Error("Config vazia.");
 const h=head(u.values);if(!h)throw new Error("Cabeçalho não encontrado.");
 let cols=indicesPerguntas(h.mapa);
 const existentes=Object.values(h.mapa).filter(Number.isInteger);
 let proxima=Math.max(u.values[0].length-1,...existentes)+1;
 for(const titulo of COLUNAS_PERGUNTAS){
  const chave=n(titulo);
  if(h.mapa[chave]===undefined){
   h.mapa[chave]=proxima;
   sh.getRangeByIndexes(u.rowIndex+h.linha,u.columnIndex+proxima,1,1).values=[[titulo]];
   proxima++;
  }
 }
 cols=indicesPerguntas(h.mapa);
 return {sh,u,h,cols};
}
async function adicionarPergunta(){
 if(cadastrandoPergunta)return;
 cadastrandoPergunta=true; const botao=document.getElementById("btnPergunta");botao.disabled=true;
 try {
 const nome=document.getElementById("perguntaNome").value.trim();
 const formulario=document.getElementById("perguntaFormulario").value;
 const tipo=document.getElementById("perguntaTipo").value;
 const obrigatoria=document.getElementById("perguntaObrigatoria").value;
 const formato=document.getElementById("perguntaFormato").value;
 const coluna=document.getElementById("perguntaColuna").value.trim() || (n(nome)==="QUAL VALOR TOTAL DOS PROCESSOS"?"Valor":nome);
 if(!nome)return status("Digite o texto da pergunta.",false);
 if(!FORMATOS_COMPATIVEIS[tipo]?.includes(formato))return status("Formato incompatível com o tipo.",false);
 if(!["email","sms","ambos"].includes(formulario))return status("Formulário inválido.",false);
 if(n(coluna)==="VALOR"&&n(nome)!=="QUAL VALOR TOTAL DOS PROCESSOS")return status("A coluna Valor é reservada para a pergunta sobre o total dos processos.",false);
 const reservados=["DATA","REALIZADO POR","EMPRESA","QTDE","EMAIL RESPOSTA","E-MAIL RESPOSTA","SUPERVISOR","OBS","OBS.","HISTORICO EXTERNO","HISTÓRICO EXTERNO","MENSAGEM","ASSUNTO","ID REGISTRO","VALOR"];
 if(reservados.includes(n(nome)) || reservados.includes(n(coluna)) || /^[=+\-@]/.test(nome) || /^[=+\-@]/.test(coluna))return status("Esse nome é reservado por um campo existente do registro. Escolha outro texto.",false);
 if(!["text","int","float","date","time","datetime","bool"].includes(tipo))return status("Tipo de resposta inválido.",false);
 try{
  await Excel.run(async context=>{
   const {sh,u,h,cols}=await obterEstruturaConfig(context);
   for(let i=h.linha+1;i<u.values.length;i++){
    if(n(u.values[i][cols.pergunta])===n(nome))throw new Error("Já existe uma pergunta com esse nome.");
    if(n(u.values[i][cols.coluna])===n(coluna)&&String(u.values[i][cols.coluna]??"").trim())throw new Error("Essa coluna já está vinculada a outra pergunta.");
   }
   let linha=h.linha+1;
   while(linha<u.values.length&&String(u.values[linha][cols.pergunta]??"").trim())linha++;
   for(const [indice,valor] of [[cols.pergunta,nome],[cols.formulario,formulario],[cols.tipo,tipo],[cols.obrigatoria,obrigatoria],[cols.formato,formato],[cols.coluna,coluna]])sh.getRangeByIndexes(u.rowIndex+linha,u.columnIndex+indice,1,1).values=[[valor]];
   await context.sync();
  });
  document.getElementById("perguntaNome").value="";
  status("✓ Pergunta adicionada. Ela aparecerá no(s) formulário(s) selecionado(s).");
  await carregar();
 }catch(e){status(e.message||"Erro ao adicionar pergunta.",false)}
 } finally {cadastrandoPergunta=false;botao.disabled=false;}
}
async function adicionarCadastro(){
 const valor=document.getElementById("valorCadastro").value.trim(),tipo=document.getElementById("tipoCadastro").value;
 if(!valor)return status("Informe o valor.",false);
 if(tipo==="EMAIL RESPOSTA"&&!emailValido(valor))return status("⚠ Informe um e-mail válido. Exemplo: nome@empresa.com",false);
 try{await Excel.run(async context=>{
  const sh=context.workbook.worksheets.getItem("Config"),u=sh.getUsedRangeOrNullObject(true);
  u.load(["values","isNullObject","rowIndex","columnIndex"]);await context.sync();
  if(u.isNullObject)throw new Error("Config vazia.");
  const h=head(u.values),col=idx(h.mapa,[tipo]);if(col===-1)throw new Error("Coluna não encontrada.");
  for(let i=h.linha+1;i<u.values.length;i++)if(n(u.values[i][col])===n(valor))throw new Error("Esse valor já está cadastrado.");
  let r=h.linha+1;while(r<u.values.length&&String(u.values[r][col]??"").trim())r++;
  sh.getRangeByIndexes(u.rowIndex+r,u.columnIndex+col,1,1).values=[[valor]];await context.sync();
 });document.getElementById("valorCadastro").value="";status("✓ Cadastro adicionado.")}catch(e){status(e.message||"Erro ao adicionar.",false)}
}
async function adicionarMensagem(){
 const tipo=document.getElementById("tipoMensagem").value,nome=document.getElementById("nomeMensagem").value.trim(),texto=document.getElementById("textoMensagem").value;
 if(!nome||!texto)return status("Informe o nome e o texto da mensagem.",false);
 try{await Excel.run(async context=>{
  const sh=context.workbook.worksheets.getItem("Config"),u=sh.getUsedRangeOrNullObject(true);
  u.load(["values","isNullObject","rowIndex","columnIndex"]);await context.sync();
  const h=head(u.values);if(!h)throw new Error("Cabeçalho não encontrado.");
  const nc=idx(h.mapa,[tipo==="email"?"NOME MENSAGEM EMAIL":"NOME MENSAGEM SMS"]),tc=idx(h.mapa,[tipo==="email"?"TEXTO MENSAGEM EMAIL":"TEXTO MENSAGEM SMS"]);
  if(nc===-1||tc===-1)throw new Error("Colunas de mensagem não encontradas.");
  for(let i=h.linha+1;i<u.values.length;i++)if(n(u.values[i][nc])===n(nome))throw new Error("Esse nome de mensagem já existe.");
  let r=h.linha+1;while(r<u.values.length&&(String(u.values[r][nc]??"").trim()||String(u.values[r][tc]??"").trim()))r++;
  sh.getRangeByIndexes(u.rowIndex+r,u.columnIndex+nc,1,1).values=[[nome]];
  sh.getRangeByIndexes(u.rowIndex+r,u.columnIndex+tc,1,1).values=[[texto]];await context.sync();
 });document.getElementById("nomeMensagem").value="";document.getElementById("textoMensagem").value="";status("✓ Mensagem adicionada.")}catch(e){status(e.message||"Erro ao adicionar.",false)}
}
async function salvarLimite(){
 const limite=Number(document.getElementById("limite").value);
 if(!Number.isInteger(limite)||limite<=0)return status("O limite deve ser um inteiro maior que zero.",false);
 try{await Excel.run(async context=>{
  const sh=context.workbook.worksheets.getItem("Config"),u=sh.getUsedRangeOrNullObject(true);
  u.load(["values","isNullObject","rowIndex","columnIndex"]);await context.sync();
  const h=head(u.values),col=idx(h.mapa,["LIMITE SMS","LIMITE_SMS"]);if(col===-1)throw new Error("Coluna LIMITE SMS não encontrada.");
  let r=h.linha+1;while(r<u.values.length&&!Number.isInteger(Number(u.values[r][col])))r++;
  sh.getRangeByIndexes(u.rowIndex+r,u.columnIndex+col,1,1).values=[[limite]];await context.sync();
 });status("✓ Limite SMS atualizado.")}catch(e){status(e.message||"Erro ao salvar.",false)}
}
