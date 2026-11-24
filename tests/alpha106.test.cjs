const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),vm=require('node:vm');
const source=fs.readFileSync('perguntas.js','utf8');
function env(value){
 const c=vm.createContext({document:{getElementById:()=>({value})},normalizar:v=>String(v??'').toUpperCase(),obterIndice:(m,a)=>m[a[0].toUpperCase()]??-1});
 vm.runInContext(source,c);return c;
}
for(const [input,expected] of [['', ''],['0',0],['R$ 1.234,56',1234.56],['1234.56',1234.56],['1.234',1234]])test('Valor '+JSON.stringify(input),()=>assert.equal(vm.runInContext('coletarRespostasPersonalizadas()[0].valor',env(input)),expected));
for(const input of ['1.2,34','12 34,56','abc','1,2,3'])test('Rejeita '+input,()=>assert.throws(()=>vm.runInContext('coletarRespostasPersonalizadas()',env(input))));
test('Duplicação monetária',()=>assert.equal(vm.runInContext('formatarDuplicacao({tipo:"float"},1234.56)',env('')),'1234,56'));
test('Duplicação percentual',()=>assert.equal(vm.runInContext('formatarDuplicacao({tipo:"float",formato:"percentual"},0.25)',env('')),'25'));
test('Duplicação booleana',()=>assert.equal(vm.runInContext('formatarDuplicacao({tipo:"bool"},false)',env('')),'Não'));
test('Coluna Valor ausente e opcional não cria cabeçalho',()=>vm.runInContext('prepararColunasPersonalizadas({}, {isNullObject:false,values:[["DATA"]]},0,{DATA:0},[{nome:"Valor",valor:""}])',env('')));
test('Coluna Valor ausente com resposta falha',()=>assert.throws(()=>vm.runInContext('prepararColunasPersonalizadas({}, {isNullObject:false,values:[["DATA"]]},0,{DATA:0},[{nome:"Valor",valor:1}])',env(''))));

function configEnv(){
 const fields=Object.fromEntries(Object.entries({perguntaNome:'',perguntaFormulario:'ambos',perguntaTipo:'float',perguntaFormato:'contabil',perguntaObrigatoria:'nao',perguntaColuna:'',btnPergunta:''}).map(([k,value])=>[k,{value,options:[],after(){}}]));
 const c=vm.createContext({Office:{onReady(){}},document:{getElementById:id=>fields[id],createElement:()=>({})}});
 vm.runInContext(fs.readFileSync('config.js','utf8'),c);return {c,fields};
}
test('Edição preserva coluna, tipo e destino',()=>{
 const {c,fields}=configEnv();
 vm.runInContext('iniciarEdicao({nome:"Original",coluna:"Total",tipo:"float",formato:"contabil",formulario:"ambos",obrigatoria:true})',c);
 assert.equal(fields.perguntaColuna.value,'Total');
 for(const id of ['perguntaColuna','perguntaTipo','perguntaFormulario'])assert.equal(fields[id].disabled,true);
 assert.equal(fields.perguntaObrigatoria.value,'sim');
});
test('Configuração antiga permanece ativa',()=>{
 const {c}=configEnv();
 assert.equal(vm.runInContext('lerPerguntas([[],["Teste","ambos","text","SIM"]],0,{linha:0,mapa:{"PERGUNTA PERSONALIZADA":0,FORMULARIO:1,"TIPO RESPOSTA":2,OBRIGATORIA:3}})[0].ativo',c),true);
});
test('Configuração desativada permanece identificável',()=>{
 const {c}=configEnv();
 assert.equal(vm.runInContext('lerPerguntas([[],["Teste","ambos","text","SIM","NAO"]],0,{linha:0,mapa:{"PERGUNTA PERSONALIZADA":0,FORMULARIO:1,"TIPO RESPOSTA":2,OBRIGATORIA:3,"PERGUNTA ATIVA":4}})[0].ativo',c),false);
});
test('Coluna existente não é recriada',()=>vm.runInContext('prepararColunasPersonalizadas({}, {isNullObject:false,values:[["DATA","Valor"]]},0,{DATA:0,VALOR:1},[{nome:"Valor",valor:123}])',env('')));
test('Campo obrigatório rejeita vazio',()=>assert.throws(()=>vm.runInContext('perguntasPersonalizadas[0].obrigatoria=true;coletarRespostasPersonalizadas()',env(''))));
test('Data duplicada',()=>assert.equal(vm.runInContext('formatarDuplicacao({tipo:"date"},45292)',env('')),'2024-01-01'));
test('Hora duplicada',()=>assert.equal(vm.runInContext('formatarDuplicacao({tipo:"time"},0.5)',env('')),'12:00'));
test('Novo cabeçalho fica após toda largura utilizada',()=>{
 const c=env('');assert.equal(vm.runInContext('let pos;prepararColunasPersonalizadas({getRangeByIndexes:(r,j)=>{pos=j;return {}}},{isNullObject:false,rowIndex:2,columnIndex:3,values:[["DATA","",""]]},0,{DATA:0},[{nome:"Extra",valor:1}]);pos',c),6);
});

test('UX001 aviso local de duplicidade',()=>{const {c}=configEnv();const fields={status:{},statusPergunta:{}};c.document.getElementById=id=>fields[id];vm.runInContext('status("A coluna Saldo já existe",false)',c);assert.match(fields.statusPergunta.textContent,/Saldo/);});
test('UX002 Valor configurado em ambos formulários',()=>{for(const tipo of ['sms','email']){const c=env('');c.TIPO_FORMULARIO=tipo;c.document.getElementById=()=>null;vm.runInContext('carregarPerguntasPersonalizadas([[],["Total atualizado","ambos","float","SIM","moeda","Valor","SIM"]],{linha:0,mapa:{"PERGUNTA PERSONALIZADA":0,FORMULARIO:1,"TIPO RESPOSTA":2,OBRIGATORIA:3,"FORMATO EXCEL":4,"COLUNA EXCEL":5,"PERGUNTA ATIVA":6}})',c);assert.equal(vm.runInContext('perguntasPersonalizadas[0].nome',c),'Total atualizado');assert.equal(vm.runInContext('perguntasPersonalizadas[0].obrigatoria',c),true);assert.equal(vm.runInContext('perguntasPersonalizadas.length',c),1);}});
test('UX003 interface de edição e estilos',()=>{assert.match(fs.readFileSync('config.html','utf8'),/id="editorPergunta"/);assert.match(fs.readFileSync('config.css','utf8'),/\.pergunta-item/);});
test('UX004 contador respeita limite e troca de tipo',()=>{const {c}=configEnv();const fields={tipoMensagem:{value:'sms'},textoMensagem:{value:'123456'},contadorMensagemSMS:{},btnMensagem:{}};c.document.getElementById=id=>fields[id];vm.runInContext('limiteMensagemSMS=5;atualizarContadorSMS()',c);assert.equal(fields.btnMensagem.disabled,true);assert.match(fields.contadorMensagemSMS.textContent,/6 \/ 5/);fields.textoMensagem.value='12345';vm.runInContext('atualizarContadorSMS()',c);assert.equal(fields.btnMensagem.disabled,false);fields.tipoMensagem.value='email';vm.runInContext('atualizarContadorSMS()',c);assert.equal(fields.contadorMensagemSMS.hidden,true);});

function splashEnv(){const c=vm.createContext({Office:{onReady(){}},atob});vm.runInContext(fs.readFileSync("taskpane.js","utf8"),c);return c;}
test("Splash saudação comum",()=>{const c=splashEnv();assert.equal(vm.runInContext('textoBoasVindas({name:"Matheus Passos"}).nome',c),"Matheus Passos");assert.equal(vm.runInContext('textoBoasVindas({name:"Matheus Passos"}).mensagem',c),"");});
test("Splash saudação especial apenas nome exato",()=>{const c=splashEnv();assert.equal(vm.runInContext('textoBoasVindas({name:" Kamylli Hemer "}).saudacao',c),"Bem-vinda");assert.match(vm.runInContext('textoBoasVindas({name:"Kamylli Hemer"}).mensagem',c),/Maionese/);assert.equal(vm.runInContext('textoBoasVindas({name:"Outra Kamylli Hemer"}).mensagem',c),"");});
test("Splash fallback sem SSO",()=>assert.equal(vm.runInContext('textoBoasVindas().nome',splashEnv()),""));
test("Splash decodifica nome Unicode",()=>{const c=splashEnv();c.token="x."+Buffer.from(JSON.stringify({name:"João"})).toString("base64url")+".x";assert.equal(vm.runInContext('dadosBoasVindas(token).name',c),"João");});
