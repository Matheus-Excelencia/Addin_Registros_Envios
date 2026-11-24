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
