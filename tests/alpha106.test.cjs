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
