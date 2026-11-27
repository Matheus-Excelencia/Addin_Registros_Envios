process.chdir(require('node:path').join(__dirname,'../public'));
const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),{webcrypto}=require('node:crypto');
function env(url='https://tenant.sharepoint.com/workbook.xlsx',demo=false){const map=new Map();let now=100000;const c=vm.createContext({window:{REGISTROS_DEMO:demo},Office:{context:{document:{getFilePropertiesAsync(cb){cb({status:'succeeded',value:{url}})}}}},crypto:webcrypto,TextEncoder,setTimeout,clearTimeout,Date:{now:()=>now},sessionStorage:{getItem:k=>map.get(k)??null,setItem:(k,v)=>map.set(k,v),removeItem:k=>map.delete(k)}});vm.runInContext(fs.readFileSync('registro-security.js','utf8'),c);const api=vm.runInContext('RegistroSecurity',c);return {c,map,api,store:async type=>api.store({tipo:type,empresa:'Exemplo'},await api.scope()),advance:n=>now+=n,change:v=>url=v}}
for(const v of ['=1+1','+5511999999999','-texto','@texto','  =1+1','\t=1+1','\u200b=1+1',"'=1+1","'literal"])test('Texto literal '+JSON.stringify(v),()=>assert.equal(env().api.cell(v),"'"+v));
for(const v of [-123.45,0,42,true,false,'','texto normal','email@example.com'])test('Tipo preservado '+JSON.stringify(v),()=>assert.equal(env().api.cell(v),v));
test('Rejeita não finito',()=>assert.throws(()=>env().api.cell(Infinity)));
test('Mesmo arquivo uma vez',async()=>{const e=env();assert.equal(await e.store('email'),true);assert.equal((await e.api.take('email')).empresa,'Exemplo');assert.equal(await e.api.take('email'),null)});
test('Arquivo diferente mesmo nome',async()=>{const e=env();await e.store('sms');e.change('https://other.sharepoint.com/workbook.xlsx');assert.equal(await e.api.take('sms'),null);assert.equal(e.map.size,0)});
test('Origem do histórico divergente',async()=>{const e=env(),source=await e.api.scope();e.change('https://other.sharepoint.com/workbook.xlsx');assert.equal(await e.api.store({tipo:'email'},source),false)});
test('Sem URL bloqueia',async()=>assert.equal(await env('').store('email'),false));
test('Expiração',async()=>{const e=env();await e.store('email');e.advance(120001);assert.equal(await e.api.take('email'),null)});
test('Canal diferente',async()=>{const e=env();await e.store('sms');assert.equal(await e.api.take('email'),null);assert.equal(e.map.size,0)});
test('Legado descartado',async()=>{const e=env();e.map.set('registroDuplicado',JSON.stringify({tipo:'email'}));assert.equal(await e.api.take('email'),null);assert.equal(e.map.size,0)});
test('Demo e Office separados',async()=>{const e=env('',true);await e.store('email');e.c.window.REGISTROS_DEMO=false;e.change('https://tenant.sharepoint.com/workbook.xlsx');assert.equal(await e.api.take('email'),null)});
test('URL não gravada',async()=>{const e=env();await e.store('email');assert.doesNotMatch(e.map.get('registroDuplicado'),/sharepoint\.com/)});
for(const p of ['email','sms'])test('Gravação protegida '+p,()=>{const s=fs.readFileSync(p+'.js','utf8');assert.match(s,/intervaloRegistro\.values = \[valores\.map\(RegistroSecurity\.cell\)\]/);assert.match(s,/gravarCadastroOrdenado\(folha, usado, info, \[coluna\], \[campo\.valor\]\)/);assert.match(s,/await RegistroSecurity\.take/);assert.doesNotMatch(s,/console\.error\(erro\)/)});
test('Config protege gravações',()=>{
 const s=fs.readFileSync('config.js','utf8'),e=env();
 e.c.Office.onReady=()=>{};e.c.document={};
 vm.runInContext(s,e.c);
 const write=vm.runInContext('gravarCadastroOrdenado',e.c),writes=[];
 const sh={getRangeByIndexes(row,col,height,width){const range={row,col,height,width,format:{}};writes.push(range);return range}};
 const u={rowIndex:2,columnIndex:3,values:[['NOME','TEXTO','OUTRA'],['Zeta','=1+1','preservar'],['','','preservar'],['Água','Olá\nMundo','preservar']]};
 write(sh,u,{linha:0},[0,1],['Beta','@texto']);
 assert.equal(writes.length,2);
 assert.deepEqual(writes.map(r=>r.col),[3,4]);
 assert.equal(JSON.stringify(writes.map(r=>r.values)),JSON.stringify([[['Água'],['Beta'],['Zeta']],[['Olá\nMundo'],["'@texto"],["'=1+1"]]]));
 for(const r of writes){assert.equal(r.row,3);assert.equal(r.height,3);assert.equal(r.width,1);assert.equal(r.format.wrapText,false)}
 writes.length=0;
 write(sh,{rowIndex:0,columnIndex:0,values:[['EMPRESA'],['Zeta'],['Água']]},{linha:0},[0],['Beta']);
 assert.equal(JSON.stringify(writes[0].values),JSON.stringify([['Água'],['Beta'],['Zeta']]));
 assert.equal(writes[0].format.wrapText,false);
 writes.length=0;
 write(sh,{rowIndex:0,columnIndex:0,values:[['EMPRESA']]},{linha:0},[0],['+Beta']);
 assert.equal(JSON.stringify(writes[0].values),JSON.stringify([["'+Beta"]]));
 assert.ok(s.includes('gravarCadastroOrdenado(sh,u,h,[col],[valor])'));
 assert.ok(s.includes('gravarCadastroOrdenado(sh,u,h,[nc,tc],[nome,texto])'));
 for(const v of ['valor','v'])assert.ok(s.includes('[[RegistroSecurity.cell('+v+')]]'));
});
test('Cabeçalho protegido',()=>assert.match(fs.readFileSync('perguntas.js','utf8'),/\[\[RegistroSecurity\.cell\(resposta\.nome\)\]\]/));
for(const p of ['email','sms','config','historico'])test('Helper carregado '+p,()=>{const s=fs.readFileSync(p+'.html','utf8');assert.ok(s.indexOf('registro-security.js')>=0&&s.indexOf('registro-security.js')<s.indexOf(p+'.js?'))});

for(const canal of ['email','sms'])test('Config ordenada pelo formulário '+canal,async()=>{
 const e=env(),writes=[];
 const s=fs.readFileSync(canal+'.js','utf8');
 const nome=canal==='email'?'atualizarConfig':'atualizarConfigSMS';
 const trecho=s.slice(s.indexOf('async function '+nome+'('),s.indexOf('async function preencherDuplicacao'));
 const usado={values:[['EMPRESAS','EMAIL RESPOSTA','SUPERVISORES','REALIZADO POR','OUTRA'],['Zeta','z@exemplo.com','Zoe','Zoe','intacto'],['Água','a@exemplo.com','Ana','Ana','intacto']],rowIndex:2,columnIndex:1,isNullObject:false,load(){}};
 const folha={getUsedRangeOrNullObject(){return usado},getRangeByIndexes(row,col,height,width){const r={row,col,height,width,format:{}};writes.push(r);return r}};
 e.c.context={workbook:{worksheets:{getItem(){return folha}}},sync:async()=>{}};
 e.c.dados={empresa:'Beta',emailResposta:'b@exemplo.com',supervisor:'Bruno',realizadoPor:'Bruno'};
 e.c.encontrarCabecalho=()=>({linha:0,mapa:Object.fromEntries(usado.values[0].map((v,i)=>[v,i]))});
 e.c.obterIndice=(mapa,nomes)=>nomes.map(n=>mapa[n]).find(v=>v!==undefined)??-1;
 e.c.normalizar=v=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toUpperCase();
 vm.runInContext(trecho,e.c);
 await vm.runInContext(nome+'(context,dados)',e.c);
 assert.deepEqual(writes.map(r=>r.col),canal==='email'?[1,2,3,4]:[1,3,4]);
 assert.equal(JSON.stringify(writes[0].values),JSON.stringify([['Água'],['Beta'],['Zeta']]));
 for(const r of writes){assert.equal(r.row,3);assert.equal(r.height,3);assert.equal(r.width,1);assert.equal(r.format.wrapText,false)}
 if(canal==='email')assert.equal(JSON.stringify(writes[1].values),JSON.stringify([['a@exemplo.com'],['b@exemplo.com'],['z@exemplo.com']]));
 assert.equal(JSON.stringify(writes.at(-1).values),JSON.stringify([['Ana'],['Bruno'],['Zoe']]));
 writes.length=0;e.c.dados={empresa:'agua',supervisor:'Ana',realizadoPor:'Ana',emailResposta:'a@exemplo.com'};
 await vm.runInContext(nome+'(context,dados)',e.c);assert.equal(writes.length,0);
});
