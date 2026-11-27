const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const read=p=>fs.readFileSync(path.join(__dirname,'..',p),'utf8');
const src=read('public/taskpane.js');
function env(s=src){const c=vm.createContext({});vm.runInContext(s,c);return c}
test('Menção e link exclusivos da saudação especial',()=>{const c=env();const special=vm.runInContext('textoBoasVindas({name:"Kamylli Hemer"})',c);assert.match(special.mensagem,/artigo especial/);assert.equal(special.artigo,true);assert.equal(vm.runInContext('textoBoasVindas({name:"Outra pessoa"})',c).artigo,false)});
test('Desativar atalho mantém menção',()=>{const c=env(src.replace('ARTIGO_BOAS_VINDAS_ATIVO = true','ARTIGO_BOAS_VINDAS_ATIVO = false'));const t=vm.runInContext('textoBoasVindas({name:"Kamylli Hemer"})',c);assert.match(t.mensagem,/artigo especial/);assert.equal(t.artigo,false)});
test('Artigo abre via Office somente ao clicar',()=>{const c=env();let url,prevented=false;c.Office={context:{ui:{openBrowserWindow:v=>url=v}}};c.evento={preventDefault(){prevented=true}};vm.runInContext('abrirArtigoBoasVindas(evento)',c);assert.equal(url,'https://addin-registros-envios.vercel.app/artigo.html');assert.equal(prevented,true)});
test('Link inicialmente oculto e fallback seguro',()=>assert.match(read('public/taskpane.html'),/id="artigoBoasVindas" href="https:\/\/addin-registros-envios\.vercel\.app\/artigo\.html" target="_blank" rel="noopener noreferrer" hidden/));
test('Licença corporativa sem concessão aberta',()=>{assert.match(read('LICENSE'),/não constitui concessão de licença de código aberto/);assert.match(read('README.md'),/\]\(LICENSE\)/)});
