const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path');
const root=path.join(__dirname,'..'),read=p=>fs.readFileSync(path.join(root,p),'utf8');
test('Cabeçalho aprovado e versão da landing preservados',()=>{const s=read('public/index.html');assert.match(s,/h\.getBoundingClientRect\(\)\.top<=1/);assert.match(s,/Excel Web · Alpha 2\.0/);assert.doesNotMatch(s,/Excel Web · Alpha 1\.06/);});
test('Links antigos da preview redirecionam sem duplicar aplicação',()=>{const c=JSON.parse(read('vercel.json'));assert.ok(c.redirects.some(r=>r.source==='/preview-alpha20/:path*'&&r.destination==='/:path*'&&r.permanent===false));});
test('Plano preserva backup e bloqueio de publicação',()=>{const s=read('docs/PLANO-PUBLICACAO-ALPHA-20.md');assert.match(s,/aa438e4dafdc6d02cd3abe9c39aaa474cf12d91c/);assert.match(s,/dpl_C9eoJ42UFFW6RVw4jKi9pK3Epj2x/);assert.match(s,/NÃO AUTORIZADO PARA PUBLICAÇÃO/);});
