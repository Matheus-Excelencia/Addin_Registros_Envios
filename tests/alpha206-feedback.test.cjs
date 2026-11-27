const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path');const read=p=>fs.readFileSync(path.join(__dirname,'../public',p),'utf8');
for(const p of ['email','sms'])test('Validação local '+p,()=>{assert.match(read(p+'.html'),/novalidate/);assert.match(read(p+'.js'),/Feedback\.validate/);assert.match(read(p+'.js'),/Feedback\.error\("qtde"/);assert.match(read(p+'.js'),/Feedback\.showError/)});
test('Erros personalizados identificam campo',()=>assert.match(read('perguntas.js'),/fieldId:"perguntaPersonalizada_"\+i/));
