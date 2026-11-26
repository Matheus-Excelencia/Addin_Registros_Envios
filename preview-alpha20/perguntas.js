/* PERGUNTAS PERSONALIZADAS CONFIGURADAS NA ABA CONFIG */
let perguntasPersonalizadas = [{nome:"Qual valor total dos processos?", coluna:"Valor", formato:"contabil", tipo:"float", obrigatoria:false}];
function carregarPerguntasPersonalizadas(valores, info) {
    const iPergunta = obterIndice(info.mapa, ["PERGUNTA PERSONALIZADA"]);
    const iFormulario = obterIndice(info.mapa, ["FORMULARIO"]);
    const iTipo = obterIndice(info.mapa, ["TIPO RESPOSTA"]);
    const iObrigatoria = obterIndice(info.mapa, ["OBRIGATORIA"]);
    const iFormato = obterIndice(info.mapa, ["FORMATO EXCEL"]);
    const iColuna = obterIndice(info.mapa, ["COLUNA EXCEL"]);
    const iAtivo = obterIndice(info.mapa, ["PERGUNTA ATIVA"]);
    perguntasPersonalizadas = [{nome:"Qual valor total dos processos?", coluna:"Valor", formato:"contabil", tipo:"float", obrigatoria:false}];
    if ([iPergunta, iFormulario, iTipo, iObrigatoria].some(i => i < 0)) {
        renderizarPerguntasPersonalizadas();
        return;
    }
    for (let linha = info.linha + 1; linha < valores.length; linha++) {
        const r = valores[linha] || [];
        const nome = String(r[iPergunta] ?? "").trim();
        if (!nome) continue;
        if (normalizar(r[iColuna]) === "VALOR" && String(r[iFormulario]).toLowerCase()==="ambos" && String(r[iTipo]).toLowerCase()==="float") {
            const formato=String(r[iFormato]||"contabil").toLowerCase();
            perguntasPersonalizadas[0]={nome,coluna:"Valor",tipo:"float",formato:["geral","decimal","decimal4","contabil","moeda"].includes(formato)?formato:"contabil",obrigatoria:normalizar(r[iObrigatoria])==="SIM"||r[iObrigatoria]===true};
            continue;
        }
        if (iAtivo >= 0 && normalizar(r[iAtivo]) === "NAO") continue;
        const formulario = String(r[iFormulario] || "ambos").toLowerCase();
        if (formulario !== "ambos" && formulario !== TIPO_FORMULARIO) continue;
        const colunaConfigurada = String(r[iColuna] || nome).trim();
        if (["VALOR","QUAL VALOR TOTAL DOS PROCESSOS?","QUAL VALOR TOTAL DOS PROCESSOS","DATA","REALIZADO POR","EMPRESA","QTDE","EMAIL RESPOSTA","E-MAIL RESPOSTA","SUPERVISOR","OBS","OBS.","HISTORICO EXTERNO","HISTORICOEXTERNO","MENSAGEM","ASSUNTO","ID REGISTRO"].includes(normalizar(colunaConfigurada)) || perguntasPersonalizadas.some(q => normalizar(q.coluna) === normalizar(colunaConfigurada))) continue;
        perguntasPersonalizadas.push({
            nome,
            coluna: iColuna < 0 ? (normalizar(nome) === "QUAL VALOR TOTAL DOS PROCESSOS" ? "Valor" : nome) : String(r[iColuna] || (normalizar(nome) === "QUAL VALOR TOTAL DOS PROCESSOS" ? "Valor" : nome)).trim(),
            formato: iFormato < 0 ? (normalizar(nome) === "QUAL VALOR TOTAL DOS PROCESSOS" ? "contabil" : "geral") : String(r[iFormato] || "geral").toLowerCase(),
            tipo: String(r[iTipo] || "text").toLowerCase(),
            obrigatoria: normalizar(r[iObrigatoria]) === "SIM" || r[iObrigatoria] === true
        });
    }
    renderizarPerguntasPersonalizadas();
}
function renderizarPerguntasPersonalizadas() {
    const area = document.getElementById("perguntasPersonalizadas");
    if (!area) return;
    area.innerHTML = "";
    perguntasPersonalizadas.forEach((pergunta, i) => {
        const wrapper = document.createElement("div");
        wrapper.className = "campo";
        const label = document.createElement("label");
        label.htmlFor = "perguntaPersonalizada_" + i;
        label.textContent = pergunta.nome + (pergunta.obrigatoria ? " *" : " (opcional)");
        const input = document.createElement(pergunta.tipo === "text" ? "textarea" : pergunta.tipo === "bool" ? "select" : "input");
        input.id = "perguntaPersonalizada_" + i;
        input.dataset.pergunta = pergunta.nome;
        input.dataset.tipo = pergunta.tipo;
        input.rows = 2;
        if (pergunta.tipo === "bool") {
            [["", "Selecione"], ["Sim", "Sim"], ["Não", "Não"]].forEach(([v,t]) => { const opt=document.createElement("option");opt.value=v;opt.textContent=t;input.appendChild(opt); });
        } else if (["date","time","datetime"].includes(pergunta.tipo)) {
            input.type = pergunta.tipo === "datetime" ? "datetime-local" : pergunta.tipo;
        } else if (pergunta.tipo === "int") {
            input.type = "number";
            input.step = "1";
            input.inputMode = "numeric";
        } else if (pergunta.tipo === "float") {
            input.type = "text";
            input.inputMode = "decimal";
            input.placeholder = "Ex.: 1234,56 ou 1.234,56";
        } else {
            input.placeholder = "Digite sua resposta";
        }
        input.required = pergunta.obrigatoria;
        if (pergunta.obrigatoria) input.setAttribute("aria-required", "true");
        wrapper.appendChild(label);
        wrapper.appendChild(input);
        area.appendChild(wrapper);
    });
}
function coletarRespostasPersonalizadas() {
    const respostas = [];
    perguntasPersonalizadas.forEach((pergunta, i) => {
        const campo = document.getElementById("perguntaPersonalizada_" + i);
        const bruto = campo ? campo.value.trim() : "";
        if (!bruto) {
            if (pergunta.obrigatoria) throw new Error('Preencha o campo "' + pergunta.nome + '".');
            respostas.push({nome: pergunta.coluna, valor: "", formato: pergunta.formato, tipo: pergunta.tipo});
            return;
        }
        if (pergunta.tipo === "int") {
            const numero = Number(bruto);
            if (!Number.isSafeInteger(numero)) throw new Error('O campo "' + pergunta.nome + '" deve ser um número inteiro.');
            respostas.push({nome: pergunta.coluna, valor: numero / (pergunta.formato === "percentual" ? 100 : 1), formato: pergunta.formato, tipo: pergunta.tipo});
        } else if (pergunta.tipo === "float") {
            let normalizado = bruto.replace(/^R\$\s*/i, "");
            if (/^[+-]?\d{1,3}(?:\.\d{3})+$/.test(normalizado)) normalizado = normalizado.replace(/\./g, "");
            if (normalizado.includes(",")) {
                if (!/^[+-]?(?:\d+|\d{1,3}(?:\.\d{3})+),\d+$/.test(normalizado)) throw new Error("Número decimal inválido: " + pergunta.nome);
                normalizado = normalizado.replace(/\./g, "").replace(",", ".");
            }
            if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(normalizado) || !Number.isFinite(Number(normalizado))) {
                throw new Error('O campo "' + pergunta.nome + '" deve ser um número decimal válido. Exemplo: 1.234,56.');
            }
            respostas.push({nome: pergunta.coluna, valor: Number(normalizado) / (pergunta.formato === "percentual" ? 100 : 1), formato: pergunta.formato, tipo: pergunta.tipo});
        } else {
            if (pergunta.tipo === "date" || pergunta.tipo === "datetime") {
                const d = new Date(pergunta.tipo === "date" ? bruto + "T00:00:00" : bruto);
                if (Number.isNaN(d.getTime())) throw new Error('Data inválida em "' + pergunta.nome + '".');
                const serial = (Date.UTC(d.getFullYear(),d.getMonth(),d.getDate(),d.getHours(),d.getMinutes())-Date.UTC(1899,11,30))/86400000;
                respostas.push({nome: pergunta.coluna, valor: serial, formato: pergunta.formato === "geral" ? (pergunta.tipo === "date" ? "data" : "datahora") : pergunta.formato, tipo: pergunta.tipo});
            } else if (pergunta.tipo === "time") {
                if (!/^\d{2}:\d{2}$/.test(bruto)) throw new Error('Hora inválida em "' + pergunta.nome + '".');
                const [h,m]=bruto.split(":").map(Number);
                if(h>23||m>59)throw new Error('Hora inválida em "' + pergunta.nome + '".');
                respostas.push({nome: pergunta.coluna, valor: (h*60+m)/1440, formato: pergunta.formato === "geral" ? "hora" : pergunta.formato, tipo: pergunta.tipo});
            } else respostas.push({nome: pergunta.coluna, valor: pergunta.tipo === "bool" ? bruto === "Sim" : (/^[=+\-@]/.test(bruto) ? "'" + bruto : bruto), formato: pergunta.formato, tipo: pergunta.tipo});
        }
    });
    return respostas;
}
function prepararColunasPersonalizadas(folha, usado, linhaCabecalho, mapa, respostas) {
    let proximaColuna = Math.max(usado.isNullObject ? 0 : usado.values[0].length, Math.max(-1, ...Object.values(mapa).filter(Number.isInteger)) + 1);
    respostas.forEach(function(resposta) {
        let indice = obterIndice(mapa, [resposta.nome]);
        if (indice === -1) {
            if (normalizar(resposta.nome) === "VALOR") { if (resposta.valor === "") return; throw new Error("Coluna Valor ausente; confira a planilha."); }
            indice = proximaColuna++;
            mapa[normalizar(resposta.nome)] = indice;
            folha.getRangeByIndexes(
                (usado.isNullObject ? 0 : usado.rowIndex) + linhaCabecalho,
                (usado.isNullObject ? 0 : usado.columnIndex) + indice,
                1, 1
            ).values = [[resposta.nome]];
        }
    });
}
function formatarDuplicacao(q,v){
 if(v===undefined||v===null||v==="")return "";
 if(q.tipo==="bool")return v===true||String(v).toLowerCase()==="true"||v==="Sim"?"Sim":"Não";
 if(typeof v==="number"){
  if(q.tipo==="date"||q.tipo==="datetime"){const d=new Date(Date.UTC(1899,11,30)+Math.round(v*86400000));return d.toISOString().slice(0,q.tipo==="date"?10:16);}
  if(q.tipo==="time"){const m=Math.round(v*1440)%1440;return String(Math.floor(m/60)).padStart(2,"0")+":"+String(m%60).padStart(2,"0");}
  if(q.formato==="percentual")v*=100;
  if(q.tipo==="float")return String(v).replace(".",",");
 }
 return String(v);
}
