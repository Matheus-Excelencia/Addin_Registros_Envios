let configInfo = null;
let mensagensSMS = [];
let statusLog = [];
let limiteSMS = 160;

Office.onReady(async function () {
    document.getElementById("btnVoltar").addEventListener("click", function () {
        window.location.href = "taskpane.html";
    });

    document.getElementById("btnCancelar").addEventListener("click", limparFormulario);
    document.getElementById("formSms").addEventListener("submit", salvarSMS);
    document.getElementById("mensagem").addEventListener("change", mostrarTextoMensagem);

    prepararLog();
    await carregarDados();
});

function prepararLog() {
    const elemento = document.getElementById("mensagemStatus");
    if (!elemento) return;
    elemento.style.whiteSpace = "pre-line";
    elemento.style.fontSize = "12px";
    elemento.style.lineHeight = "1.4";
    elemento.style.padding = "8px 10px";
    elemento.style.marginBottom = "12px";
}

function atualizarLog() {
    const elemento = document.getElementById("mensagemStatus");
    if (!elemento) return;
    elemento.style.whiteSpace = "pre-line";
    elemento.style.fontSize = "12px";
    elemento.style.lineHeight = "1.4";
    elemento.textContent = statusLog.join("\n");
    elemento.className = "status aviso";
}

function adicionarLog(texto) {
    statusLog.push(texto);
    atualizarLog();
}

function limparStatus() {
    statusLog = [];
    const elemento = document.getElementById("mensagemStatus");
    if (!elemento) return;
    elemento.textContent = "";
    elemento.className = "status";
}

function mostrarStatus(texto, tipo) {
    const elemento = document.getElementById("mensagemStatus");
    if (!elemento) return;
    elemento.textContent = texto;
    elemento.className = "status " + tipo;
}

function limparFormulario() {
    document.getElementById("formSms").reset();
    document.getElementById("mensagem").selectedIndex = 0;
    document.getElementById("textoMensagem").value = "";
    atualizarContador();
    limparStatus();
    identificarUsuario();
}

function normalizar(valor) {
    return String(valor || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .toUpperCase();
}

function obterIndice(mapa, nomes) {
    for (const nome of nomes) {
        const chave = normalizar(nome);
        if (mapa[chave] !== undefined) return mapa[chave];
    }
    return -1;
}

function encontrarCabecalho(valores) {
    for (let linha = 0; linha < Math.min(valores.length, 30); linha++) {
        const mapa = {};
        valores[linha].forEach(function (valor, indice) {
            if (valor !== null && valor !== undefined && String(valor).trim() !== "") {
                mapa[normalizar(valor)] = indice;
            }
        });

        const temEmpresa = obterIndice(mapa, ["EMPRESAS", "EMPRESA"]) !== -1;
        const temSupervisor = obterIndice(mapa, ["SUPERVISORES", "SUPERVISOR"]) !== -1;

        if (temEmpresa && temSupervisor) return { linha: linha, mapa: mapa };
    }
    return null;
}

function obterValoresColuna(valores, info, nomes) {
    const indice = obterIndice(info.mapa, Array.isArray(nomes) ? nomes : [nomes]);
    if (indice === -1) return [];
    const resultado = [];
    for (let linha = info.linha + 1; linha < valores.length; linha++) {
        const valor = valores[linha][indice];
        if (valor !== null && valor !== undefined && String(valor).trim() !== "") {
            resultado.push(String(valor).trim());
        }
    }
    return [...new Set(resultado)];
}

function preencherLista(id, valores) {
    const lista = document.getElementById(id);
    if (!lista) return;
    lista.innerHTML = "";
    valores.forEach(function (valor) {
        const option = document.createElement("option");
        option.value = valor;
        lista.appendChild(option);
    });
}

async function carregarDados() {
    limparStatus();
    statusLog = [];
    adicionarLog("● Conectando ao Config...");

    try {
        await Excel.run(async function (context) {
            const folha = context.workbook.worksheets.getItem("Config");
            const usado = folha.getUsedRangeOrNullObject();
            usado.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex", "isNullObject"]);
            await context.sync();

            if (usado.isNullObject) throw new Error("A aba Config está vazia.");

            configInfo = encontrarCabecalho(usado.values);
            if (!configInfo) throw new Error("Cabeçalho do Config não encontrado.");

            const empresas = obterValoresColuna(usado.values, configInfo, ["EMPRESAS", "EMPRESA"]);
            preencherLista("listaEmpresas", empresas);
            adicionarLog(empresas.length > 0 ? "✓ Empresas " + empresas.length : "⚠ Empresas 0");

            const supervisores = obterValoresColuna(usado.values, configInfo, ["SUPERVISORES", "SUPERVISOR"]);
            preencherLista("listaSupervisores", supervisores);
            adicionarLog(supervisores.length > 0 ? "✓ Supervisores " + supervisores.length : "⚠ Supervisores 0");

            const realizadoPor = obterValoresColuna(usado.values, configInfo, ["REALIZADO POR", "REALIZADO_POR", "REALIZADOPOR"]);
            preencherLista("listaRealizadoPor", realizadoPor);
            adicionarLog(realizadoPor.length > 0 ? "✓ Realizado por " + realizadoPor.length : "⚠ Realizado por 0");

            const indiceLimite = obterIndice(configInfo.mapa, ["LIMITE SMS", "LIMITE_SMS"]);
            if (indiceLimite !== -1) {
                for (let linha = configInfo.linha + 1; linha < usado.values.length; linha++) {
                    const numero = Number(usado.values[linha][indiceLimite]);
                    if (Number.isFinite(numero) && numero > 0) {
                        limiteSMS = numero;
                        break;
                    }
                }
            }

            const campoMensagem = document.getElementById("textoMensagem");
            campoMensagem.maxLength = limiteSMS;
            atualizarContador();

            carregarMensagensSMS(usado.values, configInfo);
        });

        await identificarUsuario();
    } catch (erro) {
        console.error(erro);
        adicionarLog("⚠ Erro no Config");
        mostrarStatus(statusLog.join("\n"), "aviso");
        await identificarUsuario();
    }
}

function carregarMensagensSMS(valores, info) {
    const indiceNome = obterIndice(info.mapa, ["NOME MENSAGEM SMS", "NOME MENSAGEM"]);
    const indiceTexto = obterIndice(info.mapa, ["TEXTO MENSAGEM SMS", "TEXTO MENSAGEM"]);
    const select = document.getElementById("mensagem");

    select.innerHTML = "";
    mensagensSMS = [];

    if (indiceNome === -1 || indiceTexto === -1) {
        const option = document.createElement("option");
        option.value = "";
        option.textContent = "Mensagens de SMS não encontradas";
        select.appendChild(option);
        adicionarLog("⚠ Mensagens de SMS 0");
        return;
    }

    const inicial = document.createElement("option");
    inicial.value = "";
    inicial.textContent = "Selecione uma mensagem";
    select.appendChild(inicial);

    for (let linha = info.linha + 1; linha < valores.length; linha++) {
        const nome = valores[linha][indiceNome];
        const texto = valores[linha][indiceTexto];
        if (nome !== null && nome !== undefined && String(nome).trim() !== "") {
            const mensagem = {
                nome: String(nome).trim(),
                texto: String(texto || "")
            };
            mensagensSMS.push(mensagem);
            const option = document.createElement("option");
            option.value = mensagem.nome;
            option.textContent = mensagem.nome;
            select.appendChild(option);
        }
    }

    adicionarLog(mensagensSMS.length > 0 ? "✓ Mensagens SMS " + mensagensSMS.length : "⚠ Mensagens SMS 0");
}

function mostrarTextoMensagem() {
    const nome = document.getElementById("mensagem").value;
    const mensagem = mensagensSMS.find(function (item) { return item.nome === nome; });
    const campo = document.getElementById("textoMensagem");
    campo.value = mensagem ? mensagem.texto : "";
    atualizarContador();
}

function atualizarContador() {
    const campo = document.getElementById("textoMensagem");
    const contador = document.getElementById("contador");
    if (!campo) return;
    const quantidade = campo.value.length;
    if (contador) contador.textContent = quantidade + " / " + limiteSMS;
    campo.style.borderColor = quantidade > limiteSMS ? "red" : "";
}

async function identificarUsuario() {
    const campo = document.getElementById("realizadoPor");
    try {
        const token = await Office.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true
        });
        const dados = decodificarToken(token);
        const nome = dados.name || dados.preferred_username || dados.email || dados.upn;
        if (nome) {
            campo.value = nome;
            campo.readOnly = false;
            adicionarLog("✓ SSO: " + nome);
        } else {
            adicionarLog("⚠ SSO sem nome");
        }
    } catch (erro) {
        console.error(erro);
        campo.readOnly = false;
        adicionarLog("⚠ SSO não identificado");
    }
}

function decodificarToken(token) {
    const partes = token.split(".");
    if (partes.length !== 3) throw new Error("Token inválido.");
    let payload = partes[1].replace(/-/g, "+").replace(/_/g, "/");
    while (payload.length % 4 !== 0) payload += "=";
    return JSON.parse(decodeURIComponent(atob(payload).split("").map(function (c) {
        return "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
    }).join("")));
}

function obterDataHoraAtual() {
    const agora = new Date();
    return String(agora.getDate()).padStart(2, "0") + "/" +
        String(agora.getMonth() + 1).padStart(2, "0") + "/" +
        agora.getFullYear() + " " +
        String(agora.getHours()).padStart(2, "0") + ":" +
        String(agora.getMinutes()).padStart(2, "0") + ":" +
        String(agora.getSeconds()).padStart(2, "0");
}

async function salvarSMS(event) {
    event.preventDefault();

    const btn = document.getElementById("btnSalvar");
    btn.disabled = true;
    btn.textContent = "Salvando...";
    limparStatus();

    try {
        const realizadoPor = document.getElementById("realizadoPor").value.trim();
        const empresa = document.getElementById("empresa").value.trim();
        const qtde = document.getElementById("qtde").value.trim();
        const supervisor = document.getElementById("supervisor").value.trim();
        const obs = document.getElementById("obs").value.trim();
        const nomeMensagem = document.getElementById("mensagem").value;
        const textoMensagem = document.getElementById("textoMensagem").value;

        if (!empresa) throw new Error("Informe a empresa.");
        if (!qtde || Number(qtde) <= 0 || !Number.isInteger(Number(qtde))) throw new Error("A Qtde deve ser um número inteiro maior que zero.");
        if (!supervisor) throw new Error("Informe o Supervisor.");
        if (!nomeMensagem) throw new Error("Selecione uma mensagem.");
        if (!textoMensagem) throw new Error("A mensagem selecionada não possui texto.");
        if (textoMensagem.length > limiteSMS) throw new Error("A mensagem possui " + textoMensagem.length + " caracteres. O limite é " + limiteSMS + ".");

        await Excel.run(async function (context) {
            await adicionarRegistroSMS(context, {
                realizadoPor,
                empresa,
                qtde: Number(qtde),
                supervisor,
                obs,
                textoMensagem
            });

            await atualizarConfigSMS(context, {
                empresa,
                supervisor,
                realizadoPor
            });

            await context.sync();
        });

        mostrarStatus("✓ Registro de SMS salvo!", "sucesso");
        document.getElementById("formSms").reset();
        document.getElementById("mensagem").selectedIndex = 0;
        document.getElementById("textoMensagem").value = "";
        atualizarContador();
        await identificarUsuario();
    } catch (erro) {
        console.error(erro);
        mostrarStatus(erro.message || "Erro ao salvar.", "erro");
    } finally {
        btn.disabled = false;
        btn.textContent = "Salvar";
    }
}

async function adicionarRegistroSMS(context, dados) {
    const folha = context.workbook.worksheets.getItem("SMS");
    const usado = folha.getUsedRangeOrNullObject();
    usado.load(["values", "rowCount", "columnCount", "isNullObject", "rowIndex"]);
    await context.sync();

    let linhaCabecalho = -1;
    let mapa = {};

    if (!usado.isNullObject) {
        for (let linha = 0; linha < Math.min(usado.values.length, 20); linha++) {
            const atual = {};
            usado.values[linha].forEach(function (valor, indice) {
                if (valor !== null && valor !== undefined && String(valor).trim() !== "") {
                    atual[normalizar(valor)] = indice;
                }
            });
            if (atual["EMPRESA"] !== undefined && (atual["DATA"] !== undefined || atual["REALIZADO POR"] !== undefined)) {
                linhaCabecalho = linha;
                mapa = atual;
                break;
            }
        }
    }

    if (linhaCabecalho === -1) {
        folha.getRange("A1:G1").values = [["Data", "Realizado por", "Empresa", "Qtde", "Supervisor", "Obs.", "Mensagem"]];
        linhaCabecalho = 0;
        ["DATA", "REALIZADO POR", "EMPRESA", "QTDE", "SUPERVISOR", "OBS.", "MENSAGEM"].forEach(function (nome, indice) {
            mapa[nome] = indice;
        });
    }

    const numeroColunas = Math.max(7, usado.isNullObject ? 7 : usado.columnCount);
    const valores = new Array(numeroColunas).fill("");

    function colocar(nomes, valor) {
        const indice = obterIndice(mapa, nomes);
        if (indice !== -1) valores[indice] = valor;
    }

    colocar(["DATA"], obterDataHoraAtual());
    colocar(["REALIZADO POR"], dados.realizadoPor);
    colocar(["EMPRESA"], dados.empresa);
    colocar(["QTDE"], dados.qtde);
    colocar(["SUPERVISOR"], dados.supervisor);
    colocar(["OBS", "OBS."], dados.obs);
    colocar(["MENSAGEM"], dados.textoMensagem);

    const proximaLinha = usado.isNullObject ? 1 : usado.rowIndex + usado.rowCount;
    folha.getRangeByIndexes(proximaLinha, 0, 1, numeroColunas).values = [valores];
}

async function atualizarConfigSMS(context, dados) {
    const folha = context.workbook.worksheets.getItem("Config");
    const usado = folha.getUsedRangeOrNullObject();
    usado.load(["values", "rowCount", "isNullObject", "rowIndex", "columnIndex"]);
    await context.sync();
    if (usado.isNullObject) return;

    const info = encontrarCabecalho(usado.values);
    if (!info) return;

    const campos = [
        { nomes: ["EMPRESAS", "EMPRESA"], valor: dados.empresa },
        { nomes: ["SUPERVISORES", "SUPERVISOR"], valor: dados.supervisor },
        { nomes: ["REALIZADO POR", "REALIZADO_POR", "REALIZADOPOR"], valor: dados.realizadoPor }
    ];

    for (const campo of campos) {
        if (!campo.valor) continue;
        const coluna = obterIndice(info.mapa, campo.nomes);
        if (coluna === -1) continue;

        let existe = false;
        for (let linha = info.linha + 1; linha < usado.values.length; linha++) {
            if (normalizar(usado.values[linha][coluna]) === normalizar(campo.valor)) {
                existe = true;
                break;
            }
        }
        if (existe) continue;

        let linhaDestino = info.linha + 1;
        while (linhaDestino < usado.values.length && String(usado.values[linhaDestino][coluna] || "").trim() !== "") {
            linhaDestino++;
        }

        folha.getRangeByIndexes(usado.rowIndex + linhaDestino, coluna, 1, 1).values = [[campo.valor]];
    }
}
