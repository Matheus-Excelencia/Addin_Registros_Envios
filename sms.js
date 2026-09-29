let configDados = {
    empresas: [],
    supervisores: [],
    realizadoPor: [],
    mensagens: [],
    textosMensagens: {},
    limiteSMS: 160
};

let nomeUsuarioSSO = "";

// ======================================================
// INICIALIZAÇÃO
// ======================================================

Office.onReady(async function () {
    const btnVoltar = document.getElementById("btnVoltar");
    const btnCancelar = document.getElementById("btnCancelar");
    const formSMS = document.getElementById("formSMS");
    const mensagem = document.getElementById("mensagem");

    if (btnVoltar) {
        btnVoltar.addEventListener("click", function () {
            window.location.href = "taskpane.html";
        });
    }

    if (btnCancelar) {
        btnCancelar.addEventListener("click", function () {
            window.location.href = "taskpane.html";
        });
    }

    if (mensagem) {
        mensagem.addEventListener("change", function () {
            atualizarTextoMensagem();
        });
    }

    if (formSMS) {
        formSMS.addEventListener("submit", salvarSMS);
    }

    criarLogInicial();

    await iniciarSMS();
});


// ======================================================
// INÍCIO
// ======================================================

async function iniciarSMS() {

    logStatus("Iniciando formulário SMS...", "info");

    try {

        await carregarSSO();

        await carregarConfig();

        atualizarLogFinal();

    } catch (erro) {

        console.error("Erro ao iniciar SMS:", erro);

        logStatus(
            "Erro ao carregar o formulário. Veja o console para detalhes.",
            "erro"
        );
    }
}


// ======================================================
// LOG
// ======================================================

function criarLogInicial() {

    const status = document.getElementById("mensagemStatus");

    if (!status) {
        console.warn("Elemento mensagemStatus não encontrado no sms.html.");
        return;
    }

    status.style.display = "block";
    status.style.fontSize = "11px";
    status.style.lineHeight = "1.5";
    status.style.padding = "8px 10px";
    status.style.marginBottom = "12px";
    status.style.borderRadius = "6px";
    status.style.background = "#f3f4f6";
    status.style.color = "#555";

    status.innerHTML = "Iniciando...";
}


function logStatus(texto, tipo = "info") {

    const status = document.getElementById("mensagemStatus");

    if (!status) {
        console.log("[SMS]", texto);
        return;
    }

    status.style.display = "block";
    status.style.fontSize = "11px";
    status.style.lineHeight = "1.5";
    status.style.padding = "8px 10px";
    status.style.marginBottom = "12px";
    status.style.borderRadius = "6px";

    if (tipo === "erro") {
        status.style.background = "#fdeaea";
        status.style.color = "#a12626";
    } else if (tipo === "sucesso") {
        status.style.background = "#e7f5ec";
        status.style.color = "#1d6b3b";
    } else if (tipo === "aviso") {
        status.style.background = "#fff5d9";
        status.style.color = "#7a5b00";
    } else {
        status.style.background = "#f3f4f6";
        status.style.color = "#555";
    }

    status.innerHTML = texto;
}


function atualizarLogFinal() {

    const sso = nomeUsuarioSSO
        ? `✓ SSO: ${escapeHtml(nomeUsuarioSSO)}`
        : "⚠ SSO não identificado";

    const empresas =
        `✓ Empresas: ${configDados.empresas.length}`;

    const supervisores =
        `✓ Supervisores: ${configDados.supervisores.length}`;

    const realizadoPor =
        `✓ Realizado por: ${configDados.realizadoPor.length}`;

    const mensagens =
        `✓ Mensagens SMS: ${configDados.mensagens.length}`;

    const limite =
        `✓ Limite SMS: ${configDados.limiteSMS}`;

    logStatus(
        `${sso}<br>` +
        `${empresas} &nbsp;&nbsp; ${supervisores}<br>` +
        `${realizadoPor} &nbsp;&nbsp; ${mensagens} &nbsp;&nbsp; ${limite}`,
        "sucesso"
    );
}


// ======================================================
// SSO
// ======================================================

async function carregarSSO() {

    try {

        if (!Office.context.requirements.isSetSupported("IdentityAPI", "1.3")) {

            logStatus(
                "⚠ Identity API não disponível. O campo Realizado por ficará editável.",
                "aviso"
            );

            return;
        }

        const token = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true
        });

        if (!token) {
            throw new Error("Token não retornado.");
        }

        const partes = token.split(".");

        if (partes.length < 2) {
            throw new Error("Token inválido.");
        }

        const payload = JSON.parse(
            decodeBase64Url(partes[1])
        );

        nomeUsuarioSSO =
            payload.name ||
            payload.preferred_username ||
            payload.unique_name ||
            "";

        if (!nomeUsuarioSSO) {
            throw new Error("Nome do usuário não encontrado no token.");
        }

        const campo = document.getElementById("realizadoPor");

        if (campo) {
            campo.value = nomeUsuarioSSO;
        }

        console.log("SSO encontrado:", nomeUsuarioSSO);

    } catch (erro) {

        console.error("Erro SSO:", erro);

        nomeUsuarioSSO = "";

        const campo = document.getElementById("realizadoPor");

        if (campo) {
            campo.value = "";
        }

        logStatus(
            "⚠ SSO não conseguiu identificar o usuário. O campo Realizado por pode ser preenchido manualmente.",
            "aviso"
        );
    }
}


function decodeBase64Url(str) {

    str = str.replace(/-/g, "+").replace(/_/g, "/");

    while (str.length % 4) {
        str += "=";
    }

    const decoded = atob(str);

    const bytes = new Uint8Array(decoded.length);

    for (let i = 0; i < decoded.length; i++) {
        bytes[i] = decoded.charCodeAt(i);
    }

    return new TextDecoder("utf-8").decode(bytes);
}


// ======================================================
// CONFIG
// ======================================================

async function carregarConfig() {

    const empresas = document.getElementById("empresa");
    const supervisores = document.getElementById("supervisor");
    const realizadoPor = document.getElementById("realizadoPor");
    const mensagem = document.getElementById("mensagem");

    try {

        await Excel.run(async function (context) {

            const planilha = context.workbook.worksheets.getItem("Config");

            const usado = planilha.getUsedRangeOrNullObject(true);

            usado.load([
                "isNullObject",
                "values",
                "rowIndex",
                "columnIndex",
                "rowCount",
                "columnCount"
            ]);

            await context.sync();

            if (usado.isNullObject) {
                throw new Error("A aba Config está vazia.");
            }

            const valores = usado.values;

            console.log("Config encontrada:", valores);

            const cabecalho = localizarCabecalho(valores);

            if (!cabecalho) {
                throw new Error(
                    "Não foi possível localizar os cabeçalhos da aba Config."
                );
            }

            console.log("Linha do cabeçalho:", cabecalho.linha);
            console.log("Cabeçalhos:", cabecalho.colunas);

            const colunas = cabecalho.colunas;

            // ----------------------------------------------
            // EMPRESAS
            // ----------------------------------------------

            const colEmpresa = encontrarColuna(
                colunas,
                [
                    "EMPRESAS",
                    "EMPRESA"
                ]
            );

            // ----------------------------------------------
            // SUPERVISORES
            // ----------------------------------------------

            const colSupervisor = encontrarColuna(
                colunas,
                [
                    "SUPERVISORES",
                    "SUPERVISOR"
                ]
            );

            // ----------------------------------------------
            // REALIZADO POR
            // ----------------------------------------------

            const colRealizadoPor = encontrarColuna(
                colunas,
                [
                    "REALIZADO POR",
                    "REALIZADO_POR",
                    "REALIZADOPOR"
                ]
            );

            // ----------------------------------------------
            // MENSAGEM SMS
            // ----------------------------------------------

            const colNomeSMS = encontrarColuna(
                colunas,
                [
                    "NOME MENSAGEM SMS"
                ]
            );

            const colTextoSMS = encontrarColuna(
                colunas,
                [
                    "TEXTO MENSAGEM SMS"
                ]
            );

            // ----------------------------------------------
            // LIMITE SMS
            // ----------------------------------------------

            const colLimiteSMS = encontrarColuna(
                colunas,
                [
                    "LIMITE SMS"
                ]
            );


            // ----------------------------------------------
            // VALIDAR COLUNAS
            // ----------------------------------------------

            if (colEmpresa === -1) {
                throw new Error("Coluna EMPRESAS não encontrada.");
            }

            if (colSupervisor === -1) {
                throw new Error("Coluna SUPERVISORES não encontrada.");
            }

            if (colRealizadoPor === -1) {
                throw new Error("Coluna REALIZADO POR não encontrada.");
            }

            if (colNomeSMS === -1) {
                throw new Error("Coluna NOME MENSAGEM SMS não encontrada.");
            }

            if (colTextoSMS === -1) {
                throw new Error("Coluna TEXTO MENSAGEM SMS não encontrada.");
            }


            // ----------------------------------------------
            // CARREGAR LISTAS
            // ----------------------------------------------

            configDados.empresas =
                pegarColuna(valores, cabecalho.linha, colEmpresa);

            configDados.supervisores =
                pegarColuna(valores, cabecalho.linha, colSupervisor);

            configDados.realizadoPor =
                pegarColuna(valores, cabecalho.linha, colRealizadoPor);

            const nomesSMS =
                pegarColuna(valores, cabecalho.linha, colNomeSMS);

            const textosSMS =
                pegarColuna(valores, cabecalho.linha, colTextoSMS);


            configDados.mensagens = [];

            configDados.textosMensagens = {};


            for (let i = 0; i < nomesSMS.length; i++) {

                const nome = limparTexto(nomesSMS[i]);

                if (!nome) {
                    continue;
                }

                const texto =
                    textosSMS[i] !== undefined &&
                    textosSMS[i] !== null
                        ? String(textosSMS[i])
                        : "";

                configDados.mensagens.push(nome);

                configDados.textosMensagens[nome] = texto;
            }


            // ----------------------------------------------
            // LIMITE
            // ----------------------------------------------

            configDados.limiteSMS = 160;

            if (colLimiteSMS !== -1) {

                const limites =
                    pegarColuna(
                        valores,
                        cabecalho.linha,
                        colLimiteSMS
                    );

                for (const valor of limites) {

                    if (
                        valor !== null &&
                        valor !== undefined &&
                        String(valor).trim() !== ""
                    ) {

                        const numero = Number(valor);

                        if (!isNaN(numero) && numero > 0) {
                            configDados.limiteSMS = numero;
                            break;
                        }
                    }
                }
            }


            // ----------------------------------------------
            // PREENCHER CAMPOS
            // ----------------------------------------------

            preencherDatalist(
                "listaEmpresas",
                configDados.empresas
            );

            preencherDatalist(
                "listaSupervisores",
                configDados.supervisores
            );

            preencherDatalist(
                "listaRealizadoPor",
                configDados.realizadoPor
            );

            preencherMensagens();


            // ----------------------------------------------
            // GARANTIR SSO NO CAMPO
            // ----------------------------------------------

            const campoRealizado =
                document.getElementById("realizadoPor");

            if (
                campoRealizado &&
                nomeUsuarioSSO
            ) {
                campoRealizado.value =
                    nomeUsuarioSSO;
            }

        });

    } catch (erro) {

        console.error("Erro ao carregar Config:", erro);

        logStatus(
            "❌ Erro ao consultar Config: " +
            escapeHtml(erro.message),
            "erro"
        );

        throw erro;
    }
}


// ======================================================
// LOCALIZAR CABEÇALHO
// ======================================================

function localizarCabecalho(valores) {

    for (let linha = 0; linha < Math.min(valores.length, 30); linha++) {

        const colunas = valores[linha].map(function (valor) {
            return normalizarCabecalho(valor);
        });

        const temEmpresas =
            colunas.includes("EMPRESAS");

        const temSupervisor =
            colunas.includes("SUPERVISORES");

        const temSMS =
            colunas.includes("NOME MENSAGEM SMS");

        if (
            temEmpresas &&
            temSupervisor &&
            temSMS
        ) {
            return {
                linha: linha,
                colunas: colunas
            };
        }
    }

    return null;
}


// ======================================================
// LOCALIZAR COLUNA
// ======================================================

function encontrarColuna(colunas, nomes) {

    const procurados =
        nomes.map(normalizarCabecalho);

    for (let i = 0; i < colunas.length; i++) {

        if (procurados.includes(colunas[i])) {
            return i;
        }
    }

    return -1;
}


// ======================================================
// NORMALIZAÇÃO
// ======================================================

function normalizarCabecalho(valor) {

    if (
        valor === null ||
        valor === undefined
    ) {
        return "";
    }

    return String(valor)
        .trim()
        .toUpperCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ");
}


function limparTexto(valor) {

    if (
        valor === null ||
        valor === undefined
    ) {
        return "";
    }

    return String(valor).trim();
}


// ======================================================
// PEGAR COLUNA
// ======================================================

function pegarColuna(valores, linhaCabecalho, coluna) {

    const resultado = [];

    for (
        let linha = linhaCabecalho + 1;
        linha < valores.length;
        linha++
    ) {

        const valor =
            valores[linha][coluna];

        const texto =
            limparTexto(valor);

        if (texto !== "") {
            resultado.push(texto);
        }
    }

    return resultado;
}


// ======================================================
// DATALIST
// ======================================================

function preencherDatalist(id, valores) {

    const lista =
        document.getElementById(id);

    if (!lista) {
        return;
    }

    lista.innerHTML = "";

    valores.forEach(function (valor) {

        const option =
            document.createElement("option");

        option.value = valor;

        lista.appendChild(option);
    });
}


// ======================================================
// MENSAGENS SMS
// ======================================================

function preencherMensagens() {

    const select =
        document.getElementById("mensagem");

    if (!select) {
        return;
    }

    select.innerHTML = "";

    const primeira =
        document.createElement("option");

    primeira.value = "";
    primeira.textContent =
        configDados.mensagens.length > 0
            ? "Selecione uma mensagem"
            : "Nenhuma mensagem SMS cadastrada";

    select.appendChild(primeira);


    configDados.mensagens.forEach(function (nome) {

        const option =
            document.createElement("option");

        option.value = nome;
        option.textContent = nome;

        select.appendChild(option);
    });
}


// ======================================================
// TEXTO DA MENSAGEM
// ======================================================

function atualizarTextoMensagem() {

    const select =
        document.getElementById("mensagem");

    const texto =
        document.getElementById("textoMensagem");

    if (!select || !texto) {
        return;
    }

    const nome =
        select.value;

    const mensagem =
        configDados.textosMensagens[nome] || "";

    texto.value = mensagem;

    atualizarContador(mensagem);
}


// ======================================================
// CONTADOR
// ======================================================

function atualizarContador(texto) {

    const contador =
        document.getElementById("contadorMensagem");

    if (!contador) {
        return;
    }

    const tamanho =
        texto ? texto.length : 0;

    contador.textContent =
        `${tamanho}/${configDados.limiteSMS} caracteres`;

    if (tamanho > configDados.limiteSMS) {

        contador.style.color = "#a12626";
        contador.style.fontWeight = "bold";

    } else {

        contador.style.color = "#666";
        contador.style.fontWeight = "normal";
    }
}


// ======================================================
// SALVAR SMS
// ======================================================

async function salvarSMS(evento) {

    evento.preventDefault();

    const btnSalvar =
        document.getElementById("btnSalvar");

    try {

        btnSalvar.disabled = true;


        const empresa =
            limparTexto(
                document.getElementById("empresa")?.value
            );

        const qtde =
            limparTexto(
                document.getElementById("qtde")?.value
            );

        const supervisor =
            limparTexto(
                document.getElementById("supervisor")?.value
            );

        const realizadoPor =
            limparTexto(
                document.getElementById("realizadoPor")?.value
            );

        const obs =
            limparTexto(
                document.getElementById("obs")?.value
            );

        const nomeMensagem =
            limparTexto(
                document.getElementById("mensagem")?.value
            );

        const textoMensagem =
            configDados.textosMensagens[nomeMensagem] || "";


        // ----------------------------------------------
        // VALIDAÇÕES
        // ----------------------------------------------

        if (!empresa) {
            throw new Error("Informe a empresa.");
        }

        if (!qtde || Number(qtde) <= 0) {
            throw new Error("Informe uma quantidade válida.");
        }

        if (!supervisor) {
            throw new Error("Informe o supervisor.");
        }

        if (!realizadoPor) {
            throw new Error("Informe quem realizou o registro.");
        }

        if (!nomeMensagem) {
            throw new Error("Selecione uma mensagem SMS.");
        }

        if (!textoMensagem) {
            throw new Error(
                "A mensagem selecionada não possui texto no Config."
            );
        }

        if (
            textoMensagem.length >
            configDados.limiteSMS
        ) {

            throw new Error(
                `A mensagem possui ${textoMensagem.length} caracteres. ` +
                `O limite é ${configDados.limiteSMS}.`
            );
        }


        // ----------------------------------------------
        // DATA E HORA
        // ----------------------------------------------

        const data =
            formatarDataHora(new Date());


        // ----------------------------------------------
        // GRAVAR
        // ----------------------------------------------

        await Excel.run(async function (context) {

            const planilha =
                context.workbook.worksheets.getItem("SMS");

            const usado =
                planilha.getUsedRangeOrNullObject(true);

            usado.load([
                "isNullObject",
                "rowIndex",
                "rowCount",
                "columnCount"
            ]);

            await context.sync();


            let proximaLinha;

            if (usado.isNullObject) {

                proximaLinha = 0;

            } else {

                proximaLinha =
                    usado.rowIndex +
                    usado.rowCount;
            }


            const valores = [[
                data,
                realizadoPor,
                empresa,
                Number(qtde),
                supervisor,
                obs,
                textoMensagem
            ]];

            const destino =
                planilha.getRangeByIndexes(
                    proximaLinha,
                    0,
                    1,
                    7
                );

            destino.values = valores;

            await context.sync();
        });


        // ----------------------------------------------
        // ATUALIZAR CONFIG
        // ----------------------------------------------

        await adicionarConfigSeNovo(
            "EMPRESAS",
            empresa
        );

        await adicionarConfigSeNovo(
            "SUPERVISORES",
            supervisor
        );

        await adicionarConfigSeNovo(
            "REALIZADO POR",
            realizadoPor
        );


        logStatus(
            "✓ Registro SMS salvo com sucesso.",
            "sucesso"
        );


        // ----------------------------------------------
        // LIMPAR FORMULÁRIO
        // ----------------------------------------------

        limparFormularioSMS();

    } catch (erro) {

        console.error(
            "Erro ao salvar SMS:",
            erro
        );

        logStatus(
            "❌ " + escapeHtml(erro.message),
            "erro"
        );

    } finally {

        btnSalvar.disabled = false;
    }
}


// ======================================================
// ADICIONAR NOVO ITEM NO CONFIG
// ======================================================

async function adicionarConfigSeNovo(
    nomeCabecalho,
    valorNovo
) {

    if (!valorNovo) {
        return;
    }

    try {

        await Excel.run(async function (context) {

            const planilha =
                context.workbook.worksheets.getItem("Config");

            const usado =
                planilha.getUsedRangeOrNullObject(true);

            usado.load([
                "isNullObject",
                "values",
                "rowIndex",
                "rowCount",
                "columnCount"
            ]);

            await context.sync();

            if (usado.isNullObject) {
                return;
            }

            const valores =
                usado.values;

            const cabecalho =
                localizarCabecalho(valores);

            if (!cabecalho) {
                return;
            }

            const coluna =
                encontrarColuna(
                    cabecalho.colunas,
                    [nomeCabecalho]
                );

            if (coluna === -1) {
                return;
            }

            const existente =
                pegarColuna(
                    valores,
                    cabecalho.linha,
                    coluna
                );

            const comparacao =
                valorNovo
                    .trim()
                    .toUpperCase();

            const jaExiste =
                existente.some(function (valor) {

                    return String(valor)
                        .trim()
                        .toUpperCase() === comparacao;
                });

            if (jaExiste) {
                return;
            }

            const linhaNova =
                usado.rowIndex +
                usado.rowCount;

            const destino =
                planilha.getRangeByIndexes(
                    linhaNova,
                    coluna,
                    1,
                    1
                );

            destino.values = [[valorNovo]];

            await context.sync();

            console.log(
                `Novo valor adicionado em ${nomeCabecalho}:`,
                valorNovo
            );
        });

    } catch (erro) {

        console.warn(
            "Não foi possível atualizar Config:",
            erro
        );
    }
}


// ======================================================
// LIMPAR FORMULÁRIO
// ======================================================

function limparFormularioSMS() {

    const form =
        document.getElementById("formSMS");

    if (form) {
        form.reset();
    }

    const realizadoPor =
        document.getElementById("realizadoPor");

    if (
        realizadoPor &&
        nomeUsuarioSSO
    ) {
        realizadoPor.value =
            nomeUsuarioSSO;
    }

    const texto =
        document.getElementById("textoMensagem");

    if (texto) {
        texto.value = "";
    }

    atualizarContador("");
}


// ======================================================
// DATA E HORA
// ======================================================

function formatarDataHora(data) {

    const dia =
        String(data.getDate()).padStart(2, "0");

    const mes =
        String(data.getMonth() + 1).padStart(2, "0");

    const ano =
        data.getFullYear();

    const hora =
        String(data.getHours()).padStart(2, "0");

    const minuto =
        String(data.getMinutes()).padStart(2, "0");

    const segundo =
        String(data.getSeconds()).padStart(2, "0");

    return `${dia}/${mes}/${ano} ${hora}:${minuto}:${segundo}`;
}


// ======================================================
// HTML
// ======================================================

function escapeHtml(valor) {

    return String(valor)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
