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
        mensagem.addEventListener("change", atualizarTextoMensagem);
    }

    if (formSMS) {
        formSMS.addEventListener("submit", salvarSMS);
    }

    criarLogInicial();

    await iniciarSMS();
});


// ======================================================
// INICIAR
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
            "❌ Erro ao carregar o formulário: " +
            escapeHtml(erro.message),
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

    logStatus(
        `${sso}<br>` +
        `✓ Empresas: ${configDados.empresas.length}` +
        ` &nbsp;&nbsp; ✓ Supervisores: ${configDados.supervisores.length}<br>` +
        `✓ Realizado por: ${configDados.realizadoPor.length}` +
        ` &nbsp;&nbsp; ✓ Mensagens SMS: ${configDados.mensagens.length}` +
        ` &nbsp;&nbsp; ✓ Limite: ${configDados.limiteSMS}`,
        "sucesso"
    );
}


// ======================================================
// SSO
// ======================================================

async function carregarSSO() {

    try {

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
            throw new Error(
                "Nome do usuário não encontrado."
            );
        }

        const campo =
            document.getElementById("realizadoPor");

        if (campo) {
            campo.value = nomeUsuarioSSO;
        }

        console.log(
            "✓ SSO encontrado:",
            nomeUsuarioSSO
        );

    } catch (erro) {

        console.warn(
            "SSO não disponível:",
            erro
        );

        nomeUsuarioSSO = "";

        const campo =
            document.getElementById("realizadoPor");

        if (campo) {
            campo.value = "";
        }

        logStatus(
            "⚠ SSO não identificado. " +
            "Realizado por pode ser preenchido manualmente.",
            "aviso"
        );
    }
}


function decodeBase64Url(str) {

    str = str
        .replace(/-/g, "+")
        .replace(/_/g, "/");

    while (str.length % 4) {
        str += "=";
    }

    const decoded = atob(str);

    const bytes =
        new Uint8Array(decoded.length);

    for (let i = 0; i < decoded.length; i++) {
        bytes[i] = decoded.charCodeAt(i);
    }

    return new TextDecoder("utf-8").decode(bytes);
}


// ======================================================
// CARREGAR CONFIG
// ======================================================

async function carregarConfig() {

    await Excel.run(async function (context) {

        const config =
            context.workbook.worksheets.getItem("Config");

        const usado =
            config.getUsedRange();

        usado.load([
            "values",
            "rowIndex",
            "rowCount",
            "columnCount"
        ]);

        await context.sync();

        const valores = usado.values;

        const cabecalho =
            localizarCabecalho(valores);

        if (!cabecalho) {
            throw new Error(
                "Cabeçalhos do Config não encontrados."
            );
        }

        const colunas =
            cabecalho.colunas;

        console.log(
            "Cabeçalhos Config:",
            colunas
        );


        const colEmpresa =
            encontrarColuna(
                colunas,
                ["EMPRESAS", "EMPRESA"]
            );

        const colSupervisor =
            encontrarColuna(
                colunas,
                ["SUPERVISORES", "SUPERVISOR"]
            );

        const colRealizado =
            encontrarColuna(
                colunas,
                [
                    "REALIZADO POR",
                    "REALIZADO_POR",
                    "REALIZADOPOR"
                ]
            );

        const colNomeSMS =
            encontrarColuna(
                colunas,
                ["NOME MENSAGEM SMS"]
            );

        const colTextoSMS =
            encontrarColuna(
                colunas,
                ["TEXTO MENSAGEM SMS"]
            );

        const colLimite =
            encontrarColuna(
                colunas,
                ["LIMITE SMS"]
            );


        if (colEmpresa === -1) {
            throw new Error(
                "Coluna EMPRESAS não encontrada."
            );
        }

        if (colSupervisor === -1) {
            throw new Error(
                "Coluna SUPERVISORES não encontrada."
            );
        }

        if (colRealizado === -1) {
            throw new Error(
                "Coluna REALIZADO POR não encontrada."
            );
        }

        if (colNomeSMS === -1) {
            throw new Error(
                "Coluna NOME MENSAGEM SMS não encontrada."
            );
        }

        if (colTextoSMS === -1) {
            throw new Error(
                "Coluna TEXTO MENSAGEM SMS não encontrada."
            );
        }


        configDados.empresas =
            pegarColuna(
                valores,
                cabecalho.linha,
                colEmpresa
            );

        configDados.supervisores =
            pegarColuna(
                valores,
                cabecalho.linha,
                colSupervisor
            );

        configDados.realizadoPor =
            pegarColuna(
                valores,
                cabecalho.linha,
                colRealizado
            );


        const nomesSMS =
            pegarColuna(
                valores,
                cabecalho.linha,
                colNomeSMS
            );

        const textosSMS =
            pegarColuna(
                valores,
                cabecalho.linha,
                colTextoSMS
            );


        configDados.mensagens = [];
        configDados.textosMensagens = {};


        for (let i = 0; i < nomesSMS.length; i++) {

            const nome =
                limparTexto(nomesSMS[i]);

            if (!nome) {
                continue;
            }

            const texto =
                textosSMS[i] !== undefined
                    ? String(textosSMS[i])
                    : "";

            configDados.mensagens.push(nome);

            configDados.textosMensagens[nome] =
                texto;
        }


        // Limite padrão
        configDados.limiteSMS = 160;


        if (colLimite !== -1) {

            const limites =
                pegarColuna(
                    valores,
                    cabecalho.linha,
                    colLimite
                );

            for (const valor of limites) {

                if (
                    valor !== null &&
                    valor !== undefined &&
                    String(valor).trim() !== ""
                ) {

                    const numero =
                        Number(valor);

                    if (
                        !isNaN(numero) &&
                        numero > 0
                    ) {
                        configDados.limiteSMS =
                            numero;
                        break;
                    }
                }
            }
        }


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


        // Recoloca SSO no campo
        const campoRealizado =
            document.getElementById("realizadoPor");

        if (
            campoRealizado &&
            nomeUsuarioSSO
        ) {
            campoRealizado.value =
                nomeUsuarioSSO;
        }


        console.log(
            "✓ Config carregado"
        );

    });
}


// ======================================================
// LOCALIZAR CABEÇALHO
// ======================================================

function localizarCabecalho(valores) {

    for (
        let linha = 0;
        linha < Math.min(valores.length, 30);
        linha++
    ) {

        const colunas =
            valores[linha].map(function (valor) {
                return normalizarCabecalho(valor);
            });


        if (
            colunas.includes("EMPRESAS") &&
            colunas.includes("SUPERVISORES") &&
            colunas.includes("REALIZADO POR") &&
            colunas.includes("NOME MENSAGEM SMS")
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
// ENCONTRAR COLUNA
// ======================================================

function encontrarColuna(colunas, nomes) {

    const procurados =
        nomes.map(normalizarCabecalho);

    for (let i = 0; i < colunas.length; i++) {

        if (
            procurados.includes(colunas[i])
        ) {
            return i;
        }
    }

    return -1;
}


// ======================================================
// PEGAR COLUNA
// ======================================================

function pegarColuna(
    valores,
    linhaCabecalho,
    coluna
) {

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
// DATALISTS
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
// MENSAGENS
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
        configDados.mensagens.length
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
// TEXTO MENSAGEM
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

    contador.style.color =
        tamanho > configDados.limiteSMS
            ? "#a12626"
            : "#666";
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
                document.getElementById("empresa").value
            );

        const qtde =
            document.getElementById("qtde").value;

        const supervisor =
            limparTexto(
                document.getElementById("supervisor").value
            );

        const realizadoPor =
            limparTexto(
                document.getElementById("realizadoPor").value
            );

        const obs =
            limparTexto(
                document.getElementById("obs").value
            );

        const nomeMensagem =
            limparTexto(
                document.getElementById("mensagem").value
            );

        const textoMensagem =
            configDados.textosMensagens[
                nomeMensagem
            ] || "";


        // ----------------------------------------------
        // VALIDAÇÕES
        // ----------------------------------------------

        if (!empresa) {
            throw new Error(
                "Informe a empresa."
            );
        }

        if (
            !qtde ||
            Number(qtde) <= 0
        ) {
            throw new Error(
                "Informe uma quantidade válida."
            );
        }

        if (!supervisor) {
            throw new Error(
                "Informe o supervisor."
            );
        }

        if (!realizadoPor) {
            throw new Error(
                "Informe o Realizado por."
            );
        }

        if (!nomeMensagem) {
            throw new Error(
                "Selecione uma mensagem SMS."
            );
        }

        if (!textoMensagem) {
            throw new Error(
                "A mensagem selecionada não possui texto."
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
        // DATA / HORA
        // ----------------------------------------------

        const data =
            formatarDataHora(new Date());


        // ----------------------------------------------
        // SALVAR NA ABA SMS
        //
        // CABEÇALHO ESTÁ NA LINHA 2
        // PRIMEIRO REGISTRO = LINHA 3
        // ----------------------------------------------

        await Excel.run(async function (context) {

            const planilha =
                context.workbook.worksheets.getItem("SMS");


            // Coluna A, a partir da linha 3
            const colunaData =
                planilha.getRange("A3:A1048576");

            colunaData.load([
                "values"
            ]);

            await context.sync();


            let proximaLinha = 3;


            for (
                let i = 0;
                i < colunaData.values.length;
                i++
            ) {

                const valor =
                    colunaData.values[i][0];

                if (
                    valor === null ||
                    valor === undefined ||
                    String(valor).trim() === ""
                ) {

                    proximaLinha =
                        i + 3;

                    break;
                }
            }


            console.log(
                "Gravando SMS na linha:",
                proximaLinha
            );


            const destino =
                planilha.getRange(
                    `A${proximaLinha}:G${proximaLinha}`
                );


            destino.values = [[
                data,
                realizadoPor,
                empresa,
                Number(qtde),
                supervisor,
                obs,
                textoMensagem
            ]];

            await context.sync();


            console.log(
                "✓ SMS gravado com sucesso."
            );
        });


        // ----------------------------------------------
        // ATUALIZAR CONFIG
        // ----------------------------------------------

        await adicionarConfig(
            "EMPRESAS",
            empresa
        );

        await adicionarConfig(
            "SUPERVISORES",
            supervisor
        );

        await adicionarConfig(
            "REALIZADO POR",
            realizadoPor
        );


        // ----------------------------------------------
        // SUCESSO
        // ----------------------------------------------

        logStatus(
            "✓ SMS salvo com sucesso.",
            "sucesso"
        );


        limparFormularioSMS();


    } catch (erro) {

        console.error(
            "Erro ao salvar SMS:",
            erro
        );

        logStatus(
            "❌ " +
            escapeHtml(erro.message),
            "erro"
        );

    } finally {

        btnSalvar.disabled = false;
    }
}


// ======================================================
// ADICIONAR NO CONFIG
// ======================================================

async function adicionarConfig(
    nomeCabecalho,
    valorNovo
) {

    if (!valorNovo) {
        return;
    }


    try {

        await Excel.run(async function (context) {

            const config =
                context.workbook.worksheets.getItem("Config");


            // Cabeçalho está na linha 2
            const intervalo =
                config.getRange("A2:I1048576");

            intervalo.load([
                "values"
            ]);

            await context.sync();


            const valores =
                intervalo.values;


            // ------------------------------------------
            // LOCALIZAR COLUNA
            // ------------------------------------------

            let coluna = -1;

            const cabecalhos =
                valores[0];


            for (
                let i = 0;
                i < cabecalhos.length;
                i++
            ) {

                if (
                    normalizarCabecalho(
                        cabecalhos[i]
                    ) ===
                    normalizarCabecalho(
                        nomeCabecalho
                    )
                ) {

                    coluna = i;
                    break;
                }
            }


            if (coluna === -1) {

                console.warn(
                    `Coluna ${nomeCabecalho} não encontrada.`
                );

                return;
            }


            // ------------------------------------------
            // VERIFICAR SE JÁ EXISTE
            // ------------------------------------------

            const valorComparar =
                valorNovo
                    .trim()
                    .toUpperCase();


            for (
                let i = 1;
                i < valores.length;
                i++
            ) {

                const atual =
                    valores[i][coluna];


                if (
                    atual !== null &&
                    atual !== undefined &&
                    String(atual)
                        .trim()
                        .toUpperCase() ===
                    valorComparar
                ) {

                    console.log(
                        `${nomeCabecalho} já existe:`,
                        valorNovo
                    );

                    return;
                }
            }


            // ------------------------------------------
            // ENCONTRAR PRIMEIRA LINHA VAZIA
            // ------------------------------------------

            let linhaDestino = 3;


            for (
                let i = 1;
                i < valores.length;
                i++
            ) {

                const atual =
                    valores[i][coluna];


                if (
                    atual === null ||
                    atual === undefined ||
                    String(atual).trim() === ""
                ) {

                    linhaDestino =
                        i + 2;

                    break;
                }
            }


            console.log(
                `Gravando ${nomeCabecalho} na linha ${linhaDestino}:`,
                valorNovo
            );


            const destino =
                config.getRange(
                    `${numeroParaColuna(coluna + 1)}${linhaDestino}`
                );


            destino.values = [[
                valorNovo
            ]];


            await context.sync();


            console.log(
                `✓ ${nomeCabecalho} salvo no Config.`
            );
        });


    } catch (erro) {

        console.error(
            `Erro ao salvar ${nomeCabecalho}:`,
            erro
        );

        // Não interrompe o salvamento do SMS
        // caso apenas o Config apresente problema.
    }
}


// ======================================================
// CONVERTER NÚMERO PARA LETRA DA COLUNA
// ======================================================

function numeroParaColuna(numero) {

    let resultado = "";

    while (numero > 0) {

        const resto =
            (numero - 1) % 26;

        resultado =
            String.fromCharCode(
                65 + resto
            ) + resultado;

        numero =
            Math.floor(
                (numero - 1) / 26
            );
    }

    return resultado;
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


    const realizado =
        document.getElementById("realizadoPor");

    if (
        realizado &&
        nomeUsuarioSSO
    ) {
        realizado.value =
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
        String(
            data.getMonth() + 1
        ).padStart(2, "0");

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
// ESCAPAR HTML
// ======================================================

function escapeHtml(valor) {

    return String(valor)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
