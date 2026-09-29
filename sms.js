// ======================================================
// SMS - NOVO REGISTRO
// ======================================================

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
    const campoMensagem = document.getElementById("mensagem");

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

    if (campoMensagem) {
        campoMensagem.addEventListener(
            "change",
            atualizarTextoMensagem
        );
    }

    if (formSMS) {
        formSMS.addEventListener(
            "submit",
            salvarSMS
        );
    }

    prepararLog();

    await iniciarSMS();
});


// ======================================================
// INICIAR SMS
// ======================================================

async function iniciarSMS() {

    logStatus(
        "Iniciando formulário...",
        "info"
    );

    try {

        await carregarSSO();

        await carregarConfig();

        atualizarLog();

    } catch (erro) {

        console.error(
            "Erro ao iniciar SMS:",
            erro
        );

        logStatus(
            "❌ Erro ao carregar: " +
            escaparHTML(erro.message),
            "erro"
        );
    }
}


// ======================================================
// LOG
// ======================================================

function prepararLog() {

    const status =
        document.getElementById("mensagemStatus");

    if (!status) {
        return;
    }

    status.style.display = "block";
    status.style.fontSize = "11px";
    status.style.lineHeight = "1.5";
    status.style.padding = "7px 9px";
    status.style.marginBottom = "12px";
    status.style.borderRadius = "6px";
    status.style.background = "#f3f4f6";
    status.style.color = "#555";

    status.innerHTML = "Iniciando...";
}


function logStatus(texto, tipo) {

    const status =
        document.getElementById("mensagemStatus");

    if (!status) {
        console.log("[SMS]", texto);
        return;
    }

    status.style.display = "block";
    status.style.fontSize = "11px";
    status.style.lineHeight = "1.5";
    status.style.padding = "7px 9px";
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


function atualizarLog() {

    const sso = nomeUsuarioSSO
        ? `✓ SSO: ${escaparHTML(nomeUsuarioSSO)}`
        : "⚠ SSO não identificado";

    const empresas =
        `✓ Empresas: ${configDados.empresas.length}`;

    const supervisores =
        `✓ Supervisores: ${configDados.supervisores.length}`;

    const realizado =
        `✓ Realizado por: ${configDados.realizadoPor.length}`;

    const mensagens =
        `✓ Mensagens SMS: ${configDados.mensagens.length}`;

    const limite =
        `✓ Limite: ${configDados.limiteSMS}`;

    logStatus(
        `${sso}<br>` +
        `${empresas} &nbsp; ${supervisores}<br>` +
        `${realizado} &nbsp; ${mensagens} &nbsp; ${limite}`,
        "sucesso"
    );
}


// ======================================================
// SSO
// ======================================================

async function carregarSSO() {

    try {

        const token =
            await OfficeRuntime.auth.getAccessToken({
                allowSignInPrompt: true,
                allowConsentPrompt: true
            });

        if (!token) {
            throw new Error(
                "Token não retornado."
            );
        }

        const partes =
            token.split(".");

        if (partes.length < 2) {
            throw new Error(
                "Token inválido."
            );
        }

        const payload =
            JSON.parse(
                decodificarBase64Url(partes[1])
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
            "✓ SSO:",
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


function decodificarBase64Url(valor) {

    valor =
        valor
            .replace(/-/g, "+")
            .replace(/_/g, "/");

    while (valor.length % 4) {
        valor += "=";
    }

    const texto =
        atob(valor);

    const bytes =
        new Uint8Array(texto.length);

    for (
        let i = 0;
        i < texto.length;
        i++
    ) {
        bytes[i] =
            texto.charCodeAt(i);
    }

    return new TextDecoder(
        "utf-8"
    ).decode(bytes);
}


// ======================================================
// CARREGAR CONFIG
// ======================================================

async function carregarConfig() {

    await Excel.run(async function (context) {

        const config =
            context.workbook.worksheets.getItem(
                "Config"
            );

        // Cabeçalho na linha 2
        // Dados começam na linha 3
        const intervalo =
            config.getRange("A2:I10000");

        intervalo.load("values");

        await context.sync();

        const valores =
            intervalo.values;

        const cabecalhos =
            valores[0];

        console.log(
            "Cabeçalhos Config:",
            cabecalhos
        );


        const colEmpresa =
            localizarColuna(
                cabecalhos,
                "EMPRESAS"
            );

        const colSupervisor =
            localizarColuna(
                cabecalhos,
                "SUPERVISORES"
            );

        const colRealizado =
            localizarColuna(
                cabecalhos,
                "REALIZADO POR"
            );

        const colNomeSMS =
            localizarColuna(
                cabecalhos,
                "NOME MENSAGEM SMS"
            );

        const colTextoSMS =
            localizarColuna(
                cabecalhos,
                "TEXTO MENSAGEM SMS"
            );

        const colLimite =
            localizarColuna(
                cabecalhos,
                "LIMITE SMS"
            );


        if (colEmpresa === -1) {
            throw new Error(
                "EMPRESAS não encontrada no Config."
            );
        }

        if (colSupervisor === -1) {
            throw new Error(
                "SUPERVISORES não encontrada no Config."
            );
        }

        if (colRealizado === -1) {
            throw new Error(
                "REALIZADO POR não encontrada no Config."
            );
        }

        if (colNomeSMS === -1) {
            throw new Error(
                "NOME MENSAGEM SMS não encontrada no Config."
            );
        }

        if (colTextoSMS === -1) {
            throw new Error(
                "TEXTO MENSAGEM SMS não encontrada no Config."
            );
        }


        // ----------------------------------------------
        // EMPRESAS
        // ----------------------------------------------

        configDados.empresas =
            pegarValoresColuna(
                valores,
                colEmpresa
            );


        // ----------------------------------------------
        // SUPERVISORES
        // ----------------------------------------------

        configDados.supervisores =
            pegarValoresColuna(
                valores,
                colSupervisor
            );


        // ----------------------------------------------
        // REALIZADO POR
        // ----------------------------------------------

        configDados.realizadoPor =
            pegarValoresColuna(
                valores,
                colRealizado
            );


        // ----------------------------------------------
        // MENSAGENS SMS
        // ----------------------------------------------

        const nomes =
            pegarValoresColuna(
                valores,
                colNomeSMS
            );

        const textos =
            pegarValoresColuna(
                valores,
                colTextoSMS
            );

        configDados.mensagens = [];
        configDados.textosMensagens = {};


        /*
         * Como as colunas são lidas separadamente,
         * precisamos manter a posição original.
         */

        for (
            let linha = 1;
            linha < valores.length;
            linha++
        ) {

            const nome =
                limpar(
                    valores[linha][colNomeSMS]
                );

            const texto =
                limpar(
                    valores[linha][colTextoSMS]
                );

            if (!nome) {
                continue;
            }

            configDados.mensagens.push(nome);

            configDados.textosMensagens[nome] =
                texto;
        }


        // ----------------------------------------------
        // LIMITE SMS
        // ----------------------------------------------

        configDados.limiteSMS = 160;

        if (colLimite !== -1) {

            for (
                let linha = 1;
                linha < valores.length;
                linha++
            ) {

                const valor =
                    valores[linha][colLimite];

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


        // ----------------------------------------------
        // PREENCHER FORMULÁRIO
        // ----------------------------------------------

        preencherLista(
            "listaEmpresas",
            configDados.empresas
        );

        preencherLista(
            "listaSupervisores",
            configDados.supervisores
        );

        preencherLista(
            "listaRealizadoPor",
            configDados.realizadoPor
        );

        preencherMensagens();


        // Mantém o SSO
        const realizado =
            document.getElementById(
                "realizadoPor"
            );

        if (
            realizado &&
            nomeUsuarioSSO
        ) {
            realizado.value =
                nomeUsuarioSSO;
        }


        console.log(
            "✓ Config carregado"
        );
    });
}


// ======================================================
// LOCALIZAR COLUNA
// ======================================================

function localizarColuna(
    cabecalhos,
    nome
) {

    const procurado =
        normalizarCabecalho(nome);

    for (
        let i = 0;
        i < cabecalhos.length;
        i++
    ) {

        if (
            normalizarCabecalho(
                cabecalhos[i]
            ) === procurado
        ) {
            return i;
        }
    }

    return -1;
}


// ======================================================
// PEGAR VALORES DE UMA COLUNA
// ======================================================

function pegarValoresColuna(
    valores,
    coluna
) {

    const resultado = [];

    for (
        let linha = 1;
        linha < valores.length;
        linha++
    ) {

        const valor =
            limpar(
                valores[linha][coluna]
            );

        if (valor) {
            resultado.push(valor);
        }
    }

    return resultado;
}


// ======================================================
// LISTAS
// ======================================================

function preencherLista(
    id,
    valores
) {

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
        document.getElementById(
            "mensagem"
        );

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


    configDados.mensagens.forEach(
        function (nome) {

            const option =
                document.createElement("option");

            option.value = nome;
            option.textContent = nome;

            select.appendChild(option);
        }
    );
}


// ======================================================
// TEXTO DA MENSAGEM
// ======================================================

function atualizarTextoMensagem() {

    const select =
        document.getElementById(
            "mensagem"
        );

    const texto =
        document.getElementById(
            "textoMensagem"
        );

    if (!select || !texto) {
        return;
    }

    const nome =
        select.value;

    const conteudo =
        configDados.textosMensagens[nome] || "";

    texto.value = conteudo;

    atualizarContador(
        conteudo
    );
}


// ======================================================
// CONTADOR
// ======================================================

function atualizarContador(texto) {

    const contador =
        document.getElementById(
            "contadorMensagem"
        );

    if (!contador) {
        return;
    }

    const quantidade =
        texto
            ? texto.length
            : 0;

    contador.textContent =
        `${quantidade}/${configDados.limiteSMS} caracteres`;

    if (
        quantidade >
        configDados.limiteSMS
    ) {

        contador.style.color =
            "#a12626";

        contador.style.fontWeight =
            "bold";

    } else {

        contador.style.color =
            "#666";

        contador.style.fontWeight =
            "normal";
    }
}


// ======================================================
// SALVAR SMS
// ======================================================

async function salvarSMS(evento) {

    evento.preventDefault();

    const btnSalvar =
        document.getElementById(
            "btnSalvar"
        );

    try {

        btnSalvar.disabled = true;


        // ----------------------------------------------
        // 1. PEGAR CAMPOS
        // ----------------------------------------------

        logStatus(
            "1/4 Verificando dados...",
            "info"
        );


        const empresa =
            limpar(
                document.getElementById(
                    "empresa"
                ).value
            );

        const qtde =
            document.getElementById(
                "qtde"
            ).value;

        const supervisor =
            limpar(
                document.getElementById(
                    "supervisor"
                ).value
            );

        const realizadoPor =
            limpar(
                document.getElementById(
                    "realizadoPor"
                ).value
            );

        const obs =
            limpar(
                document.getElementById(
                    "obs"
                ).value
            );

        const nomeMensagem =
            limpar(
                document.getElementById(
                    "mensagem"
                ).value
            );


        const textoMensagem =
            configDados
                .textosMensagens[
                    nomeMensagem
                ] || "";


        console.log(
            "Dados SMS:",
            {
                empresa,
                qtde,
                supervisor,
                realizadoPor,
                obs,
                nomeMensagem,
                textoMensagem
            }
        );


        // ----------------------------------------------
        // 2. VALIDAÇÕES
        // ----------------------------------------------

        if (!empresa) {
            throw new Error(
                "Empresa não preenchida."
            );
        }

        if (
            !qtde ||
            Number(qtde) <= 0
        ) {
            throw new Error(
                "Quantidade inválida."
            );
        }

        if (!supervisor) {
            throw new Error(
                "Supervisor não preenchido."
            );
        }

        if (!realizadoPor) {
            throw new Error(
                "Realizado por não preenchido."
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
                `A mensagem possui ${textoMensagem.length} ` +
                `caracteres. O limite é ` +
                `${configDados.limiteSMS}.`
            );
        }


        // ----------------------------------------------
        // 3. SALVAR NA ABA SMS
        // ----------------------------------------------

        logStatus(
            "2/4 Salvando na aba SMS...",
            "info"
        );


        const data =
            formatarDataHora(
                new Date()
            );


        await Excel.run(
            async function (context) {

                const sms =
                    context.workbook.worksheets.getItem(
                        "SMS"
                    );


                /*
                 * Cabeçalho = linha 2
                 * Dados = linha 3 em diante
                 *
                 * Vamos olhar SOMENTE a coluna A.
                 */

                const colunaA =
                    sms.getRange(
                        "A3:A10000"
                    );

                colunaA.load(
                    "values"
                );

                await context.sync();


                let linhaDestino = 3;


                for (
                    let i = 0;
                    i < colunaA.values.length;
                    i++
                ) {

                    const valor =
                        colunaA.values[i][0];

                    if (
                        valor === null ||
                        valor === undefined ||
                        String(valor).trim() === ""
                    ) {

                        linhaDestino =
                            i + 3;

                        break;
                    }
                }


                console.log(
                    "Linha SMS:",
                    linhaDestino
                );


                const destino =
                    sms.getRange(
                        `A${linhaDestino}:G${linhaDestino}`
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
                    "✓ Registro SMS gravado."
                );
            }
        );


        // ----------------------------------------------
        // 4. CONFIG
        // ----------------------------------------------

        logStatus(
            "3/4 SMS gravado. Atualizando Config...",
            "info"
        );


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


        logStatus(
            "4/4 ✓ SMS salvo com sucesso!",
            "sucesso"
        );


        // ----------------------------------------------
        // LIMPAR FORMULÁRIO
        // ----------------------------------------------

        document.getElementById(
            "empresa"
        ).value = "";

        document.getElementById(
            "qtde"
        ).value = "";

        document.getElementById(
            "supervisor"
        ).value = "";

        document.getElementById(
            "obs"
        ).value = "";

        document.getElementById(
            "mensagem"
        ).value = "";

        const texto =
            document.getElementById(
                "textoMensagem"
            );

        if (texto) {
            texto.value = "";
        }

        atualizarContador("");


        // Mantém o usuário do SSO
        const realizado =
            document.getElementById(
                "realizadoPor"
            );

        if (
            realizado &&
            nomeUsuarioSSO
        ) {
            realizado.value =
                nomeUsuarioSSO;
        }


    } catch (erro) {

        console.error(
            "❌ ERRO AO SALVAR SMS:",
            erro
        );

        logStatus(
            "❌ " +
            escaparHTML(
                erro.message
            ),
            "erro"
        );

    } finally {

        btnSalvar.disabled = false;
    }
}


// ======================================================
// SALVAR NOVO VALOR NO CONFIG
// ======================================================

async function adicionarConfig(
    nomeColuna,
    valorNovo
) {

    if (!valorNovo) {
        return;
    }


    try {

        await Excel.run(
            async function (context) {

                const config =
                    context.workbook.worksheets.getItem(
                        "Config"
                    );


                /*
                 * Cabeçalho na linha 2.
                 *
                 * A = Empresas
                 * B = Email
                 * C = Supervisores
                 * D = Realizado por
                 * ...
                 */

                const intervalo =
                    config.getRange(
                        "A2:I10000"
                    );

                intervalo.load(
                    "values"
                );

                await context.sync();


                const valores =
                    intervalo.values;

                const cabecalhos =
                    valores[0];


                const coluna =
                    localizarColuna(
                        cabecalhos,
                        nomeColuna
                    );


                if (coluna === -1) {

                    throw new Error(
                        `Coluna ${nomeColuna} não encontrada no Config.`
                    );
                }


                // --------------------------------------
                // VERIFICAR DUPLICADO
                // --------------------------------------

                const comparacao =
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
                        comparacao
                    ) {

                        console.log(
                            `${nomeColuna} já existe:`,
                            valorNovo
                        );

                        return;
                    }
                }


                // --------------------------------------
                // PRIMEIRA LINHA VAZIA
                // --------------------------------------

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


                const letra =
                    numeroParaColuna(
                        coluna + 1
                    );


                console.log(
                    `Gravando ${nomeColuna}:`,
                    `${letra}${linhaDestino}`,
                    valorNovo
                );


                const destino =
                    config.getRange(
                        `${letra}${linhaDestino}`
                    );


                destino.values = [[
                    valorNovo
                ]];


                await context.sync();


                console.log(
                    `✓ ${nomeColuna} adicionado ao Config.`
                );
            }
        );


    } catch (erro) {

        console.error(
            `Erro ao atualizar ${nomeColuna}:`,
            erro
        );

        /*
         * O registro SMS já foi salvo.
         * Portanto, não apagamos nem cancelamos o registro
         * se apenas a atualização do Config falhar.
         */

        logStatus(
            `⚠ SMS salvo, mas não foi possível atualizar ${nomeColuna} no Config.`,
            "aviso"
        );
    }
}


// ======================================================
// NÚMERO → LETRA DA COLUNA
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
// DATA + HORA
// ======================================================

function formatarDataHora(data) {

    const dia =
        String(
            data.getDate()
        ).padStart(2, "0");

    const mes =
        String(
            data.getMonth() + 1
        ).padStart(2, "0");

    const ano =
        data.getFullYear();

    const hora =
        String(
            data.getHours()
        ).padStart(2, "0");

    const minuto =
        String(
            data.getMinutes()
        ).padStart(2, "0");

    const segundo =
        String(
            data.getSeconds()
        ).padStart(2, "0");

    return (
        `${dia}/${mes}/${ano} ` +
        `${hora}:${minuto}:${segundo}`
    );
}


// ======================================================
// LIMPAR TEXTO
// ======================================================

function limpar(valor) {

    if (
        valor === null ||
        valor === undefined
    ) {
        return "";
    }

    return String(valor).trim();
}


// ======================================================
// NORMALIZAR CABEÇALHO
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
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .replace(
            /\s+/g,
            " "
        );
}


// ======================================================
// ESCAPAR HTML
// ======================================================

function escaparHTML(valor) {

    return String(valor)
        .replace(
            /&/g,
            "&amp;"
        )
        .replace(
            /</g,
            "&lt;"
        )
        .replace(
            />/g,
            "&gt;"
        )
        .replace(
            /"/g,
            "&quot;"
        )
        .replace(
            /'/g,
            "&#039;"
        );
}
