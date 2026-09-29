/* ============================================================
   REGISTROS DE ENVIOS - SMS
   ============================================================ */

let configInfo = null;
let mensagensSMS = [];
let limiteSMS = 160;

Office.onReady(async (info) => {

    if (info.host !== Office.HostType.Excel) {

        mostrarStatus(
            "Este formulário precisa ser aberto dentro do Excel.",
            "erro"
        );

        return;
    }

    document
        .getElementById("btnSalvar")
        .addEventListener("click", salvar);

    document
        .getElementById("btnCancelar")
        .addEventListener("click", cancelar);

    document
        .getElementById("mensagem")
        .addEventListener(
            "change",
            carregarTextoMensagem
        );

    await inicializar();
});


/* ============================================================
   INICIALIZAÇÃO
   ============================================================ */

async function inicializar() {

    try {

        mostrarStatus(
            "Carregando dados...",
            "info"
        );

        await prepararConfig();

        await carregarListas();

        const usuario =
            await obterUsuarioLogado();

        if (usuario) {

            document.getElementById(
                "realizadoPor"
            ).value = usuario;

            adicionarOpcao(
                "listaRealizadoPor",
                usuario
            );

            mostrarStatus(
                "Usuário identificado automaticamente: " +
                usuario,
                "sucesso"
            );

        } else {

            document.getElementById(
                "realizadoPor"
            ).placeholder =
                "Digite o nome manualmente";

            mostrarStatus(
                "Não foi possível identificar o usuário automaticamente. Preencha manualmente.",
                "aviso"
            );
        }

    } catch (erro) {

        console.error(erro);

        mostrarStatus(
            "Não foi possível carregar todas as listas. Você poderá preencher os campos manualmente.",
            "aviso"
        );
    }
}


/* ============================================================
   PREPARAR CONFIG
   ============================================================ */

async function prepararConfig() {

    await Excel.run(async (context) => {

        const sheet =
            context.workbook.worksheets
                .getItem("Config");

        const used =
            sheet.getUsedRangeOrNullObject(true);

        used.load([
            "isNullObject",
            "values",
            "rowCount",
            "columnCount"
        ]);

        await context.sync();

        if (used.isNullObject) {

            throw new Error(
                "A aba Config está vazia."
            );
        }

        const valores =
            used.values;

        configInfo =
            identificarConfig(valores);

        /*
         * Se as colunas SMS ainda não existirem,
         * criamos automaticamente.
         */

        let houveAlteracao = false;

        if (configInfo.nomeMensagemSMS === -1) {

            sheet
                .getCell(
                    configInfo.headerRow,
                    6
                )
                .values =
                [["NOME MENSAGEM SMS"]];

            configInfo.nomeMensagemSMS = 6;

            houveAlteracao = true;
        }

        if (configInfo.textoMensagemSMS === -1) {

            sheet
                .getCell(
                    configInfo.headerRow,
                    7
                )
                .values =
                [["TEXTO MENSAGEM SMS"]];

            configInfo.textoMensagemSMS = 7;

            houveAlteracao = true;
        }

        if (configInfo.limiteSMS === -1) {

            sheet
                .getCell(
                    configInfo.headerRow,
                    8
                )
                .values =
                [["LIMITE SMS"]];

            configInfo.limiteSMS = 8;

            sheet
                .getCell(
                    configInfo.headerRow + 1,
                    8
                )
                .values =
                [[160]];

            houveAlteracao = true;
        }

        if (houveAlteracao) {
            await context.sync();
        }

        /*
         * Agora carregamos novamente para pegar
         * o valor atualizado do limite.
         */

        const used2 =
            sheet.getUsedRange();

        used2.load("values");

        await context.sync();

        const valoresAtualizados =
            used2.values;

        const linhaLimite =
            configInfo.headerRow + 1;

        if (
            valoresAtualizados[linhaLimite] &&
            valoresAtualizados[linhaLimite][
                configInfo.limiteSMS
            ] !== undefined
        ) {

            const valor =
                Number(
                    valoresAtualizados[
                        linhaLimite
                    ][
                        configInfo.limiteSMS
                    ]
                );

            if (
                Number.isFinite(valor) &&
                valor > 0
            ) {

                limiteSMS =
                    Math.min(valor, 160);
            }
        }

        atualizarContador();
    });
}


/* ============================================================
   IDENTIFICAR CONFIG
   ============================================================ */

function normalizar(texto) {

    return String(texto || "")
        .trim()
        .toUpperCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
}


function identificarConfig(valores) {

    let headerRow = -1;

    for (let i = 0; i < valores.length; i++) {

        const linha =
            valores[i].map(v => normalizar(v));

        if (
            linha.includes("EMPRESAS") &&
            linha.includes("SUPERVISORES")
        ) {

            headerRow = i;
            break;
        }
    }

    if (headerRow === -1) {

        throw new Error(
            "Não encontrei os cabeçalhos da Config."
        );
    }

    const headers =
        valores[headerRow];

    function achar(...nomes) {

        for (const nome of nomes) {

            const indice =
                headers.findIndex(
                    h =>
                        normalizar(h) ===
                        normalizar(nome)
                );

            if (indice !== -1) {
                return indice;
            }
        }

        return -1;
    }

    return {

        headerRow,

        empresas:
            achar("EMPRESAS"),

        supervisores:
            achar("SUPERVISORES"),

        realizadoPor:
            achar(
                "REALIZADO POR",
                "REALIZADO_POR"
            ),

        nomeMensagemSMS:
            achar(
                "NOME MENSAGEM SMS",
                "MENSAGENS SMS",
                "NOME SMS"
            ),

        textoMensagemSMS:
            achar(
                "TEXTO MENSAGEM SMS",
                "TEXTO COMPLETO SMS",
                "TEXTO SMS"
            ),

        limiteSMS:
            achar(
                "LIMITE SMS",
                "LIMITE_SMS"
            )
    };
}


/* ============================================================
   CARREGAR LISTAS
   ============================================================ */

async function carregarListas() {

    await Excel.run(async (context) => {

        const sheet =
            context.workbook.worksheets
                .getItem("Config");

        const used =
            sheet.getUsedRange();

        used.load("values");

        await context.sync();

        const valores =
            used.values;

        const dados =
            valores.slice(
                configInfo.headerRow + 1
            );

        preencherDatalist(
            "listaEmpresas",
            dados,
            configInfo.empresas
        );

        preencherDatalist(
            "listaSupervisores",
            dados,
            configInfo.supervisores
        );

        preencherDatalist(
            "listaRealizadoPor",
            dados,
            configInfo.realizadoPor
        );

        carregarMensagensSMS(
            dados,
            configInfo.nomeMensagemSMS,
            configInfo.textoMensagemSMS
        );
    });
}


/* ============================================================
   DATALIST
   ============================================================ */

function preencherDatalist(
    id,
    dados,
    coluna
) {

    const lista =
        document.getElementById(id);

    lista.innerHTML = "";

    if (
        coluna === -1 ||
        coluna === null ||
        coluna === undefined
    ) {
        return;
    }

    const valores = [];

    for (const linha of dados) {

        const valor =
            String(
                linha[coluna] || ""
            ).trim();

        if (
            valor &&
            !valores.some(
                v =>
                    normalizar(v) ===
                    normalizar(valor)
            )
        ) {

            valores.push(valor);
        }
    }

    valores.forEach(valor => {

        const option =
            document.createElement("option");

        option.value = valor;

        lista.appendChild(option);
    });
}


function adicionarOpcao(
    id,
    valor
) {

    if (!valor) return;

    const lista =
        document.getElementById(id);

    const existe =
        [...lista.options].some(
            option =>
                normalizar(option.value) ===
                normalizar(valor)
        );

    if (!existe) {

        const option =
            document.createElement("option");

        option.value = valor;

        lista.appendChild(option);
    }
}


/* ============================================================
   MENSAGENS SMS
   ============================================================ */

function carregarMensagensSMS(
    dados,
    colunaNome,
    colunaTexto
) {

    const select =
        document.getElementById(
            "mensagem"
        );

    select.innerHTML =
        '<option value="">Selecione uma mensagem SMS</option>';

    mensagensSMS = [];

    if (
        colunaNome === -1 ||
        colunaTexto === -1
    ) {

        mostrarStatus(
            "As colunas de mensagens SMS não foram encontradas na Config.",
            "aviso"
        );

        return;
    }

    for (const linha of dados) {

        const nome =
            String(
                linha[colunaNome] || ""
            ).trim();

        const texto =
            String(
                linha[colunaTexto] || ""
            ).trim();

        if (!nome) continue;

        mensagensSMS.push({
            nome,
            texto
        });

        const option =
            document.createElement("option");

        option.value = nome;
        option.textContent = nome;

        select.appendChild(option);
    }

    if (mensagensSMS.length === 0) {

        mostrarStatus(
            "Nenhuma mensagem SMS foi cadastrada na Config.",
            "aviso"
        );
    }
}


function carregarTextoMensagem() {

    const nome =
        document.getElementById(
            "mensagem"
        ).value;

    const item =
        mensagensSMS.find(
            mensagem =>
                mensagem.nome === nome
        );

    const texto =
        item ? item.texto : "";

    document.getElementById(
        "textoMensagem"
    ).value = texto;

    atualizarContador();
}


/* ============================================================
   CONTADOR SMS
   ============================================================ */

function atualizarContador() {

    const campo =
        document.getElementById(
            "textoMensagem"
        );

    const contador =
        document.getElementById(
            "contador"
        );

    if (!campo || !contador) {
        return;
    }

    const quantidade =
        Array.from(
            campo.value
        ).length;

    contador.textContent =
        `${quantidade} / ${limiteSMS}`;

    if (quantidade > limiteSMS) {

        contador.classList.add(
            "excedido"
        );

    } else {

        contador.classList.remove(
            "excedido"
        );
    }
}


/* ============================================================
   USUÁRIO
   ============================================================ */

async function obterUsuarioLogado() {

    try {

        const token =
            await Office.auth.getAccessToken({
                allowSignInPrompt: true,
                allowConsentPrompt: true
            });

        if (!token) {
            return null;
        }

        const dados =
            decodificarJWT(token);

        if (!dados) {
            return null;
        }

        return (
            dados.name ||
            dados.preferred_username ||
            dados.upn ||
            dados.email ||
            null
        );

    } catch (erro) {

        console.warn(
            "Falha ao identificar usuário:",
            erro
        );

        return null;
    }
}


function decodificarJWT(token) {

    try {

        const partes =
            token.split(".");

        if (partes.length < 2) {
            return null;
        }

        const base64 =
            partes[1]
                .replace(/-/g, "+")
                .replace(/_/g, "/");

        const json =
            decodeURIComponent(
                atob(base64)
                    .split("")
                    .map(
                        c =>
                            "%" +
                            (
                                "00" +
                                c.charCodeAt(0)
                                    .toString(16)
                            ).slice(-2)
                    )
                    .join("")
            );

        return JSON.parse(json);

    } catch (erro) {

        return null;
    }
}


/* ============================================================
   SALVAR SMS
   ============================================================ */

async function salvar() {

    const realizadoPor =
        document.getElementById(
            "realizadoPor"
        ).value.trim();

    const empresa =
        document.getElementById(
            "empresa"
        ).value.trim();

    const qtde =
        document.getElementById(
            "qtde"
        ).value.trim();

    const supervisor =
        document.getElementById(
            "supervisor"
        ).value.trim();

    const obs =
        document.getElementById(
            "obs"
        ).value.trim();

    const mensagemNome =
        document.getElementById(
            "mensagem"
        ).value.trim();

    const mensagemTexto =
        document.getElementById(
            "textoMensagem"
        ).value;


    /* ----------------------------
       VALIDAÇÕES
       ---------------------------- */

    if (!realizadoPor) {

        mostrarStatus(
            "Informe o campo 'Realizado por'.",
            "erro"
        );

        return;
    }

    if (!empresa) {

        mostrarStatus(
            "Informe a empresa.",
            "erro"
        );

        return;
    }

    if (
        !qtde ||
        !Number.isInteger(
            Number(qtde)
        ) ||
        Number(qtde) <= 0
    ) {

        mostrarStatus(
            "Informe uma quantidade válida.",
            "erro"
        );

        return;
    }

    if (!supervisor) {

        mostrarStatus(
            "Informe o supervisor.",
            "erro"
        );

        return;
    }

    if (!mensagemNome) {

        mostrarStatus(
            "Selecione uma mensagem SMS.",
            "erro"
        );

        return;
    }

    if (!mensagemTexto) {

        mostrarStatus(
            "A mensagem selecionada não possui texto.",
            "erro"
        );

        return;
    }

    const quantidadeCaracteres =
        Array.from(
            mensagemTexto
        ).length;

    if (
        quantidadeCaracteres >
        limiteSMS
    ) {

        mostrarStatus(
            `A mensagem possui ${quantidadeCaracteres} caracteres. O limite é ${limiteSMS}.`,
            "erro"
        );

        return;
    }


    try {

        desabilitarBotoes(true);

        mostrarStatus(
            "Salvando registro...",
            "info"
        );

        await atualizarConfig(
            empresa,
            supervisor,
            realizadoPor
        );

        await adicionarRegistroSMS({
            realizadoPor,
            empresa,
            qtde: Number(qtde),
            supervisor,
            obs,
            mensagem: mensagemTexto
        });

        limparFormulario();

        mostrarStatus(
            "SMS salvo com sucesso!",
            "sucesso"
        );

    } catch (erro) {

        console.error(erro);

        mostrarStatus(
            "Erro ao salvar: " +
            (erro.message || erro),
            "erro"
        );

    } finally {

        desabilitarBotoes(false);
    }
}


/* ============================================================
   ATUALIZAR CONFIG
   ============================================================ */

async function atualizarConfig(
    empresa,
    supervisor,
    realizadoPor
) {

    await Excel.run(async (context) => {

        const sheet =
            context.workbook.worksheets
                .getItem("Config");

        const used =
            sheet.getUsedRange();

        used.load("values");

        await context.sync();

        const valores =
            used.values;

        const dados =
            valores.slice(
                configInfo.headerRow + 1
            );

        const ultimaLinha =
            Math.max(
                valores.length,
                configInfo.headerRow + 2
            );

        const itens = [
            {
                coluna: configInfo.empresas,
                valor: empresa
            },
            {
                coluna: configInfo.supervisores,
                valor: supervisor
            },
            {
                coluna: configInfo.realizadoPor,
                valor: realizadoPor
            }
        ];

        for (const item of itens) {

            if (
                item.coluna === -1 ||
                item.coluna === null ||
                item.coluna === undefined
            ) {
                continue;
            }

            const existe =
                dados.some(
                    linha =>
                        normalizar(
                            linha[item.coluna]
                        ) ===
                        normalizar(
                            item.valor
                        )
                );

            if (!existe) {

                const linhaExcel =
                    ultimaLinha + 1;

                sheet
                    .getCell(
                        linhaExcel - 1,
                        item.coluna
                    )
                    .values =
                    [[item.valor]];
            }
        }

        await context.sync();
    });
}


/* ============================================================
   ADICIONAR SMS
   ============================================================ */

async function adicionarRegistroSMS(
    dados
) {

    await Excel.run(async (context) => {

        const sheet =
            context.workbook.worksheets
                .getItem("SMS");

        const used =
            sheet.getUsedRangeOrNullObject(true);

        used.load([
            "isNullObject",
            "values"
        ]);

        await context.sync();

        let valores = [];

        if (!used.isNullObject) {
            valores = used.values;
        }

        let header =
            encontrarCabecalhoSMS(
                valores
            );

        /*
         * Se a aba ainda não estiver estruturada,
         * criamos os cabeçalhos.
         */

        if (!header) {

            const headers = [
                "Data",
                "Realizado por",
                "Empresa",
                "Qtde",
                "Supervisor",
                "Obs",
                "Mensagem"
            ];

            sheet
                .getRange("A1:G1")
                .values = [headers];

            header = {
                row: 0,
                mapa: {
                    "DATA": 0,
                    "REALIZADO POR": 1,
                    "EMPRESA": 2,
                    "QTDE": 3,
                    "SUPERVISOR": 4,
                    "OBS": 5,
                    "MENSAGEM": 6
                },
                columnCount: 7
            };

            valores = [headers];
        }

        const linhaDestino =
            header.row + 1 +
            contarLinhasDados(
                valores,
                header.row
            );

        const linha =
            new Array(
                header.columnCount
            ).fill("");

        preencherColuna(
            linha,
            header,
            "Data",
            formatarData(new Date())
        );

        preencherColuna(
            linha,
            header,
            "Realizado por",
            dados.realizadoPor
        );

        preencherColuna(
            linha,
            header,
            "Empresa",
            dados.empresa
        );

        preencherColuna(
            linha,
            header,
            "Qtde",
            dados.qtde
        );

        preencherColuna(
            linha,
            header,
            "Supervisor",
            dados.supervisor
        );

        preencherColuna(
            linha,
            header,
            "Obs",
            dados.obs
        );

        preencherColuna(
            linha,
            header,
            "Mensagem",
            dados.mensagem
        );

        sheet
            .getRangeByIndexes(
                linhaDestino,
                0,
                1,
                linha.length
            )
            .values = [linha];

        await context.sync();
    });
}


/* ============================================================
   CABEÇALHO SMS
   ============================================================ */

function encontrarCabecalhoSMS(
    valores
) {

    for (
        let r = 0;
        r < valores.length;
        r++
    ) {

        const headers =
            valores[r].map(
                v => normalizar(v)
            );

        if (
            headers.includes("DATA") &&
            headers.includes("EMPRESA") &&
            headers.includes("QTDE") &&
            headers.includes("SUPERVISOR") &&
            headers.includes("MENSAGEM")
        ) {

            const mapa = {};

            headers.forEach(
                (valor, coluna) => {
                    mapa[valor] = coluna;
                }
            );

            return {
                row: r,
                columnCount: headers.length,
                mapa
            };
        }
    }

    return null;
}


/* ============================================================
   UTILIDADES
   ============================================================ */

function preencherColuna(
    linha,
    header,
    nome,
    valor
) {

    const coluna =
        header.mapa[
            normalizar(nome)
        ];

    if (
        coluna !== undefined &&
        coluna >= 0
    ) {

        linha[coluna] = valor;
    }
}


function contarLinhasDados(
    valores,
    headerRow
) {

    let quantidade = 0;

    for (
        let i = headerRow + 1;
        i < valores.length;
        i++
    ) {

        const linha =
            valores[i];

        const possuiDados =
            linha.some(
                valor =>
                    String(
                        valor || ""
                    ).trim() !== ""
            );

        if (possuiDados) {
            quantidade++;
        }
    }

    return quantidade;
}


function formatarData(data) {

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

    return `${dia}/${mes}/${ano}`;
}


function limparFormulario() {

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

    document.getElementById(
        "textoMensagem"
    ).value = "";

    atualizarContador();
}


function cancelar() {

    limparFormulario();

    mostrarStatus(
        "Formulário limpo.",
        "info"
    );
}


function desabilitarBotoes(
    valor
) {

    document.getElementById(
        "btnSalvar"
    ).disabled = valor;

    document.getElementById(
        "btnCancelar"
    ).disabled = valor;
}


function mostrarStatus(
    mensagem,
    tipo
) {

    const status =
        document.getElementById(
            "status"
        );

    status.textContent =
        mensagem;

    status.className =
        "status " +
        (tipo || "");
}
