/* ============================================================
   REGISTROS DE ENVIOS - E-MAIL
   ============================================================ */

let configInfo = null;
let mensagensEmail = [];

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
        .addEventListener("change", carregarTextoMensagem);

    await inicializar();
});


/* ============================================================
   INICIALIZAÇÃO
   ============================================================ */

async function inicializar() {

    try {

        mostrarStatus("Carregando dados...", "info");

        await Excel.run(async (context) => {

            const sheet = context.workbook.worksheets.getItem("Config");

            const used = sheet.getUsedRangeOrNullObject(true);

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

            configInfo = identificarConfig(
                used.values
            );
        });

        await carregarListas();

        const usuario = await obterUsuarioLogado();

        if (usuario) {

            document.getElementById(
                "realizadoPor"
            ).value = usuario;

            adicionarOpcao(
                "listaRealizadoPor",
                usuario
            );

            mostrarStatus(
                "Usuário identificado automaticamente: " + usuario,
                "sucesso"
            );

        } else {

            document.getElementById(
                "realizadoPor"
            ).placeholder = "Digite o nome manualmente";

            mostrarStatus(
                "Não foi possível identificar o usuário automaticamente. Preencha 'Realizado por' manualmente.",
                "aviso"
            );
        }

    } catch (erro) {

        console.error(erro);

        mostrarStatus(
            "Não foi possível carregar todas as listas. Os campos continuam disponíveis para preenchimento manual.",
            "aviso"
        );
    }
}


/* ============================================================
   CONFIG
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

        const linha = valores[i]
            .map(v => normalizar(v));

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
            "Não encontrei a linha de cabeçalhos da aba Config."
        );
    }

    const headers = valores[headerRow];

    function achar(...nomes) {

        for (const nome of nomes) {

            const indice = headers.findIndex(
                h => normalizar(h) === normalizar(nome)
            );

            if (indice !== -1) {
                return indice;
            }
        }

        return -1;
    }

    return {
        headerRow,

        empresas: achar("EMPRESAS"),

        emails: achar(
            "EMAILS_RESPOSTA",
            "EMAIL RESPOSTA"
        ),

        supervisores: achar("SUPERVISORES"),

        realizadoPor: achar(
            "REALIZADO POR",
            "REALIZADO_POR"
        ),

        nomeMensagemEmail: achar(
            "NOME MENSAGEM EMAIL",
            "NOME MENSAGEM E-MAIL",
            "MENSAGENS",
            "NOME MENSAGEM"
        ),

        textoMensagemEmail: achar(
            "TEXTO MENSAGEM EMAIL",
            "TEXTO MENSAGEM E-MAIL",
            "TEXTO COMPLETO",
            "TEXTO MENSAGEM"
        )
    };
}


/* ============================================================
   CARREGAR LISTAS
   ============================================================ */

async function carregarListas() {

    await Excel.run(async (context) => {

        const sheet =
            context.workbook.worksheets.getItem("Config");

        const used =
            sheet.getUsedRange();

        used.load([
            "values",
            "rowCount",
            "columnCount"
        ]);

        await context.sync();

        const valores = used.values;

        const dados = valores.slice(
            configInfo.headerRow + 1
        );

        preencherDatalist(
            "listaEmpresas",
            dados,
            configInfo.empresas
        );

        preencherDatalist(
            "listaEmails",
            dados,
            configInfo.emails
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

        carregarMensagensEmail(
            dados,
            configInfo.nomeMensagemEmail,
            configInfo.textoMensagemEmail
        );
    });
}


/* ============================================================
   DATALISTS
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
            String(linha[coluna] || "").trim();

        if (
            valor &&
            !valores.some(
                v => normalizar(v) === normalizar(valor)
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


function adicionarOpcao(id, valor) {

    if (!valor) return;

    const lista =
        document.getElementById(id);

    const existe =
        [...lista.options].some(
            o =>
                normalizar(o.value) ===
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
   MENSAGENS E-MAIL
   ============================================================ */

function carregarMensagensEmail(
    dados,
    colunaNome,
    colunaTexto
) {

    const select =
        document.getElementById("mensagem");

    select.innerHTML =
        '<option value="">Selecione uma mensagem</option>';

    mensagensEmail = [];

    if (
        colunaNome === -1 ||
        colunaTexto === -1
    ) {
        return;
    }

    for (const linha of dados) {

        const nome =
            String(linha[colunaNome] || "").trim();

        const texto =
            String(linha[colunaTexto] || "").trim();

        if (!nome) continue;

        mensagensEmail.push({
            nome,
            texto
        });

        const option =
            document.createElement("option");

        option.value = nome;
        option.textContent = nome;

        select.appendChild(option);
    }
}


function carregarTextoMensagem() {

    const nome =
        document.getElementById("mensagem").value;

    const item =
        mensagensEmail.find(
            m => m.nome === nome
        );

    document.getElementById(
        "textoMensagem"
    ).value = item ? item.texto : "";
}


/* ============================================================
   USUÁRIO LOGADO
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

        console.warn(
            "Não foi possível ler o token.",
            erro
        );

        return null;
    }
}


/* ============================================================
   SALVAR
   ============================================================ */

async function salvar() {

    const realizadoPor =
        document.getElementById("realizadoPor")
            .value.trim();

    const empresa =
        document.getElementById("empresa")
            .value.trim();

    const qtde =
        document.getElementById("qtde")
            .value.trim();

    const emailResposta =
        document.getElementById("emailResposta")
            .value.trim();

    const supervisor =
        document.getElementById("supervisor")
            .value.trim();

    const obs =
        document.getElementById("obs")
            .value.trim();

    const historicoExterno =
        document.getElementById("historicoExterno")
            .value.trim();

    const mensagem =
        document.getElementById("mensagem")
            .value.trim();

    const textoMensagem =
        document.getElementById("textoMensagem")
            .value;

    const assunto =
        document.getElementById("assunto")
            .value.trim();


    /* ----------------------------
       VALIDAÇÃO
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
        !Number.isInteger(Number(qtde)) ||
        Number(qtde) <= 0
    ) {

        mostrarStatus(
            "Informe uma quantidade válida.",
            "erro"
        );

        return;
    }

    if (!emailResposta) {

        mostrarStatus(
            "Informe o E-mail Resposta.",
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

    if (!historicoExterno) {

        mostrarStatus(
            "Informe o Histórico Externo.",
            "erro"
        );

        return;
    }

    if (!mensagem) {

        mostrarStatus(
            "Selecione uma mensagem.",
            "erro"
        );

        return;
    }

    if (!assunto) {

        mostrarStatus(
            "Informe o assunto.",
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

        await adicionarRegistroEmail({
            realizadoPor,
            empresa,
            qtde: Number(qtde),
            emailResposta,
            supervisor,
            obs,
            historicoExterno,
            textoMensagem,
            assunto
        });

        limparFormulario();

        mostrarStatus(
            "Registro salvo com sucesso!",
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

        const valores = used.values;

        const headerRow =
            configInfo.headerRow;

        const dados =
            valores.slice(headerRow + 1);

        const ultimaLinha =
            Math.max(
                valores.length,
                headerRow + 2
            );

        const valoresNovos = [
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

        for (const item of valoresNovos) {

            if (
                item.coluna === -1 ||
                item.coluna === null ||
                item.coluna === undefined
            ) {
                continue;
            }

            const jaExiste =
                dados.some(
                    linha =>
                        normalizar(
                            linha[item.coluna]
                        ) === normalizar(item.valor)
                );

            if (!jaExiste) {

                const linhaExcel =
                    ultimaLinha + 1;

                sheet
                    .getCell(
                        linhaExcel - 1,
                        item.coluna
                    )
                    .values = [[item.valor]];
            }
        }

        await context.sync();
    });
}


/* ============================================================
   ADICIONAR E-MAIL
   ============================================================ */

async function adicionarRegistroEmail(dados) {

    await Excel.run(async (context) => {

        const sheet =
            context.workbook.worksheets
                .getItem("Email");

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
                "A aba Email não possui estrutura."
            );
        }

        const valores =
            used.values;

        const header = encontrarCabecalhoEmail(
            valores
        );

        if (!header) {
            throw new Error(
                "Não encontrei os cabeçalhos da aba Email."
            );
        }

        const linhaDestino =
            header.row + 1 +
            contarLinhasDados(
                valores,
                header.row
            );

        const linha =
            new Array(
                Math.max(
                    header.columnCount,
                    10
                )
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
            "Email Resposta",
            dados.emailResposta
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
            "Historico Externo",
            dados.historicoExterno
        );

        preencherColuna(
            linha,
            header,
            "Mensagem",
            dados.textoMensagem
        );

        preencherColuna(
            linha,
            header,
            "Assunto",
            dados.assunto
        );

        const range =
            sheet.getRangeByIndexes(
                linhaDestino,
                0,
                1,
                linha.length
            );

        range.values = [linha];

        await context.sync();
    });
}


/* ============================================================
   CABEÇALHOS E-MAIL
   ============================================================ */

function encontrarCabecalhoEmail(valores) {

    for (let r = 0; r < valores.length; r++) {

        const headers =
            valores[r].map(v => normalizar(v));

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


function preencherColuna(
    linha,
    header,
    nome,
    valor
) {

    const chave =
        normalizar(nome);

    const coluna =
        header.mapa[chave];

    if (
        coluna !== undefined &&
        coluna >= 0
    ) {
        linha[coluna] = valor;
    }
}


/* ============================================================
   UTILIDADES
   ============================================================ */

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
                    String(valor || "").trim() !== ""
            );

        if (possuiDados) {
            quantidade++;
        }
    }

    return quantidade;
}


function formatarData(data) {

    const dia =
        String(data.getDate())
            .padStart(2, "0");

    const mes =
        String(data.getMonth() + 1)
            .padStart(2, "0");

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
        "emailResposta"
    ).value = "";

    document.getElementById(
        "supervisor"
    ).value = "";

    document.getElementById(
        "obs"
    ).value = "";

    document.getElementById(
        "historicoExterno"
    ).value = "";

    document.getElementById(
        "mensagem"
    ).value = "";

    document.getElementById(
        "textoMensagem"
    ).value = "";

    document.getElementById(
        "assunto"
    ).value = "";
}


function cancelar() {

    limparFormulario();

    mostrarStatus(
        "Formulário limpo.",
        "info"
    );
}


function desabilitarBotoes(valor) {

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
        document.getElementById("status");

    status.textContent = mensagem;

    status.className =
        "status " + (tipo || "");
}
