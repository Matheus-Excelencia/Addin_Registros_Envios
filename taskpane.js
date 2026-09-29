/* ============================================================
   REGISTROS DE ENVIOS
   TASK PANE - EXCEL
   ============================================================ */

const ABA_EMAIL = "Email";
const ABA_CONFIG = "Config";

let configuracao = {
    empresas: [],
    emails: [],
    supervisores: [],
    realizadoPor: [],
    mensagens: []
};

let mensagemSelecionada = null;


/* ============================================================
   INICIALIZAÇÃO
   ============================================================ */

Office.onReady(async (info) => {

    if (info.host !== Office.HostType.Excel) {
        mostrarStatus(
            "Este complemento deve ser utilizado dentro do Excel.",
            "erro"
        );
        return;
    }

    document
        .getElementById("btnSalvar")
        .addEventListener("click", salvarRegistro);

    document
        .getElementById("btnCancelar")
        .addEventListener("click", limparFormulario);

    document
        .getElementById("mensagem")
        .addEventListener("change", selecionarMensagem);

    await iniciar();

});


/* ============================================================
   INICIAR
   ============================================================ */

async function iniciar() {

    try {

        mostrarStatus("Carregando configurações...", "info");

        await carregarConfiguracao();

        await identificarUsuario();

        preencherListas();

        preencherMensagens();

        mostrarStatus("Formulário pronto.", "sucesso");

    } catch (erro) {

        console.error("Erro ao iniciar:", erro);

        mostrarStatus(
            "Erro ao carregar o formulário: " + obterMensagemErro(erro),
            "erro"
        );

    }

}


/* ============================================================
   NORMALIZAÇÃO
   ============================================================ */

function normalizarTexto(valor) {

    if (valor === null || valor === undefined) {
        return "";
    }

    return String(valor)
        .trim()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toUpperCase();

}


function normalizarCabecalho(valor) {

    return normalizarTexto(valor)
        .replace(/_/g, " ")
        .replace(/\s+/g, " ");

}


/* ============================================================
   LOCALIZAR CABEÇALHO
   ============================================================ */

function localizarCabecalho(valores, nomesAceitos) {

    const procurados = nomesAceitos.map(normalizarCabecalho);

    for (let linha = 0; linha < valores.length; linha++) {

        for (let coluna = 0; coluna < valores[linha].length; coluna++) {

            const valor = normalizarCabecalho(valores[linha][coluna]);

            if (procurados.includes(valor)) {

                return {
                    linha,
                    coluna
                };

            }

        }

    }

    return null;

}


/* ============================================================
   CARREGAR CONFIGURAÇÃO
   ============================================================ */

async function carregarConfiguracao() {

    await Excel.run(async (context) => {

        const planilha = context.workbook.worksheets.getItem(ABA_CONFIG);

        const usado = planilha.getUsedRangeOrNullObject(true);

        usado.load([
            "values",
            "rowCount",
            "columnCount",
            "rowIndex"
        ]);

        await context.sync();

        if (usado.isNullObject) {

            throw new Error(
                "A aba Config está vazia."
            );

        }

        const valores = usado.values;

        const cabEmpresa = localizarCabecalho(
            valores,
            ["EMPRESAS"]
        );

        const cabEmail = localizarCabecalho(
            valores,
            [
                "EMAIL RESPOSTA",
                "EMAILS RESPOSTA",
                "EMAILS_RESPOSTA"
            ]
        );

        const cabSupervisor = localizarCabecalho(
            valores,
            ["SUPERVISORES"]
        );

        const cabRealizado = localizarCabecalho(
            valores,
            ["REALIZADO POR"]
        );

        const cabNomeMensagem = localizarCabecalho(
            valores,
            [
                "NOME MENSAGEM",
                "MENSAGENS"
            ]
        );

        const cabTextoMensagem = localizarCabecalho(
            valores,
            [
                "TEXTO MENSAGEM",
                "TEXTO COMPLETO"
            ]
        );

        if (!cabEmpresa) {
            throw new Error("Não encontrei a coluna EMPRESAS na aba Config.");
        }

        if (!cabEmail) {
            throw new Error("Não encontrei a coluna EMAIL RESPOSTA na aba Config.");
        }

        if (!cabSupervisor) {
            throw new Error("Não encontrei a coluna SUPERVISORES na aba Config.");
        }

        if (!cabNomeMensagem) {
            throw new Error("Não encontrei a coluna NOME MENSAGEM na aba Config.");
        }

        if (!cabTextoMensagem) {
            throw new Error("Não encontrei a coluna TEXTO MENSAGEM na aba Config.");
        }


        configuracao.empresas = obterColunaConfig(
            valores,
            cabEmpresa
        );

        configuracao.emails = obterColunaConfig(
            valores,
            cabEmail
        );

        configuracao.supervisores = obterColunaConfig(
            valores,
            cabSupervisor
        );

        configuracao.realizadoPor = cabRealizado
            ? obterColunaConfig(valores, cabRealizado)
            : [];

        configuracao.mensagens = [];


        const inicioDados = Math.max(
            cabNomeMensagem.linha,
            cabTextoMensagem.linha
        ) + 1;


        for (
            let i = inicioDados;
            i < valores.length;
            i++
        ) {

            const nome =
                valores[i][cabNomeMensagem.coluna];

            const texto =
                valores[i][cabTextoMensagem.coluna];

            if (
                nome !== null &&
                nome !== undefined &&
                String(nome).trim() !== ""
            ) {

                configuracao.mensagens.push({

                    nome: String(nome).trim(),

                    texto:
                        texto === null ||
                        texto === undefined
                            ? ""
                            : String(texto)

                });

            }

        }

    });

}


/* ============================================================
   OBTER COLUNA DA CONFIG
   ============================================================ */

function obterColunaConfig(valores, cabecalho) {

    const lista = [];

    for (
        let i = cabecalho.linha + 1;
        i < valores.length;
        i++
    ) {

        const valor =
            valores[i][cabecalho.coluna];

        if (
            valor !== null &&
            valor !== undefined &&
            String(valor).trim() !== ""
        ) {

            const texto = String(valor).trim();

            if (
                !lista.some(
                    item =>
                        normalizarTexto(item) ===
                        normalizarTexto(texto)
                )
            ) {

                lista.push(texto);

            }

        }

    }

    return lista;

}


/* ============================================================
   PREENCHER LISTAS
   ============================================================ */

function preencherListas() {

    preencherDatalist(
        "listaEmpresas",
        configuracao.empresas
    );

    preencherDatalist(
        "listaEmails",
        configuracao.emails
    );

    preencherDatalist(
        "listaSupervisores",
        configuracao.supervisores
    );

    preencherDatalist(
        "listaRealizadoPor",
        configuracao.realizadoPor
    );

}


/* ============================================================
   DATALIST
   ============================================================ */

function preencherDatalist(id, valores) {

    const lista = document.getElementById(id);

    if (!lista) {
        return;
    }

    lista.innerHTML = "";

    valores.forEach(valor => {

        const option =
            document.createElement("option");

        option.value = valor;

        lista.appendChild(option);

    });

}


/* ============================================================
   MENSAGENS
   ============================================================ */

function preencherMensagens() {

    const select =
        document.getElementById("mensagem");

    select.innerHTML =
        '<option value="">Selecione uma mensagem</option>';

    configuracao.mensagens.forEach(
        (item, indice) => {

            const option =
                document.createElement("option");

            option.value = String(indice);

            option.textContent = item.nome;

            select.appendChild(option);

        }
    );

}


function selecionarMensagem() {

    const select =
        document.getElementById("mensagem");

    const indice = select.value;

    const campoTexto =
        document.getElementById("textoMensagem");

    if (
        indice === "" ||
        !configuracao.mensagens[indice]
    ) {

        mensagemSelecionada = null;

        campoTexto.value = "";

        return;

    }

    mensagemSelecionada =
        configuracao.mensagens[indice];

    campoTexto.value =
        mensagemSelecionada.texto;

}


/* ============================================================
   IDENTIFICAR USUÁRIO
   ============================================================ */

async function identificarUsuario() {

    const campo =
        document.getElementById("realizadoPor");

    try {

        const token =
            await obterTokenSSO();

        if (!token) {
            throw new Error("Não foi possível obter o token.");
        }

        const payload =
            decodificarToken(token);

        console.log("Payload SSO:", payload);


        const nome =
            payload.name ||
            payload.preferred_username ||
            payload.unique_name ||
            payload.upn ||
            "";


        if (nome) {

            campo.value = nome;

            campo.readOnly = true;

            adicionarNaListaLocal(
                configuracao.realizadoPor,
                nome
            );

            preencherDatalist(
                "listaRealizadoPor",
                configuracao.realizadoPor
            );

            mostrarStatus(
                "Usuário identificado automaticamente: " + nome,
                "sucesso"
            );

            return;

        }

        throw new Error(
            "O token não possui o nome do usuário."
        );

    } catch (erro) {

        console.warn(
            "Não foi possível identificar automaticamente:",
            erro
        );

        campo.readOnly = false;

        campo.placeholder =
            "Digite seu nome";

        mostrarStatus(
            "Não foi possível identificar automaticamente. Informe seu nome.",
            "aviso"
        );

    }

}


/* ============================================================
   TOKEN SSO
   ============================================================ */

async function obterTokenSSO() {

    if (
        Office &&
        Office.auth &&
        typeof Office.auth.getAccessToken === "function"
    ) {

        return await Office.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true
        });

    }

    throw new Error(
        "A API de autenticação do Office não está disponível."
    );

}


/* ============================================================
   DECODIFICAR JWT
   ============================================================ */

function decodificarToken(token) {

    const partes =
        token.split(".");

    if (partes.length !== 3) {

        throw new Error(
            "Token SSO inválido."
        );

    }

    let base64 =
        partes[1]
            .replace(/-/g, "+")
            .replace(/_/g, "/");

    while (base64.length % 4) {
        base64 += "=";
    }

    const json =
        decodeURIComponent(
            atob(base64)
                .split("")
                .map(
                    c =>
                        "%" +
                        (
                            "00" +
                            c.charCodeAt(0).toString(16)
                        ).slice(-2)
                )
                .join("")
        );

    return JSON.parse(json);

}


/* ============================================================
   SALVAR REGISTRO
   ============================================================ */

async function salvarRegistro() {

    const botao =
        document.getElementById("btnSalvar");

    try {

        botao.disabled = true;

        mostrarStatus(
            "Validando informações...",
            "info"
        );


        const dados =
            coletarFormulario();


        const erroValidacao =
            validarFormulario(dados);


        if (erroValidacao) {

            mostrarStatus(
                erroValidacao,
                "erro"
            );

            return;

        }


        mostrarStatus(
            "Atualizando listas de configuração...",
            "info"
        );


        await atualizarConfig(dados);


        mostrarStatus(
            "Salvando registro no Excel...",
            "info"
        );


        await adicionarRegistroEmail(dados);


        limparFormulario();


        mostrarStatus(
            "Registro salvo com sucesso!",
            "sucesso"
        );


    } catch (erro) {

        console.error(
            "Erro ao salvar:",
            erro
        );

        mostrarStatus(
            "Erro ao salvar: " +
            obterMensagemErro(erro),
            "erro"
        );

    } finally {

        botao.disabled = false;

    }

}


/* ============================================================
   COLETAR FORMULÁRIO
   ============================================================ */

function coletarFormulario() {

    const mensagem =
        mensagemSelecionada
            ? mensagemSelecionada.texto
            : "";


    return {

        data:
            new Date(),

        realizadoPor:
            document
                .getElementById("realizadoPor")
                .value
                .trim(),

        empresa:
            document
                .getElementById("empresa")
                .value
                .trim(),

        qtde:
            Number(
                document
                    .getElementById("qtde")
                    .value
            ),

        emailResposta:
            document
                .getElementById("emailResposta")
                .value
                .trim(),

        supervisor:
            document
                .getElementById("supervisor")
                .value
                .trim(),

        obs:
            document
                .getElementById("obs")
                .value
                .trim(),

        historicoExterno:
            document
                .getElementById("historicoExterno")
                .value
                .trim(),

        mensagem: mensagem,

        nomeMensagem:
            mensagemSelecionada
                ? mensagemSelecionada.nome
                : "",

        assunto:
            document
                .getElementById("assunto")
                .value
                .trim()

    };

}


/* ============================================================
   VALIDAÇÃO
   ============================================================ */

function validarFormulario(dados) {

    if (!dados.realizadoPor) {

        return "Informe quem realizou o envio.";

    }

    if (!dados.empresa) {

        return "Informe a empresa.";

    }

    if (
        !Number.isInteger(dados.qtde) ||
        dados.qtde <= 0
    ) {

        return "A Qtde deve ser um número inteiro maior que zero.";

    }

    if (!dados.emailResposta) {

        return "Informe o Email Resposta.";

    }


    const emailValido =
        /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    if (
        !emailValido.test(
            dados.emailResposta
        )
    ) {

        return "Informe um Email Resposta válido.";

    }


    if (!dados.supervisor) {

        return "Informe o supervisor.";

    }

    if (!dados.historicoExterno) {

        return "Informe o Histórico Externo.";

    }

    if (!dados.nomeMensagem) {

        return "Selecione uma mensagem.";

    }

    if (!dados.mensagem) {

        return "A mensagem selecionada não possui texto.";

    }

    if (!dados.assunto) {

        return "Informe o assunto.";

    }

    return null;

}


/* ============================================================
   ADICIONAR REGISTRO NA ABA EMAIL
   ============================================================ */

async function adicionarRegistroEmail(dados) {

    await Excel.run(async (context) => {

        const planilha =
            context.workbook.worksheets.getItem(
                ABA_EMAIL
            );


        const tabela =
            obterTabelaEmail(planilha);


        const cabecalho =
            tabela
                .getHeaderRowRange();

        cabecalho.load("values");

        await context.sync();


        const headers =
            cabecalho.values[0];


        const mapa =
            criarMapaCabecalhos(headers);


        const linha =
            new Array(headers.length).fill("");


        preencherCelulaPorCabecalho(
            linha,
            mapa,
            ["DATA"],
            formatarData(dados.data)
        );


        preencherCelulaPorCabecalho(
            linha,
            mapa,
            ["REALIZADO POR"],
            dados.realizadoPor
        );


        preencherCelulaPorCabecalho(
            linha,
            mapa,
            ["EMPRESA"],
            dados.empresa
        );


        preencherCelulaPorCabecalho(
            linha,
            mapa,
            ["QTDE"],
            dados.qtde
        );


        preencherCelulaPorCabecalho(
            linha,
            mapa,
            [
                "EMAIL RESPOSTA",
                "EMAILS RESPOSTA"
            ],
            dados.emailResposta
        );


        preencherCelulaPorCabecalho(
            linha,
            mapa,
            ["SUPERVISOR"],
            dados.supervisor
        );


        preencherCelulaPorCabecalho(
            linha,
            mapa,
            ["OBS"],
            dados.obs
        );


        preencherCelulaPorCabecalho(
            linha,
            mapa,
            [
                "HISTORICO EXTERNO",
                "HISTÓRICO EXTERNO"
            ],
            dados.historicoExterno
        );


        preencherCelulaPorCabecalho(
            linha,
            mapa,
            ["MENSAGEM"],
            dados.mensagem
        );


        preencherCelulaPorCabecalho(
            linha,
            mapa,
            ["ASSUNTO"],
            dados.assunto
        );


        tabela.rows.add(
            null,
            [linha]
        );


        await context.sync();

    });

}


/* ============================================================
   OBTER TABELA EMAIL
   ============================================================ */

function obterTabelaEmail(planilha) {

    const tabelas =
        planilha.tables;

    const tabela =
        tabelas.getItemOrNullObject("Tabela1");

    tabela.load("name");

    return tabela;

}


/* ============================================================
   MAPA DE CABEÇALHOS
   ============================================================ */

function criarMapaCabecalhos(headers) {

    const mapa = {};

    headers.forEach(
        (header, indice) => {

            const chave =
                normalizarCabecalho(header);

            mapa[chave] = indice;

        }
    );

    return mapa;

}


/* ============================================================
   PREENCHER CÉLULA
   ============================================================ */

function preencherCelulaPorCabecalho(
    linha,
    mapa,
    nomes,
    valor
) {

    for (const nome of nomes) {

        const chave =
            normalizarCabecalho(nome);

        if (
            mapa[chave] !== undefined
        ) {

            linha[mapa[chave]] =
                valor;

            return true;

        }

    }

    return false;

}


/* ============================================================
   ATUALIZAR CONFIG
   ============================================================ */

async function atualizarConfig(dados) {

    await Excel.run(async (context) => {

        const planilha =
            context.workbook.worksheets.getItem(
                ABA_CONFIG
            );


        const usado =
            planilha.getUsedRangeOrNullObject(true);

        usado.load([
            "values",
            "rowCount",
            "columnCount",
            "rowIndex"
        ]);


        await context.sync();


        if (usado.isNullObject) {

            throw new Error(
                "A aba Config não possui estrutura."
            );

        }


        const valores =
            usado.values;


        const cabEmpresa =
            localizarCabecalho(
                valores,
                ["EMPRESAS"]
            );


        const cabEmail =
            localizarCabecalho(
                valores,
                [
                    "EMAIL RESPOSTA",
                    "EMAILS RESPOSTA",
                    "EMAILS_RESPOSTA"
                ]
            );


        const cabSupervisor =
            localizarCabecalho(
                valores,
                ["SUPERVISORES"]
            );


        const cabRealizado =
            localizarCabecalho(
                valores,
                ["REALIZADO POR"]
            );


        if (
            !cabEmpresa ||
            !cabEmail ||
            !cabSupervisor
        ) {

            throw new Error(
                "Não foi possível localizar as colunas da aba Config."
            );

        }


        let proximaLinha =
            usado.rowIndex +
            usado.rowCount;


        await adicionarValorConfig(
            context,
            planilha,
            valores,
            cabEmpresa,
            dados.empresa,
            proximaLinha
        );


        await adicionarValorConfig(
            context,
            planilha,
            valores,
            cabEmail,
            dados.emailResposta,
            proximaLinha
        );


        await adicionarValorConfig(
            context,
            planilha,
            valores,
            cabSupervisor,
            dados.supervisor,
            proximaLinha
        );


        if (cabRealizado) {

            await adicionarValorConfig(
                context,
                planilha,
                valores,
                cabRealizado,
                dados.realizadoPor,
                proximaLinha
            );

        }


        await context.sync();

    });


    /*
       Atualiza as listas locais também.
       Assim o novo valor já aparece no formulário
       sem precisar fechar e abrir o complemento.
    */

    adicionarNaListaLocal(
        configuracao.empresas,
        dados.empresa
    );

    adicionarNaListaLocal(
        configuracao.emails,
        dados.emailResposta
    );

    adicionarNaListaLocal(
        configuracao.supervisores,
        dados.supervisor
    );

    adicionarNaListaLocal(
        configuracao.realizadoPor,
        dados.realizadoPor
    );


    preencherListas();

}


/* ============================================================
   ADICIONAR VALOR NA CONFIG
   ============================================================ */

async function adicionarValorConfig(
    context,
    planilha,
    valores,
    cabecalho,
    novoValor,
    proximaLinha
) {

    if (!novoValor) {
        return;
    }


    const valorNormalizado =
        normalizarTexto(novoValor);


    let existe = false;


    for (
        let i = cabecalho.linha + 1;
        i < valores.length;
        i++
    ) {

        const atual =
            valores[i][cabecalho.coluna];


        if (
            atual !== null &&
            atual !== undefined &&
            normalizarTexto(atual) ===
            valorNormalizado
        ) {

            existe = true;

            break;

        }

    }


    if (existe) {
        return;
    }


    const celula =
        planilha.getCell(
            proximaLinha,
            cabecalho.coluna
        );


    celula.values = [
        [novoValor]
    ];

}


/* ============================================================
   ADICIONAR LOCALMENTE
   ============================================================ */

function adicionarNaListaLocal(
    lista,
    valor
) {

    if (!valor) {
        return;
    }


    const existe =
        lista.some(
            item =>
                normalizarTexto(item) ===
                normalizarTexto(valor)
        );


    if (!existe) {

        lista.push(valor);

    }

}


/* ============================================================
   LIMPAR FORMULÁRIO
   ============================================================ */

function limparFormulario() {

    document.getElementById("empresa").value = "";

    document.getElementById("qtde").value = "";

    document.getElementById("emailResposta").value = "";

    document.getElementById("supervisor").value = "";

    document.getElementById("obs").value = "";

    document.getElementById("historicoExterno").value = "";

    document.getElementById("mensagem").value = "";

    document.getElementById("textoMensagem").value = "";

    document.getElementById("assunto").value = "";

    mensagemSelecionada = null;

}


/* ============================================================
   DATA
   ============================================================ */

function formatarData(data) {

    const dia =
        String(data.getDate()).padStart(2, "0");

    const mes =
        String(data.getMonth() + 1).padStart(2, "0");

    const ano =
        data.getFullYear();

    return `${dia}/${mes}/${ano}`;

}


/* ============================================================
   STATUS
   ============================================================ */

function mostrarStatus(
    mensagem,
    tipo
) {

    const status =
        document.getElementById("status");


    status.textContent =
        mensagem;


    status.className =
        "status " + tipo;


}


/* ============================================================
   ERRO
   ============================================================ */

function obterMensagemErro(erro) {

    if (!erro) {
        return "Erro desconhecido.";
    }

    if (erro.message) {
        return erro.message;
    }

    return String(erro);

}
