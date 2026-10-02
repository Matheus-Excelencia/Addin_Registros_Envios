let configInfo = null;
let mensagensEmail = [];
let statusLog = [];


/* =========================================================
   INÍCIO
========================================================= */

Office.onReady(async function () {

    document
        .getElementById("btnVoltar")
        .addEventListener("click", function () {
            window.location.href = "taskpane.html";
        });

    document
        .getElementById("btnCancelar")
        .addEventListener("click", limparFormulario);

    document
        .getElementById("formEmail")
        .addEventListener("submit", salvarEmail);

    document
        .getElementById("mensagem")
        .addEventListener("change", mostrarTextoMensagem);

    prepararLog();

    await carregarDados();
    preencherDuplicacao();
});


/* =========================================================
   LOG / STATUS
========================================================= */

function prepararLog() {

    const elemento =
        document.getElementById("mensagemStatus");

    if (!elemento) return;

    elemento.style.whiteSpace = "pre-line";
    elemento.style.fontSize = "12px";
    elemento.style.lineHeight = "1.4";
    elemento.style.padding = "8px 10px";
    elemento.style.marginBottom = "12px";
}


function atualizarLog() {

    const elemento =
        document.getElementById("mensagemStatus");

    if (!elemento) return;

    elemento.style.whiteSpace = "pre-line";
    elemento.style.fontSize = "12px";
    elemento.style.lineHeight = "1.4";

    elemento.textContent =
        statusLog.join("\n");

    elemento.className = "status aviso";
}


function adicionarLog(texto) {

    statusLog.push(texto);

    atualizarLog();
}


function limparStatus() {

    statusLog = [];

    const elemento =
        document.getElementById("mensagemStatus");

    if (!elemento) return;

    elemento.textContent = "";
    elemento.className = "status";
}


function mostrarStatus(texto, tipo) {

    const elemento =
        document.getElementById("mensagemStatus");

    if (!elemento) return;

    elemento.textContent = texto;
    elemento.className = "status " + tipo;
}


/* =========================================================
   LIMPAR FORMULÁRIO
========================================================= */

function limparFormulario() {

    document
        .getElementById("formEmail")
        .reset();

    document
        .getElementById("mensagem")
        .selectedIndex = 0;

    document
        .getElementById("textoMensagem")
        .value = "";

    limparStatus();

    identificarUsuario();
}


/* =========================================================
   NORMALIZAÇÃO
========================================================= */

function normalizar(valor) {

    return String(valor || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .toUpperCase();
}


/* =========================================================
   LOCALIZAR CABEÇALHO DO CONFIG
========================================================= */

function encontrarCabecalho(valores) {

    for (
        let linha = 0;
        linha < Math.min(valores.length, 30);
        linha++
    ) {

        const mapa = {};

        valores[linha].forEach(function (
            valor,
            indice
        ) {

            if (
                valor !== null &&
                valor !== undefined &&
                String(valor).trim() !== ""
            ) {

                mapa[normalizar(valor)] = indice;
            }
        });


        const temEmpresa =
            obterIndice(
                mapa,
                ["EMPRESAS", "EMPRESA"]
            ) !== -1;


        const temSupervisor =
            obterIndice(
                mapa,
                ["SUPERVISORES", "SUPERVISOR"]
            ) !== -1;


        if (
            temEmpresa &&
            temSupervisor
        ) {

            return {
                linha: linha,
                mapa: mapa
            };
        }
    }

    return null;
}


/* =========================================================
   OBTER ÍNDICE
========================================================= */

function obterIndice(mapa, nomes) {

    for (const nome of nomes) {

        const chave =
            normalizar(nome);

        if (
            mapa[chave] !== undefined
        ) {

            return mapa[chave];
        }
    }

    return -1;
}


/* =========================================================
   OBTER VALORES
========================================================= */

function obterValoresColuna(
    valores,
    info,
    nomes
) {

    const indice =
        obterIndice(
            info.mapa,
            Array.isArray(nomes)
                ? nomes
                : [nomes]
        );


    if (indice === -1) {

        return [];
    }


    const resultado = [];


    for (
        let linha = info.linha + 1;
        linha < valores.length;
        linha++
    ) {

        const valor =
            valores[linha][indice];


        if (
            valor !== null &&
            valor !== undefined &&
            String(valor).trim() !== ""
        ) {

            resultado.push(
                String(valor).trim()
            );
        }
    }


    return [
        ...new Set(resultado)
    ];
}


/* =========================================================
   PREENCHER LISTA
========================================================= */

function preencherLista(
    id,
    valores
) {

    const lista =
        document.getElementById(id);

    if (!lista) return;

    lista.innerHTML = "";


    valores.forEach(function (valor) {

        const option =
            document.createElement("option");

        option.value = valor;

        lista.appendChild(option);
    });
}


/* =========================================================
   CARREGAR DADOS DO CONFIG
========================================================= */

async function carregarDados() {

    limparStatus();

    statusLog = [];

    adicionarLog("● Conectando ao Config...");


    try {

        await Excel.run(
            async function (context) {

                const folha =
                    context.workbook.worksheets.getItem(
                        "Config"
                    );


                const usado =
                    folha.getUsedRangeOrNullObject(true);


                usado.load([
                    "values",
                    "rowIndex",
                    "isNullObject"
                ]);


                await context.sync();


                if (usado.isNullObject) {

                    throw new Error(
                        "A aba Config está vazia."
                    );
                }


                configInfo =
                    encontrarCabecalho(
                        usado.values
                    );


                if (!configInfo) {

                    throw new Error(
                        "Cabeçalho do Config não encontrado."
                    );
                }


                /* EMPRESAS */

                const empresas =
                    obterValoresColuna(
                        usado.values,
                        configInfo,
                        [
                            "EMPRESAS",
                            "EMPRESA"
                        ]
                    );


                preencherLista(
                    "listaEmpresas",
                    empresas
                );


                adicionarLog(
                    empresas.length > 0
                        ? "✓ Empresas " + empresas.length
                        : "⚠ Empresas 0"
                );


                /* E-MAILS */

                const emails =
                    obterValoresColuna(
                        usado.values,
                        configInfo,
                        [
                            "EMAILS_RESPOSTA",
                            "E-MAILS RESPOSTA",
                            "EMAIL RESPOSTA",
                            "E-MAIL RESPOSTA"
                        ]
                    );


                preencherLista(
                    "listaEmails",
                    emails
                );


                adicionarLog(
                    emails.length > 0
                        ? "✓ E-mails " + emails.length
                        : "⚠ E-mails 0"
                );


                /* SUPERVISORES */

                const supervisores =
                    obterValoresColuna(
                        usado.values,
                        configInfo,
                        [
                            "SUPERVISORES",
                            "SUPERVISOR"
                        ]
                    );


                preencherLista(
                    "listaSupervisores",
                    supervisores
                );


                adicionarLog(
                    supervisores.length > 0
                        ? "✓ Supervisores " + supervisores.length
                        : "⚠ Supervisores 0"
                );


                /* REALIZADO POR */

                const realizadoPor =
                    obterValoresColuna(
                        usado.values,
                        configInfo,
                        [
                            "REALIZADO POR",
                            "REALIZADO_POR",
                            "REALIZADOPOR"
                        ]
                    );


                preencherLista(
                    "listaRealizadoPor",
                    realizadoPor
                );


                adicionarLog(
                    realizadoPor.length > 0
                        ? "✓ Realizado por " + realizadoPor.length
                        : "⚠ Realizado por 0"
                );


                /* MENSAGENS */

                carregarMensagens(
                    usado.values,
                    configInfo
                );

            }
        );


        await identificarUsuario();


    }
    catch (erro) {

        console.error(erro);

        adicionarLog(
            "⚠ Erro no Config"
        );

        mostrarStatus(
            statusLog.join("\n"),
            "aviso"
        );


        await identificarUsuario();
    }
}


/* =========================================================
   MENSAGENS
========================================================= */

function carregarMensagens(
    valores,
    info
) {

    const indiceNome =
        obterIndice(
            info.mapa,
            [
                "NOME MENSAGEM EMAIL",
                "MENSAGENS",
                "NOME MENSAGEM"
            ]
        );


    const indiceTexto =
        obterIndice(
            info.mapa,
            [
                "TEXTO MENSAGEM EMAIL",
                "TEXTO COMPLETO",
                "TEXTO MENSAGEM"
            ]
        );


    const select =
        document.getElementById("mensagem");


    select.innerHTML = "";

    mensagensEmail = [];


    if (
        indiceNome === -1 ||
        indiceTexto === -1
    ) {

        const option =
            document.createElement("option");

        option.value = "";

        option.textContent =
            "Mensagens de E-mail não encontradas";

        select.appendChild(option);


        adicionarLog(
            "⚠ Mensagens de E-mail 0"
        );

        return;
    }


    const inicial =
        document.createElement("option");

    inicial.value = "";

    inicial.textContent =
        "Selecione uma mensagem";

    select.appendChild(inicial);


    for (
        let linha = info.linha + 1;
        linha < valores.length;
        linha++
    ) {

        const nome =
            valores[linha][indiceNome];

        const texto =
            valores[linha][indiceTexto];


        if (
            nome !== null &&
            nome !== undefined &&
            String(nome).trim() !== ""
        ) {

            const mensagem = {

                nome:
                    String(nome).trim(),

                texto:
                    String(texto || "")

            };


            mensagensEmail.push(
                mensagem
            );


            const option =
                document.createElement("option");

            option.value =
                mensagem.nome;

            option.textContent =
                mensagem.nome;

            select.appendChild(option);
        }
    }


    adicionarLog(
        mensagensEmail.length > 0
            ? "✓ Mensagens " + mensagensEmail.length
            : "⚠ Mensagens 0"
    );
}


/* =========================================================
   TEXTO DA MENSAGEM
========================================================= */

function mostrarTextoMensagem() {

    const nome =
        document.getElementById(
            "mensagem"
        ).value;


    const mensagem =
        mensagensEmail.find(
            function (item) {
                return item.nome === nome;
            }
        );


    document.getElementById(
        "textoMensagem"
    ).value =
        mensagem
            ? mensagem.texto
            : "";
}


/* =========================================================
   SSO
========================================================= */

async function identificarUsuario() {

    const campo =
        document.getElementById(
            "realizadoPor"
        );


    try {

        const token =
            await Office.auth.getAccessToken({
                allowSignInPrompt: true,
                allowConsentPrompt: true
            });


        const dados =
            decodificarToken(token);


        const nome =
            dados.name ||
            dados.preferred_username ||
            dados.email ||
            dados.upn;


        if (nome) {

            campo.value = nome;

            campo.readOnly = false;


            adicionarLog(
                "✓ SSO: " + nome
            );

        }
        else {

            adicionarLog(
                "⚠ SSO sem nome"
            );
        }

    }
    catch (erro) {

        console.error(erro);

        campo.readOnly = false;

        adicionarLog(
            "⚠ SSO não identificado"
        );
    }
}


/* =========================================================
   TOKEN
========================================================= */

function decodificarToken(token) {

    const partes =
        token.split(".");


    if (partes.length !== 3) {

        throw new Error(
            "Token inválido."
        );
    }


    let payload =
        partes[1]
            .replace(/-/g, "+")
            .replace(/_/g, "/");


    while (
        payload.length % 4 !== 0
    ) {

        payload += "=";
    }


    return JSON.parse(
        decodeURIComponent(
            atob(payload)
                .split("")
                .map(function (c) {

                    return "%" +
                        (
                            "00" +
                            c.charCodeAt(0)
                                .toString(16)
                        ).slice(-2);

                })
                .join("")
        )
    );
}


/* =========================================================
   DATA + HORA
========================================================= */

function obterDataHoraAtual() {

    const agora = new Date();

    // Retorna um número serial do Excel usando a data/hora local do navegador.
    // Isso evita que o Excel interprete "dd/mm/yyyy" como "mm/dd/yyyy".
    const dataUTC = Date.UTC(
        agora.getFullYear(),
        agora.getMonth(),
        agora.getDate(),
        agora.getHours(),
        agora.getMinutes(),
        agora.getSeconds()
    );

    const epochExcel = Date.UTC(1899, 11, 30);

    return (dataUTC - epochExcel) / 86400000;
}


/* =========================================================
   SALVAR
========================================================= */

async function salvarEmail(event) {

    event.preventDefault();

    limparStatus();

    const btn =
        document.getElementById(
            "btnSalvar"
        );


    btn.disabled = true;

    btn.textContent =
        "Salvando...";


    try {

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


        const emailResposta =
            document.getElementById(
                "emailResposta"
            ).value.trim();


        const supervisor =
            document.getElementById(
                "supervisor"
            ).value.trim();


        const obs =
            document.getElementById(
                "obs"
            ).value.trim();


        const historicoExterno =
            document.getElementById(
                "historicoExterno"
            ).value.trim();


        const nomeMensagem =
            document.getElementById(
                "mensagem"
            ).value;


        const textoMensagem =
            document.getElementById(
                "textoMensagem"
            ).value;


        const assunto = document.getElementById("assunto").value.trim();
        const referenciaCampanha = document.getElementById("referenciaCampanha").value.trim();


        if (!empresa)
            throw new Error(
                "Informe a empresa."
            );


        if (
            !qtde ||
            Number(qtde) <= 0 ||
            !Number.isInteger(
                Number(qtde)
            )
        )
            throw new Error(
                "A Qtde deve ser um número inteiro maior que zero."
            );


        if (!emailResposta)
            throw new Error(
                "Informe o E-mail Resposta."
            );


        if (!supervisor)
            throw new Error(
                "Informe o Supervisor."
            );


        if (!historicoExterno)
            throw new Error(
                "Informe o Histórico Externo."
            );


        if (!nomeMensagem)
            throw new Error(
                "Selecione uma mensagem."
            );


        if (!textoMensagem)
            throw new Error(
                "A mensagem selecionada não possui texto."
            );


        if (!assunto) throw new Error("Informe o assunto.");


        await Excel.run(
            async function (context) {

                await adicionarRegistroEmail(
                    context,
                    {
                        realizadoPor,
                        empresa,
                        qtde: Number(qtde),
                        emailResposta,
                        supervisor,
                        obs,
                        historicoExterno,
                        textoMensagem,
                        assunto,
                        referenciaCampanha,
                        idRegistro: gerarIdRegistro()
                    }
                );


                await atualizarConfig(
                    context,
                    {
                        empresa,
                        emailResposta,
                        supervisor,
                        realizadoPor
                    }
                );


                await context.sync();
            }
        );


        mostrarStatus(
            "✓ Registro de E-mail salvo!",
            "sucesso"
        );


        document
            .getElementById(
                "formEmail"
            )
            .reset();


        document
            .getElementById(
                "mensagem"
            )
            .selectedIndex = 0;


        document.getElementById("textoMensagem").value = "";
        document.getElementById("referenciaCampanha").value = "";


        await identificarUsuario();

    }
    catch (erro) {

        console.error(erro);

        mostrarStatus(
            erro.message ||
            "Erro ao salvar.",
            "erro"
        );

    }
    finally {

        btn.disabled = false;

        btn.textContent =
            "Salvar";
    }
}


function validarEstruturaRegistro(mapa, camposObrigatorios, quantidadeColunasEsperadas) {
    const ausentes = camposObrigatorios.filter(function (nomes) {
        const indice = obterIndice(mapa, nomes);
        return indice === -1 || indice >= quantidadeColunasEsperadas;
    });

    if (ausentes.length > 0) {
        throw new Error(
            "A estrutura da aba não corresponde ao padrão esperado. Coluna(s) ausente(s) ou fora da estrutura: " +
            ausentes.map(function (nomes) { return nomes[0]; }).join(", ")
        );
    }
}


/* =========================================================
   ADICIONAR REGISTRO
========================================================= */

async function adicionarRegistroEmail(
    context,
    dados
) {

    const folha =
        context.workbook.worksheets.getItem(
            "Email"
        );


    const usado =
        folha.getUsedRangeOrNullObject(true);


    usado.load([
        "values",
        "isNullObject",
        "rowIndex",
        "columnIndex"
    ]);


    await context.sync();


    let linhaCabecalho = -1;

    let mapa = {};


    if (!usado.isNullObject) {

        for (
            let linha = 0;
            linha < Math.min(
                usado.values.length,
                20
            );
            linha++
        ) {

            const atual = {};


            usado.values[linha].forEach(
                function (
                    valor,
                    indice
                ) {

                    if (
                        valor !== null &&
                        valor !== undefined &&
                        String(valor).trim() !== ""
                    ) {

                        atual[
                            normalizar(valor)
                        ] = indice;
                    }
                }
            );


            if (
                atual["EMPRESA"] !== undefined &&
                (
                    atual["DATA"] !== undefined ||
                    atual["REALIZADO POR"] !== undefined
                )
            ) {

                linhaCabecalho =
                    linha;

                mapa = atual;

                break;
            }
        }
    }


    if (linhaCabecalho === -1) {

        folha
            .getRange("A1:J1")
            .values = [[
                "Data",
                "Realizado por",
                "Empresa",
                "Qtde",
                "Email Resposta",
                "Supervisor",
                "Obs",
                "Historico Externo",
                "Mensagem",
                "Assunto"
            ]];

        linhaCabecalho = 0;

        [
            "DATA",
            "REALIZADO POR",
            "EMPRESA",
            "QTDE",
            "EMAIL RESPOSTA",
            "SUPERVISOR",
            "OBS",
            "HISTORICO EXTERNO",
            "MENSAGEM",
            "ASSUNTO"
        ].forEach(function (
            nome,
            indice
        ) {
            mapa[nome] = indice;
        });
    }


    const numeroColunas = 12;
    if (mapa["ID REGISTRO"] === undefined) { folha.getRangeByIndexes(usado.rowIndex + linhaCabecalho, usado.columnIndex + 10, 1, 2).values = [["ID Registro", "Referência Campanha"]]; mapa["ID REGISTRO"] = 10; mapa["REFERÊNCIA CAMPANHA"] = 11; mapa["REFERENCIA CAMPANHA"] = 11; }
    validarEstruturaRegistro(mapa, [
        ["DATA"],
        ["REALIZADO POR"],
        ["EMPRESA"],
        ["QTDE"],
        ["EMAIL RESPOSTA", "E-MAIL RESPOSTA"],
        ["SUPERVISOR"],
        ["OBS"],
        ["HISTORICO EXTERNO", "HISTÓRICO EXTERNO"],
        ["MENSAGEM"],
        ["ASSUNTO"],
        ["ID REGISTRO"],
        ["REFERENCIA CAMPANHA", "REFERÊNCIA CAMPANHA"]
    ], numeroColunas);


    const valores =
        new Array(numeroColunas)
            .fill("");


    function colocar(
        nomes,
        valor
    ) {

        const indice =
            obterIndice(
                mapa,
                nomes
            );

        if (indice !== -1)
            valores[indice] = valor;
    }


    colocar(
        ["DATA"],
        obterDataHoraAtual()
    );

    colocar(
        ["REALIZADO POR"],
        dados.realizadoPor
    );

    colocar(
        ["EMPRESA"],
        dados.empresa
    );

    colocar(
        ["QTDE"],
        dados.qtde
    );

    colocar(
        [
            "EMAIL RESPOSTA",
            "E-MAIL RESPOSTA"
        ],
        dados.emailResposta
    );

    colocar(
        ["SUPERVISOR"],
        dados.supervisor
    );

    colocar(
        ["OBS"],
        dados.obs
    );

    colocar(
        [
            "HISTORICO EXTERNO",
            "HISTÓRICO EXTERNO"
        ],
        dados.historicoExterno
    );

    colocar(
        ["MENSAGEM"],
        dados.textoMensagem
    );

    colocar(["ASSUNTO"], dados.assunto);
    colocar(["ID REGISTRO"], dados.idRegistro);
    colocar(["REFERENCIA CAMPANHA", "REFERÊNCIA CAMPANHA"], dados.referenciaCampanha);


    const proximaLinha = obterProximaLinhaDados(usado, 10);
    const colunaInicial = usado.isNullObject ? 0 : usado.columnIndex;
    const intervaloRegistro = folha.getRangeByIndexes(
        proximaLinha,
        colunaInicial,
        1,
        numeroColunas
    );

    intervaloRegistro.values = [valores];

    // Mantém a coluna Data como data/hora real do Excel.
    intervaloRegistro.getCell(0, 0).numberFormat = [["dd/mm/yyyy hh:mm:ss"]];

    // Desativa a quebra automática somente na célula da Mensagem.
    const indiceMensagem = obterIndice(mapa, ["MENSAGEM"]);
    if (indiceMensagem !== -1) {
        intervaloRegistro.getCell(0, indiceMensagem).format.wrapText = false;
    }
}


function gerarIdRegistro() {
    const agora = new Date();
    const parte = Math.random().toString(36).slice(2, 7).toUpperCase();
    return "ENV-" + agora.getFullYear() + String(agora.getMonth()+1).padStart(2,"0") + String(agora.getDate()).padStart(2,"0") + "-" + Date.now().toString().slice(-6) + "-" + parte;
}

function obterProximaLinhaDados(usado, quantidadeColunasEsperadas) {
    if (usado.isNullObject) return 1;
    const valores = usado.values || [];
    for (let linha = valores.length - 1; linha >= 0; linha--) {
        const colunas = valores[linha] || [];
        const limite = Math.min(quantidadeColunasEsperadas, colunas.length);
        for (let coluna = 0; coluna < limite; coluna++) {
            const valor = colunas[coluna];
            if (valor !== null && valor !== undefined && String(valor).trim() !== "") {
                return usado.rowIndex + linha + 1;
            }
        }
    }
    return usado.rowIndex + 1;
}


/* =========================================================
   ATUALIZAR CONFIG
========================================================= */

async function atualizarConfig(
    context,
    dados
) {

    const folha =
        context.workbook.worksheets.getItem(
            "Config"
        );


    const usado =
        folha.getUsedRangeOrNullObject();


    usado.load([
        "values",
        "isNullObject",
        "rowIndex",
        "columnIndex"
    ]);


    await context.sync();


    if (usado.isNullObject)
        return;


    const info =
        encontrarCabecalho(
            usado.values
        );


    if (!info)
        return;


    const campos = [
        {
            nomes: [
                "EMPRESAS",
                "EMPRESA"
            ],
            valor: dados.empresa
        },
        {
            nomes: [
                "EMAILS_RESPOSTA",
                "EMAIL RESPOSTA",
                "E-MAIL RESPOSTA"
            ],
            valor: dados.emailResposta
        },
        {
            nomes: [
                "SUPERVISORES",
                "SUPERVISOR"
            ],
            valor: dados.supervisor
        },
        {
            nomes: [
                "REALIZADO POR",
                "REALIZADO_POR",
                "REALIZADOPOR"
            ],
            valor: dados.realizadoPor
        }
    ];


    for (const campo of campos) {

        if (!campo.valor)
            continue;


        const coluna =
            obterIndice(
                info.mapa,
                campo.nomes
            );


        if (coluna === -1)
            continue;


        let linhaDestino =
            info.linha + 1;


        while (
            linhaDestino <
                usado.values.length &&
            String(
                usado.values[
                    linhaDestino
                ][coluna] || ""
            ).trim() !== ""
        ) {

            linhaDestino++;
        }


        const existentes = [];


        for (
            let linha =
                info.linha + 1;
            linha < usado.values.length;
            linha++
        ) {

            existentes.push(
                String(
                    usado.values[
                        linha
                    ][coluna] || ""
                ).trim()
            );
        }


        const existe =
            existentes.some(
                function (valor) {
                    return (
                        normalizar(valor) ===
                        normalizar(
                            campo.valor
                        )
                    );
                }
            );


        if (existe)
            continue;


        folha
            .getRangeByIndexes(
                usado.rowIndex + linhaDestino,
                usado.columnIndex + coluna,
                1,
                1
            )
            .values = [
                [campo.valor]
            ];
    }
}

function preencherDuplicacao() {
    const bruto = localStorage.getItem("registroDuplicado");
    if (!bruto) return;
    try {
        const d = JSON.parse(bruto);
        if (d.tipo !== "email") return;
        const campos = {realizadoPor:"realizadoPor",empresa:"empresa",qtde:"qtde",emailResposta:"emailResposta",supervisor:"supervisor",obs:"obs",historicoExterno:"historicoExterno",referenciaCampanha:"referenciaCampanha",assunto:"assunto"};
        Object.keys(campos).forEach(k => { if (d[k] !== undefined) document.getElementById(campos[k]).value = d[k]; });
        const opcao = mensagensEmail.find(item => item.texto === d.mensagem);
        if (opcao) { document.getElementById("mensagem").value = opcao.nome; mostrarTextoMensagem(); }
        localStorage.removeItem("registroDuplicado");
        mostrarStatus("Registro carregado para duplicação. Revise os dados antes de salvar.", "aviso");
    } catch(e) { localStorage.removeItem("registroDuplicado"); }
}
