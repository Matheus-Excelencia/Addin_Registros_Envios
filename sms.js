let configInfoSMS = null;
let mensagensSMS = [];
let limiteSMS = 160;


/* =========================================================
   INICIAR
========================================================= */

Office.onReady(async function () {

    document
        .getElementById("btnVoltar")
        .addEventListener("click", function () {

            window.location.href =
                "taskpane.html";
        });


    document
        .getElementById("btnCancelar")
        .addEventListener(
            "click",
            limparFormulario
        );


    document
        .getElementById("formSMS")
        .addEventListener(
            "submit",
            salvarSMS
        );


    document
        .getElementById("mensagem")
        .addEventListener(
            "change",
            mostrarTextoMensagem
        );


    document
        .getElementById("textoMensagem")
        .addEventListener(
            "input",
            atualizarContador
        );


    await carregarDados();

});


/* =========================================================
   STATUS
========================================================= */

function mostrarStatus(texto, tipo) {

    const elemento =
        document.getElementById(
            "mensagemStatus"
        );

    elemento.textContent = texto;

    elemento.className =
        "status " + tipo;
}


function limparStatus() {

    const elemento =
        document.getElementById(
            "mensagemStatus"
        );

    elemento.textContent = "";

    elemento.className =
        "status";
}


/* =========================================================
   LIMPAR FORMULÁRIO
========================================================= */

function limparFormulario() {

    document
        .getElementById("formSMS")
        .reset();


    document
        .getElementById("mensagem")
        .selectedIndex = 0;


    document
        .getElementById("textoMensagem")
        .value = "";


    atualizarContador();

    limparStatus();

    identificarUsuario();

}


/* =========================================================
   CARREGAR DADOS
========================================================= */

async function carregarDados() {

    try {

        limparStatus();


        await Excel.run(
            async function (context) {

                const folha =
                    context.workbook.worksheets.getItem(
                        "Config"
                    );


                await prepararConfigSMS(
                    context,
                    folha
                );


                const usado =
                    folha.getUsedRangeOrNullObject();


                usado.load([
                    "values",
                    "rowCount",
                    "columnCount",
                    "isNullObject",
                    "rowIndex",
                    "columnIndex"
                ]);


                await context.sync();


                if (usado.isNullObject) {

                    throw new Error(
                        "A aba Config está vazia."
                    );
                }


                configInfoSMS =
                    encontrarCabecalho(
                        usado.values
                    );


                if (!configInfoSMS) {

                    throw new Error(
                        "Não foi possível localizar os cabeçalhos da aba Config."
                    );
                }


                preencherLista(
                    "listaEmpresas",
                    obterValoresColuna(
                        usado.values,
                        configInfoSMS,
                        "EMPRESAS"
                    )
                );


                preencherLista(
                    "listaSupervisores",
                    obterValoresColuna(
                        usado.values,
                        configInfoSMS,
                        "SUPERVISORES"
                    )
                );


                preencherLista(
                    "listaRealizadoPor",
                    obterValoresColuna(
                        usado.values,
                        configInfoSMS,
                        "REALIZADO POR"
                    )
                );


                carregarMensagens(
                    usado.values,
                    configInfoSMS
                );


                carregarLimiteSMS(
                    usado.values,
                    configInfoSMS
                );

            }
        );


        await identificarUsuario();

    }
    catch (erro) {

        console.error(erro);


        const selectMensagem =
            document.getElementById(
                "mensagem"
            );


        selectMensagem.innerHTML = "";


        const option =
            document.createElement(
                "option"
            );


        option.value = "";

        option.textContent =
            "Não foi possível carregar as mensagens";


        selectMensagem.appendChild(option);


        mostrarStatus(
            "Não foi possível carregar todas as listas. Os campos ainda podem ser preenchidos manualmente.",
            "aviso"
        );


        await identificarUsuario();

    }
}


/* =========================================================
   PREPARAR CONFIG SMS
========================================================= */

async function prepararConfigSMS(
    context,
    folha
) {

    const usado =
        folha.getUsedRangeOrNullObject();


    usado.load([
        "values",
        "rowCount",
        "columnCount",
        "isNullObject"
    ]);


    await context.sync();


    if (usado.isNullObject) {
        return;
    }


    let linhaCabecalho = -1;


    for (
        let linha = 0;
        linha < Math.min(
            usado.values.length,
            20
        );
        linha++
    ) {

        const valores =
            usado.values[linha];


        const nomes =
            valores.map(function (valor) {

                return normalizar(valor);

            });


        if (
            nomes.includes("EMPRESAS") &&
            nomes.includes("SUPERVISORES")
        ) {

            linhaCabecalho =
                linha;

            break;
        }
    }


    if (linhaCabecalho === -1) {
        return;
    }


    const cabecalho =
        usado.values[linhaCabecalho];


    const nomesExistentes =
        cabecalho.map(function (valor) {

            return normalizar(valor);

        });


    const colunasNecessarias = [
        "NOME MENSAGEM SMS",
        "TEXTO MENSAGEM SMS",
        "LIMITE SMS"
    ];


    for (
        const nome of colunasNecessarias
    ) {

        if (
            nomesExistentes.includes(
                normalizar(nome)
            )
        ) {

            continue;
        }


        const coluna =
            usado.columnCount;


        folha
            .getRangeByIndexes(
                linhaCabecalho,
                coluna,
                1,
                1
            )
            .values = [[nome]];


        if (
            normalizar(nome) ===
            "LIMITE SMS"
        ) {

            folha
                .getRangeByIndexes(
                    linhaCabecalho + 1,
                    coluna,
                    1,
                    1
                )
                .values = [[160]];
        }


        usado.columnCount++;

        nomesExistentes.push(
            normalizar(nome)
        );
    }
}


/* =========================================================
   NORMALIZAR
========================================================= */

function normalizar(valor) {

    return String(valor || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .toUpperCase();
}


/* =========================================================
   ENCONTRAR CABEÇALHO
========================================================= */

function encontrarCabecalho(valores) {

    for (
        let linha = 0;
        linha < Math.min(
            valores.length,
            20
        );
        linha++
    ) {

        const colunas =
            valores[linha];


        const mapa = {};


        colunas.forEach(
            function (
                valor,
                indice
            ) {

                if (
                    valor !== null &&
                    valor !== undefined &&
                    valor !== ""
                ) {

                    mapa[
                        normalizar(valor)
                    ] = indice;

                }

            }
        );


        if (
            mapa["EMPRESAS"] !== undefined &&
            mapa["SUPERVISORES"] !== undefined
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

function obterIndice(
    mapa,
    nomes
) {

    for (
        const nome of nomes
    ) {

        const indice =
            mapa[
                normalizar(nome)
            ];


        if (
            indice !== undefined
        ) {

            return indice;

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
    nome
) {

    const indice =
        obterIndice(
            info.mapa,
            [nome]
        );


    if (indice === -1) {

        return [];

    }


    const resultado = [];


    for (
        let linha =
            info.linha + 1;
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


    lista.innerHTML = "";


    valores.forEach(
        function (valor) {

            const option =
                document.createElement(
                    "option"
                );


            option.value = valor;


            lista.appendChild(
                option
            );

        }
    );
}


/* =========================================================
   CARREGAR MENSAGENS SMS
========================================================= */

function carregarMensagens(
    valores,
    info
) {

    const indiceNome =
        obterIndice(
            info.mapa,
            [
                "NOME MENSAGEM SMS"
            ]
        );


    const indiceTexto =
        obterIndice(
            info.mapa,
            [
                "TEXTO MENSAGEM SMS"
            ]
        );


    const select =
        document.getElementById(
            "mensagem"
        );


    select.innerHTML = "";


    mensagensSMS = [];


    if (
        indiceNome === -1 ||
        indiceTexto === -1
    ) {

        const option =
            document.createElement(
                "option"
            );


        option.value = "";


        option.textContent =
            "Mensagens SMS não encontradas";


        select.appendChild(
            option
        );


        mostrarStatus(
            "As mensagens SMS não foram encontradas na aba Config.",
            "aviso"
        );


        return;
    }


    const opcaoInicial =
        document.createElement(
            "option"
        );


    opcaoInicial.value = "";


    opcaoInicial.textContent =
        "Selecione uma mensagem";


    select.appendChild(
        opcaoInicial
    );


    for (
        let linha =
            info.linha + 1;
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


            mensagensSMS.push(
                mensagem
            );


            const option =
                document.createElement(
                    "option"
                );


            option.value =
                mensagem.nome;


            option.textContent =
                mensagem.nome;


            select.appendChild(
                option
            );

        }
    }


    if (
        mensagensSMS.length === 0
    ) {

        mostrarStatus(
            "Nenhuma mensagem SMS foi cadastrada na aba Config.",
            "aviso"
        );
    }
}


/* =========================================================
   CARREGAR LIMITE SMS
========================================================= */

function carregarLimiteSMS(
    valores,
    info
) {

    const indice =
        obterIndice(
            info.mapa,
            ["LIMITE SMS"]
        );


    limiteSMS = 160;


    if (indice === -1) {

        atualizarContador();

        return;
    }


    for (
        let linha =
            info.linha + 1;
        linha < valores.length;
        linha++
    ) {

        const valor =
            Number(
                valores[linha][indice]
            );


        if (
            !isNaN(valor) &&
            valor > 0
        ) {

            limiteSMS =
                Math.min(
                    valor,
                    160
                );

            break;
        }
    }


    atualizarContador();
}


/* =========================================================
   MOSTRAR TEXTO DA MENSAGEM
========================================================= */

function mostrarTextoMensagem() {

    const nome =
        document.getElementById(
            "mensagem"
        ).value;


    const mensagem =
        mensagensSMS.find(
            function (item) {

                return (
                    item.nome === nome
                );

            }
        );


    document.getElementById(
        "textoMensagem"
    ).value =
        mensagem
            ? mensagem.texto
            : "";


    atualizarContador();
}


/* =========================================================
   CONTADOR
========================================================= */

function atualizarContador() {

    const campo =
        document.getElementById(
            "textoMensagem"
        );


    const contador =
        document.getElementById(
            "contadorMensagem"
        );


    if (!contador) {
        return;
    }


    const quantidade =
        campo.value.length;


    contador.textContent =
        quantidade +
        " / " +
        limiteSMS;


    if (
        quantidade > limiteSMS
    ) {

        contador.classList.add(
            "limite-excedido"
        );

    }
    else {

        contador.classList.remove(
            "limite-excedido"
        );

    }
}


/* =========================================================
   IDENTIFICAR USUÁRIO
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

        }

    }
    catch (erro) {

        console.log(
            "Identificação automática não disponível.",
            erro
        );


        campo.readOnly = false;

    }
}


/* =========================================================
   DECODIFICAR TOKEN
========================================================= */

function decodificarToken(token) {

    const partes =
        token.split(".");


    if (
        partes.length !== 3
    ) {

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
                .map(
                    function (c) {

                        return "%" +
                            (
                                "00" +
                                c.charCodeAt(0)
                                    .toString(16)
                            ).slice(-2);

                    }
                )
                .join("")
        )
    );
}


/* =========================================================
   DATA E HORA
========================================================= */

function obterDataHoraAtual() {

    const agora =
        new Date();


    const dia =
        String(
            agora.getDate()
        ).padStart(2, "0");


    const mes =
        String(
            agora.getMonth() + 1
        ).padStart(2, "0");


    const ano =
        agora.getFullYear();


    const hora =
        String(
            agora.getHours()
        ).padStart(2, "0");


    const minuto =
        String(
            agora.getMinutes()
        ).padStart(2, "0");


    const segundo =
        String(
            agora.getSeconds()
        ).padStart(2, "0");


    return (
        dia +
        "/" +
        mes +
        "/" +
        ano +
        " " +
        hora +
        ":" +
        minuto +
        ":" +
        segundo
    );
}


/* =========================================================
   SALVAR SMS
========================================================= */

async function salvarSMS(event) {

    event.preventDefault();


    limparStatus();


    const btnSalvar =
        document.getElementById(
            "btnSalvar"
        );


    btnSalvar.disabled = true;


    btnSalvar.textContent =
        "Salvando...";


    try {

        const realizadoPor =
            document
                .getElementById(
                    "realizadoPor"
                )
                .value
                .trim();


        const empresa =
            document
                .getElementById(
                    "empresa"
                )
                .value
                .trim();


        const qtde =
            document
                .getElementById(
                    "qtde"
                )
                .value
                .trim();


        const supervisor =
            document
                .getElementById(
                    "supervisor"
                )
                .value
                .trim();


        const obs =
            document
                .getElementById(
                    "obs"
                )
                .value
                .trim();


        const nomeMensagem =
            document
                .getElementById(
                    "mensagem"
                )
                .value;


        const textoMensagem =
            document
                .getElementById(
                    "textoMensagem"
                )
                .value;


        if (!empresa) {

            throw new Error(
                "Informe a empresa."
            );

        }


        if (
            !qtde ||
            Number(qtde) <= 0 ||
            !Number.isInteger(
                Number(qtde)
            )
        ) {

            throw new Error(
                "A Qtde deve ser um número inteiro maior que zero."
            );

        }


        if (!supervisor) {

            throw new Error(
                "Informe o Supervisor."
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
            limiteSMS
        ) {

            throw new Error(
                "A mensagem SMS possui " +
                textoMensagem.length +
                " caracteres. O limite é de " +
                limiteSMS +
                " caracteres."
            );

        }


        await Excel.run(
            async function (context) {

                await adicionarRegistroSMS(
                    context,
                    {
                        realizadoPor,
                        empresa,
                        qtde:
                            Number(qtde),
                        supervisor,
                        obs,
                        textoMensagem
                    }
                );


                await atualizarConfigSMS(
                    context,
                    {
                        empresa,
                        supervisor,
                        realizadoPor
                    }
                );


                await context.sync();

            }
        );


        mostrarStatus(
            "Registro de SMS salvo com sucesso!",
            "sucesso"
        );


        document
            .getElementById(
                "formSMS"
            )
            .reset();


        document
            .getElementById(
                "mensagem"
            )
            .selectedIndex = 0;


        document
            .getElementById(
                "textoMensagem"
            )
            .value = "";


        atualizarContador();


        await identificarUsuario();

    }
    catch (erro) {

        console.error(erro);


        mostrarStatus(
            erro.message ||
            "Erro ao salvar o registro.",
            "erro"
        );

    }
    finally {

        btnSalvar.disabled = false;

        btnSalvar.textContent =
            "Salvar";

    }
}


/* =========================================================
   ADICIONAR REGISTRO SMS
========================================================= */

async function adicionarRegistroSMS(
    context,
    dados
) {

    const folha =
        context.workbook.worksheets.getItem(
            "SMS"
        );


    const usado =
        folha.getUsedRangeOrNullObject();


    usado.load([
        "values",
        "rowCount",
        "columnCount",
        "isNullObject",
        "rowIndex",
        "columnIndex"
    ]);


    await context.sync();


    const cabecalhos = [

        "DATA",
        "REALIZADO POR",
        "EMPRESA",
        "QTDE",
        "SUPERVISOR",
        "OBS",
        "MENSAGEM"

    ];


    let linhaCabecalho = -1;

    let mapa = {};


    /* =====================================================
       LOCALIZAR CABEÇALHO
    ===================================================== */

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


            usado.values[linha]
                .forEach(
                    function (
                        valor,
                        indice
                    ) {

                        if (
                            valor !== null &&
                            valor !== undefined &&
                            valor !== ""
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


    /* =====================================================
       CRIAR CABEÇALHO SE NECESSÁRIO
    ===================================================== */

    if (
        linhaCabecalho === -1
    ) {

        folha
            .getRange("A1:G1")
            .values = [[

                "Data",
                "Realizado por",
                "Empresa",
                "Qtde",
                "Supervisor",
                "Obs",
                "Mensagem"

            ]];


        linhaCabecalho = 0;


        mapa = {};


        cabecalhos.forEach(
            function (
                nome,
                indice
            ) {

                mapa[nome] =
                    indice;

            }
        );
    }


    /* =====================================================
       MONTAR LINHA
    ===================================================== */

    const numeroColunas =
        Math.max(
            7,
            usado.isNullObject
                ? 7
                : usado.columnCount
        );


    const valores =
        new Array(
            numeroColunas
        ).fill("");


    const dataAtual =
        obterDataHoraAtual();


    function colocar(
        nomes,
        valor
    ) {

        const indice =
            obterIndice(
                mapa,
                nomes
            );


        if (
            indice !== -1
        ) {

            valores[indice] =
                valor;

        }
    }


    colocar(
        ["DATA"],
        dataAtual
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
        ["SUPERVISOR"],
        dados.supervisor
    );


    colocar(
        ["OBS"],
        dados.obs
    );


    colocar(
        ["MENSAGEM"],
        dados.textoMensagem
    );


    /* =====================================================
       CORREÇÃO PRINCIPAL:
       SEMPRE PEGAR A PRÓXIMA LINHA LIVRE
    ===================================================== */

    let proximaLinha;


    if (
        usado.isNullObject
    ) {

        proximaLinha = 1;

    }
    else {

        proximaLinha =
            usado.rowIndex +
            usado.rowCount;

    }


    /* =====================================================
       GRAVAR NOVO REGISTRO
    ===================================================== */

    folha
        .getRangeByIndexes(
            proximaLinha,
            0,
            1,
            numeroColunas
        )
        .values = [
            valores
        ];
}


/* =========================================================
   ATUALIZAR CONFIG SMS
========================================================= */

async function atualizarConfigSMS(
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
        "rowCount",
        "columnCount",
        "isNullObject",
        "rowIndex",
        "columnIndex"
    ]);


    await context.sync();


    if (
        usado.isNullObject
    ) {

        return;

    }


    const info =
        encontrarCabecalho(
            usado.values
        );


    if (!info) {

        return;

    }


    const campos = [

        {
            nome: "EMPRESAS",
            valor: dados.empresa
        },

        {
            nome: "SUPERVISORES",
            valor: dados.supervisor
        },

        {
            nome: "REALIZADO POR",
            valor: dados.realizadoPor
        }

    ];


    for (
        const campo of campos
    ) {

        if (!campo.valor) {

            continue;

        }


        const coluna =
            obterIndice(
                info.mapa,
                [campo.nome]
            );


        if (
            coluna === -1
        ) {

            continue;

        }


        const valoresColuna = [];


        for (
            let linha =
                info.linha + 1;
            linha < usado.values.length;
            linha++
        ) {

            valoresColuna.push(
                String(
                    usado.values[linha][coluna] ||
                    ""
                ).trim()
            );

        }


        const existe =
            valoresColuna.some(
                function (valor) {

                    return (
                        normalizar(valor) ===
                        normalizar(
                            campo.valor
                        )
                    );

                }
            );


        if (
            existe
        ) {

            continue;

        }


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


        const linhaAbsoluta =
            usado.rowIndex +
            linhaDestino;


        folha
            .getRangeByIndexes(
                linhaAbsoluta,
                coluna,
                1,
                1
            )
            .values = [
                [campo.valor]
            ];

    }
}
