let configInfo = null;
let mensagensSMS = [];
let limiteSMS = 160;


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
        .getElementById("formSms")
        .addEventListener("submit", salvarSMS);

    document
        .getElementById("mensagem")
        .addEventListener("change", mostrarTextoMensagem);

    await carregarDados();

});


function mostrarStatus(texto, tipo) {

    const elemento =
        document.getElementById("mensagemStatus");

    elemento.textContent = texto;
    elemento.className = "status " + tipo;

}


function limparStatus() {

    const elemento =
        document.getElementById("mensagemStatus");

    elemento.textContent = "";
    elemento.className = "status";

}


async function carregarDados() {

    try {

        limparStatus();

        await Excel.run(async function (context) {

            const folha =
                context.workbook.worksheets.getItem("Config");

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
                throw new Error("A aba Config está vazia.");
            }

            configInfo =
                encontrarCabecalho(usado.values);

            if (!configInfo) {
                throw new Error(
                    "Não foi possível localizar os cabeçalhos da aba Config."
                );
            }


            await prepararConfigSMS(
                context,
                folha,
                usado,
                configInfo
            );

            await context.sync();


            const usadoAtualizado =
                folha.getUsedRange();

            usadoAtualizado.load("values");

            await context.sync();

            configInfo =
                encontrarCabecalho(
                    usadoAtualizado.values
                );


            preencherLista(
                "listaEmpresas",
                obterValoresColuna(
                    usadoAtualizado.values,
                    configInfo,
                    "EMPRESAS"
                )
            );


            preencherLista(
                "listaSupervisores",
                obterValoresColuna(
                    usadoAtualizado.values,
                    configInfo,
                    "SUPERVISORES"
                )
            );


            preencherLista(
                "listaRealizadoPor",
                obterValoresColuna(
                    usadoAtualizado.values,
                    configInfo,
                    "REALIZADO POR"
                )
            );


            carregarMensagens(
                usadoAtualizado.values,
                configInfo
            );


            carregarLimiteSMS(
                usadoAtualizado.values,
                configInfo
            );

        });


        await identificarUsuario();


    } catch (erro) {

        console.error(erro);

        mostrarStatus(
            "Não foi possível carregar todas as listas. Os campos ainda podem ser preenchidos manualmente.",
            "aviso"
        );

        await identificarUsuario();

    }

}


function normalizar(valor) {

    return String(valor || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .toUpperCase();

}


function encontrarCabecalho(valores) {

    for (
        let linha = 0;
        linha < Math.min(valores.length, 20);
        linha++
    ) {

        const colunas = valores[linha];

        const mapa = {};

        colunas.forEach(
            function (valor, indice) {

                if (
                    valor !== null &&
                    valor !== undefined &&
                    valor !== ""
                ) {

                    mapa[normalizar(valor)] = indice;

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


function obterIndice(mapa, nomes) {

    for (const nome of nomes) {

        const indice =
            mapa[normalizar(nome)];

        if (indice !== undefined) {
            return indice;
        }

    }

    return -1;

}


function obterValoresColuna(
    valores,
    info,
    nome
) {

    const indice =
        obterIndice(info.mapa, [nome]);

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

    return [...new Set(resultado)];

}


function preencherLista(id, valores) {

    const lista =
        document.getElementById(id);

    lista.innerHTML = "";

    valores.forEach(
        function (valor) {

            const option =
                document.createElement("option");

            option.value = valor;

            lista.appendChild(option);

        }
    );

}


async function prepararConfigSMS(
    context,
    folha,
    usado,
    info
) {

    let ultimaColuna =
        usado.columnCount;

    const indiceNome =
        obterIndice(
            info.mapa,
            ["NOME MENSAGEM SMS"]
        );

    const indiceTexto =
        obterIndice(
            info.mapa,
            ["TEXTO MENSAGEM SMS"]
        );

    const indiceLimite =
        obterIndice(
            info.mapa,
            ["LIMITE SMS"]
        );


    if (indiceNome === -1) {

        folha
            .getRangeByIndexes(
                info.linha,
                ultimaColuna,
                1,
                1
            )
            .values = [
                ["NOME MENSAGEM SMS"]
            ];

        ultimaColuna++;

    }


    if (indiceTexto === -1) {

        folha
            .getRangeByIndexes(
                info.linha,
                ultimaColuna,
                1,
                1
            )
            .values = [
                ["TEXTO MENSAGEM SMS"]
            ];

        ultimaColuna++;

    }


    if (indiceLimite === -1) {

        folha
            .getRangeByIndexes(
                info.linha,
                ultimaColuna,
                1,
                1
            )
            .values = [
                ["LIMITE SMS"]
            ];

        folha
            .getRangeByIndexes(
                info.linha + 1,
                ultimaColuna,
                1,
                1
            )
            .values = [[160]];

    }

}


function carregarMensagens(
    valores,
    info
) {

    const indiceNome =
        obterIndice(
            info.mapa,
            ["NOME MENSAGEM SMS"]
        );

    const indiceTexto =
        obterIndice(
            info.mapa,
            ["TEXTO MENSAGEM SMS"]
        );


    const select =
        document.getElementById("mensagem");

    select.innerHTML = "";

    mensagensSMS = [];


    if (
        indiceNome === -1 ||
        indiceTexto === -1
    ) {

        const option =
            document.createElement("option");

        option.value = "";

        option.textContent =
            "Mensagens SMS não encontradas";

        select.appendChild(option);

        return;

    }


    const opcaoInicial =
        document.createElement("option");

    opcaoInicial.value = "";

    opcaoInicial.textContent =
        "Selecione uma mensagem";

    select.appendChild(opcaoInicial);


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

                nome: String(nome).trim(),

                texto: String(texto || "")

            };


            mensagensSMS.push(mensagem);


            const option =
                document.createElement("option");

            option.value =
                mensagem.nome;

            option.textContent =
                mensagem.nome;

            select.appendChild(option);

        }

    }

}


function carregarLimiteSMS(
    valores,
    info
) {

    const indice =
        obterIndice(
            info.mapa,
            ["LIMITE SMS"]
        );

    if (indice === -1) {

        limiteSMS = 160;

        return;

    }


    const valor =
        Number(
            valores[info.linha + 1]?.[indice]
        );


    if (
        Number.isFinite(valor) &&
        valor > 0
    ) {

        limiteSMS =
            Math.min(valor, 160);

    } else {

        limiteSMS = 160;

    }

}


function mostrarTextoMensagem() {

    const nome =
        document.getElementById("mensagem").value;

    const mensagem =
        mensagensSMS.find(
            function (item) {
                return item.nome === nome;
            }
        );


    const texto =
        mensagem
            ? mensagem.texto
            : "";


    document.getElementById(
        "textoMensagem"
    ).value = texto;


    atualizarContador();

}


function atualizarContador() {

    const campo =
        document.getElementById("textoMensagem");

    const contador =
        document.getElementById("contador");

    const quantidade =
        campo.value.length;


    contador.textContent =
        quantidade + " / " + limiteSMS;


    if (quantidade > limiteSMS) {

        contador.classList.add("limite");

    } else {

        contador.classList.remove("limite");

    }

}


async function identificarUsuario() {

    const campo =
        document.getElementById("realizadoPor");


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


    } catch (erro) {

        console.log(
            "Identificação automática não disponível.",
            erro
        );

        campo.readOnly = false;

    }

}


function decodificarToken(token) {

    const partes =
        token.split(".");

    if (partes.length !== 3) {
        throw new Error("Token inválido.");
    }


    let payload =
        partes[1]
            .replace(/-/g, "+")
            .replace(/_/g, "/");


    while (payload.length % 4 !== 0) {
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


async function salvarSMS(event) {

    event.preventDefault();

    limparStatus();


    const btnSalvar =
        document.getElementById("btnSalvar");

    btnSalvar.disabled = true;
    btnSalvar.textContent = "Salvando...";


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


        const supervisor =
            document.getElementById(
                "supervisor"
            ).value.trim();


        const obs =
            document.getElementById(
                "obs"
            ).value.trim();


        const nomeMensagem =
            document.getElementById(
                "mensagem"
            ).value;


        const textoMensagem =
            document.getElementById(
                "textoMensagem"
            ).value;


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
            textoMensagem.length > limiteSMS
        ) {

            throw new Error(
                "A mensagem possui " +
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
                        qtde: Number(qtde),
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
            .getElementById("formSms")
            .reset();


        await identificarUsuario();


        document
            .getElementById("mensagem")
            .selectedIndex = 0;


        document
            .getElementById("textoMensagem")
            .value = "";


        atualizarContador();


    } catch (erro) {

        console.error(erro);

        mostrarStatus(
            erro.message ||
            "Erro ao salvar o registro.",
            "erro"
        );

    } finally {

        btnSalvar.disabled = false;

        btnSalvar.textContent =
            "Salvar";

    }

}


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
        "isNullObject"
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
                        valor !== ""
                    ) {

                        atual[
                            normalizar(valor)
                        ] = indice;

                    }

                }
            );


            if (
                atual["EMPRESA"] !== undefined
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


        mapa = {

            "DATA": 0,
            "REALIZADO POR": 1,
            "EMPRESA": 2,
            "QTDE": 3,
            "SUPERVISOR": 4,
            "OBS": 5,
            "MENSAGEM": 6

        };

    }


    const numeroColunas =
        Math.max(
            7,
            usado.isNullObject
                ? 7
                : usado.columnCount
        );


    const linhaFinal =
        new Array(
            numeroColunas
        ).fill("");


    function colocar(
        nomes,
        valor
    ) {

        const indice =
            obterIndice(
                mapa,
                nomes
            );

        if (indice !== -1) {

            linhaFinal[indice] =
                valor;

        }

    }


    colocar(
        ["DATA"],
        new Date()
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


    const proximaLinha =
        linhaCabecalho +
        (
            usado.isNullObject
                ? 1
                : usado.rowCount
        );


    folha
        .getRangeByIndexes(
            proximaLinha,
            0,
            1,
            numeroColunas
        )
        .values = [
            linhaFinal
        ];

}


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
        "isNullObject"
    ]);


    await context.sync();


    if (usado.isNullObject) {
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


    for (const campo of campos) {

        if (!campo.valor) {
            continue;
        }


        const coluna =
            obterIndice(
                info.mapa,
                [campo.nome]
            );


        if (coluna === -1) {
            continue;
        }


        let existe = false;


        for (
            let linha = info.linha + 1;
            linha < usado.values.length;
            linha++
        ) {

            const valor =
                String(
                    usado.values[linha][coluna] ||
                    ""
                ).trim();


            if (
                normalizar(valor) ===
                normalizar(campo.valor)
            ) {

                existe = true;
                break;

            }

        }


        if (existe) {
            continue;
        }


        let linhaDestino =
            info.linha + 1;


        while (
            linhaDestino < usado.values.length &&
            String(
                usado.values[linhaDestino][coluna] ||
                ""
            ).trim() !== ""
        ) {

            linhaDestino++;

        }


        folha
            .getRangeByIndexes(
                linhaDestino,
                coluna,
                1,
                1
            )
            .values = [
                [campo.valor]
            ];

    }

}


function limparFormulario() {

    document
        .getElementById("formSms")
        .reset();


    document
        .getElementById("mensagem")
        .selectedIndex = 0;


    document
        .getElementById("textoMensagem")
        .value = "";


    atualizarContador();


    identificarUsuario();

}
