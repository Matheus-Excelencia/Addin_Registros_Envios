let configInfo = null;
let mensagensEmail = [];

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

    await carregarDados();

});


function mostrarStatus(texto, tipo) {

    const elemento = document.getElementById("mensagemStatus");

    elemento.textContent = texto;
    elemento.className = "status " + tipo;

}


function limparStatus() {

    const elemento = document.getElementById("mensagemStatus");

    elemento.textContent = "";
    elemento.className = "status";

}


async function carregarDados() {

    try {

        limparStatus();

        await Excel.run(async function (context) {

            const folha = context.workbook.worksheets.getItem("Config");

            const usado = folha.getUsedRangeOrNullObject();

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

            configInfo = encontrarCabecalho(
                usado.values
            );

            if (!configInfo) {
                throw new Error(
                    "Não foi possível localizar os cabeçalhos da aba Config."
                );
            }

            preencherLista(
                "listaEmpresas",
                obterValoresColuna(
                    usado.values,
                    configInfo,
                    "EMPRESAS"
                )
            );

            preencherLista(
                "listaEmails",
                obterValoresColuna(
                    usado.values,
                    configInfo,
                    "EMAILS_RESPOSTA"
                )
            );

            preencherLista(
                "listaSupervisores",
                obterValoresColuna(
                    usado.values,
                    configInfo,
                    "SUPERVISORES"
                )
            );

            preencherLista(
                "listaRealizadoPor",
                obterValoresColuna(
                    usado.values,
                    configInfo,
                    "REALIZADO POR"
                )
            );

            carregarMensagens(
                usado.values,
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

    for (let linha = 0; linha < Math.min(valores.length, 20); linha++) {

        const colunas = valores[linha];

        const mapa = {};

        colunas.forEach(function (valor, indice) {

            if (valor !== null && valor !== undefined && valor !== "") {

                mapa[normalizar(valor)] = indice;

            }

        });

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

        const indice = mapa[normalizar(nome)];

        if (indice !== undefined) {
            return indice;
        }

    }

    return -1;

}


function obterValoresColuna(valores, info, nome) {

    const indice = obterIndice(
        info.mapa,
        [nome]
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

        const valor = valores[linha][indice];

        if (
            valor !== null &&
            valor !== undefined &&
            String(valor).trim() !== ""
        ) {

            resultado.push(String(valor).trim());

        }

    }

    return [...new Set(resultado)];

}


function preencherLista(id, valores) {

    const lista = document.getElementById(id);

    lista.innerHTML = "";

    valores.forEach(function (valor) {

        const option = document.createElement("option");

        option.value = valor;

        lista.appendChild(option);

    });

}


function carregarMensagens(valores, info) {

    const indiceNome = obterIndice(
        info.mapa,
        [
            "NOME MENSAGEM EMAIL",
            "MENSAGENS"
        ]
    );

    const indiceTexto = obterIndice(
        info.mapa,
        [
            "TEXTO MENSAGEM EMAIL",
            "TEXTO COMPLETO"
        ]
    );

    const select = document.getElementById("mensagem");

    select.innerHTML = "";

    mensagensEmail = [];

    if (indiceNome === -1 || indiceTexto === -1) {

        const option = document.createElement("option");

        option.value = "";
        option.textContent =
            "Mensagens de E-mail não encontradas";

        select.appendChild(option);

        return;

    }

    const opcaoInicial = document.createElement("option");

    opcaoInicial.value = "";
    opcaoInicial.textContent =
        "Selecione uma mensagem";

    select.appendChild(opcaoInicial);

    for (
        let linha = info.linha + 1;
        linha < valores.length;
        linha++
    ) {

        const nome = valores[linha][indiceNome];
        const texto = valores[linha][indiceTexto];

        if (
            nome !== null &&
            nome !== undefined &&
            String(nome).trim() !== ""
        ) {

            const mensagem = {
                nome: String(nome).trim(),
                texto: String(texto || "")
            };

            mensagensEmail.push(mensagem);

            const option = document.createElement("option");

            option.value = mensagem.nome;
            option.textContent = mensagem.nome;

            select.appendChild(option);

        }

    }

}


function mostrarTextoMensagem() {

    const nome = document.getElementById("mensagem").value;

    const mensagem = mensagensEmail.find(
        function (item) {
            return item.nome === nome;
        }
    );

    document.getElementById("textoMensagem").value =
        mensagem ? mensagem.texto : "";

}


async function identificarUsuario() {

    const campo = document.getElementById("realizadoPor");

    try {

        const token = await Office.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true
        });

        const dados = decodificarToken(token);

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

    const partes = token.split(".");

    if (partes.length !== 3) {
        throw new Error("Token inválido.");
    }

    let payload = partes[1]
        .replace(/-/g, "+")
        .replace(/_/g, "/");

    while (payload.length % 4 !== 0) {
        payload += "=";
    }

    return JSON.parse(
        decodeURIComponent(
            atob(payload)
                .split("")
                .map(function (c) {
                    return "%" +
                        ("00" + c.charCodeAt(0).toString(16))
                            .slice(-2);
                })
                .join("")
        )
    );

}


async function salvarEmail(event) {

    event.preventDefault();

    limparStatus();

    const btnSalvar =
        document.getElementById("btnSalvar");

    btnSalvar.disabled = true;
    btnSalvar.textContent = "Salvando...";

    try {

        const realizadoPor =
            document.getElementById("realizadoPor").value.trim();

        const empresa =
            document.getElementById("empresa").value.trim();

        const qtde =
            document.getElementById("qtde").value.trim();

        const emailResposta =
            document.getElementById("emailResposta").value.trim();

        const supervisor =
            document.getElementById("supervisor").value.trim();

        const obs =
            document.getElementById("obs").value.trim();

        const historicoExterno =
            document.getElementById("historicoExterno").value.trim();

        const nomeMensagem =
            document.getElementById("mensagem").value;

        const textoMensagem =
            document.getElementById("textoMensagem").value;

        const assunto =
            document.getElementById("assunto").value.trim();


        if (!empresa) {
            throw new Error("Informe a empresa.");
        }

        if (!qtde || Number(qtde) <= 0 || !Number.isInteger(Number(qtde))) {
            throw new Error("A Qtde deve ser um número inteiro maior que zero.");
        }

        if (!emailResposta) {
            throw new Error("Informe o E-mail Resposta.");
        }

        if (!supervisor) {
            throw new Error("Informe o Supervisor.");
        }

        if (!historicoExterno) {
            throw new Error("Informe o Histórico Externo.");
        }

        if (!nomeMensagem) {
            throw new Error("Selecione uma mensagem.");
        }

        if (!textoMensagem) {
            throw new Error("A mensagem selecionada não possui texto.");
        }

        if (!assunto) {
            throw new Error("Informe o assunto.");
        }

        await Excel.run(async function (context) {

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
                    assunto
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

        });

        mostrarStatus(
            "Registro de E-mail salvo com sucesso!",
            "sucesso"
        );

        document.getElementById("formEmail").reset();

        await identificarUsuario();

        document.getElementById("mensagem").selectedIndex = 0;
        document.getElementById("textoMensagem").value = "";

    } catch (erro) {

        console.error(erro);

        mostrarStatus(
            erro.message || "Erro ao salvar o registro.",
            "erro"
        );

    } finally {

        btnSalvar.disabled = false;
        btnSalvar.textContent = "Salvar";

    }

}


async function adicionarRegistroEmail(context, dados) {

    const folha =
        context.workbook.worksheets.getItem("Email");

    const usado =
        folha.getUsedRangeOrNullObject();

    usado.load([
        "values",
        "rowCount",
        "columnCount",
        "isNullObject"
    ]);

    await context.sync();


    const cabecalhos = [
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
    ];


    let linhaCabecalho = -1;
    let mapa = {};

    if (!usado.isNullObject) {

        for (
            let linha = 0;
            linha < Math.min(usado.values.length, 20);
            linha++
        ) {

            const atual = {};

            usado.values[linha].forEach(
                function (valor, indice) {

                    if (
                        valor !== null &&
                        valor !== undefined &&
                        valor !== ""
                    ) {

                        atual[normalizar(valor)] = indice;

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

                linhaCabecalho = linha;
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

        mapa = {};

        cabecalhos.forEach(
            function (nome, indice) {
                mapa[nome] = indice;
            }
        );

    }


    const valores = new Array(
        Math.max(
            10,
            Object.keys(mapa).length
        )
    ).fill("");


    const dataAtual = new Date();

    function colocar(nomes, valor) {

        const indice = obterIndice(
            mapa,
            nomes
        );

        if (indice !== -1) {
            valores[indice] = valor;
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
        ["EMAIL RESPOSTA"],
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
        ["HISTORICO EXTERNO", "HISTÓRICO EXTERNO"],
        dados.historicoExterno
    );

    colocar(
        ["MENSAGEM"],
        dados.textoMensagem
    );

    colocar(
        ["ASSUNTO"],
        dados.assunto
    );


    const ultimaLinha =
        linhaCabecalho +
        (
            usado.isNullObject
                ? 1
                : usado.rowCount
        );

    const numeroColunas =
        Math.max(
            valores.length,
            usado.isNullObject
                ? 10
                : usado.columnCount
        );

    const linhaFinal =
        new Array(numeroColunas).fill("");

    valores.forEach(
        function (valor, indice) {
            linhaFinal[indice] = valor;
        }
    );


    folha
        .getRangeByIndexes(
            ultimaLinha,
            0,
            1,
            numeroColunas
        )
        .values = [linhaFinal];

}


async function atualizarConfig(context, dados) {

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
        return;
    }

    const info =
        encontrarCabecalho(usado.values);

    if (!info) {
        return;
    }


    const campos = [
        {
            nome: "EMPRESAS",
            valor: dados.empresa
        },
        {
            nome: "EMAILS_RESPOSTA",
            valor: dados.emailResposta
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

        const valoresColuna = [];

        for (
            let linha = info.linha + 1;
            linha < usado.values.length;
            linha++
        ) {

            valoresColuna.push(
                String(
                    usado.values[linha][coluna] || ""
                ).trim()
            );

        }

        const existe =
            valoresColuna.some(
                function (valor) {
                    return normalizar(valor) ===
                        normalizar(campo.valor);
                }
            );

        if (existe) {
            continue;
        }

        let linhaDestino =
            info.linha + 1;

        while (
            linhaDestino < usado.values.length &&
            String(
                usado.values[linhaDestino][coluna] || ""
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
            .values = [[campo.valor]];

    }

}
