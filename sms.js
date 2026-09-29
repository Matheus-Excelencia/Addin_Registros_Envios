/* =========================================================
   SMS - Office Add-in
   Config:
   A = EMPRESAS
   B = EMAIL RESPOSTA
   C = SUPERVISORES
   D = REALIZADO POR
   E = NOME MENSAGEM EMAIL
   F = TEXTO MENSAGEM EMAIL
   G = NOME MENSAGEM SMS
   H = TEXTO MENSAGEM SMS
   I = LIMITE SMS

   Cabeçalho: linha 2
   Dados: linha 3 em diante
========================================================= */

Office.onReady(function () {

    console.log("SMS iniciado.");

    const form = document.getElementById("formSMS");
    const btnSalvar = document.getElementById("btnSalvar");
    const btnCancelar = document.getElementById("btnCancelar");

    if (btnSalvar) {
        // Evita que o botão dispare submit automaticamente
        btnSalvar.type = "button";

        btnSalvar.addEventListener("click", function (evento) {
            evento.preventDefault();
            salvarSMS();
        });
    }

    if (form) {
        form.addEventListener("submit", function (evento) {
            evento.preventDefault();
            salvarSMS();
        });
    }

    if (btnCancelar) {
        btnCancelar.addEventListener("click", function () {
            limparFormulario();
        });
    }

    carregarDadosConfig();
});


/* =========================================================
   FUNÇÕES AUXILIARES
========================================================= */

function logStatus(mensagem) {
    const campo = document.getElementById("mensagemStatus");

    if (campo) {
        campo.textContent = mensagem;
    }

    console.log(mensagem);
}


function obterValor(id) {
    const elemento = document.getElementById(id);

    if (!elemento) {
        return "";
    }

    return elemento.value.trim();
}


function preencherLista(id, valores) {

    const lista = document.getElementById(id);

    if (!lista) {
        console.warn("Lista não encontrada:", id);
        return;
    }

    lista.innerHTML = "";

    valores.forEach(function (valor) {

        if (!valor) {
            return;
        }

        const option = document.createElement("option");
        option.value = valor;

        lista.appendChild(option);
    });
}


function valorUnico(valores) {

    return [...new Set(
        valores
            .map(function (v) {
                return String(v || "").trim();
            })
            .filter(function (v) {
                return v !== "";
            })
    )];
}


/* =========================================================
   CARREGAR CONFIG
========================================================= */

async function carregarDadosConfig() {

    logStatus("Carregando Config...");

    try {

        await Excel.run(async function (context) {

            const config = context.workbook.worksheets.getItem("Config");

            /*
             * Linha 2 = cabeçalho
             * Vamos ler diretamente A2:I1000.
             */

            const intervalo = config.getRange("A2:I1000");

            intervalo.load("values");

            await context.sync();

            const dados = intervalo.values;

            if (!dados || dados.length === 0) {
                throw new Error("Não foi possível ler a aba Config.");
            }

            /*
             * A primeira linha do intervalo A2:I1000
             * corresponde à linha 2 da planilha.
             */

            const cabecalho = dados[0].map(function (valor) {
                return String(valor || "").trim().toUpperCase();
            });

            console.log("Cabeçalhos encontrados:", cabecalho);

            /*
             * Como sua estrutura é fixa, validamos
             * diretamente a primeira linha.
             */

            if (cabecalho[0] !== "EMPRESAS") {
                throw new Error(
                    'A célula A2 deveria ser "EMPRESAS", mas foi encontrado: "' +
                    cabecalho[0] +
                    '"'
                );
            }

            if (cabecalho[2] !== "SUPERVISORES") {
                throw new Error(
                    'A célula C2 deveria ser "SUPERVISORES", mas foi encontrado: "' +
                    cabecalho[2] +
                    '"'
                );
            }

            if (cabecalho[3] !== "REALIZADO POR") {
                throw new Error(
                    'A célula D2 deveria ser "REALIZADO POR", mas foi encontrado: "' +
                    cabecalho[3] +
                    '"'
                );
            }

            if (cabecalho[6] !== "NOME MENSAGEM SMS") {
                throw new Error(
                    'A célula G2 deveria ser "NOME MENSAGEM SMS", mas foi encontrado: "' +
                    cabecalho[6] +
                    '"'
                );
            }

            /*
             * Ignora a linha 2 e pega somente os dados.
             */

            const empresas = [];
            const supervisores = [];
            const realizadoPor = [];
            const nomesMensagens = [];
            const textosMensagens = [];

            for (let i = 1; i < dados.length; i++) {

                const linha = dados[i];

                if (linha[0]) {
                    empresas.push(String(linha[0]).trim());
                }

                if (linha[2]) {
                    supervisores.push(String(linha[2]).trim());
                }

                if (linha[3]) {
                    realizadoPor.push(String(linha[3]).trim());
                }

                if (linha[6]) {

                    nomesMensagens.push(
                        String(linha[6]).trim()
                    );

                    textosMensagens.push(
                        String(linha[7] || "").trim()
                    );
                }
            }

            const empresasUnicas = valorUnico(empresas);
            const supervisoresUnicos = valorUnico(supervisores);
            const realizadoUnicos = valorUnico(realizadoPor);

            preencherLista(
                "listaEmpresas",
                empresasUnicas
            );

            preencherLista(
                "listaSupervisores",
                supervisoresUnicos
            );

            preencherLista(
                "listaRealizadoPor",
                realizadoUnicos
            );

            /*
             * Mensagens SMS
             */

            const listaMensagens =
                document.getElementById("mensagem");

            if (listaMensagens) {

                listaMensagens.innerHTML =
                    '<option value="">Selecione uma mensagem</option>';

                for (let i = 0; i < nomesMensagens.length; i++) {

                    const option =
                        document.createElement("option");

                    option.value = nomesMensagens[i];

                    /*
                     * Guarda o texto completo da mensagem.
                     */

                    option.dataset.texto =
                        textosMensagens[i];

                    option.textContent =
                        nomesMensagens[i];

                    listaMensagens.appendChild(option);
                }
            }

            /*
             * Também coloca o limite padrão de 160.
             */

            const campoMensagem =
                document.getElementById("textoMensagem");

            if (campoMensagem) {

                campoMensagem.maxLength = 160;

                campoMensagem.addEventListener(
                    "input",
                    atualizarContador
                );
            }

            /*
             * Quando escolher uma mensagem,
             * preenche o texto completo.
             */

            if (listaMensagens) {

                listaMensagens.addEventListener(
                    "change",
                    function () {

                        const opcao =
                            listaMensagens.options[
                                listaMensagens.selectedIndex
                            ];

                        const texto =
                            opcao?.dataset?.texto || "";

                        if (campoMensagem) {
                            campoMensagem.value = texto;
                            atualizarContador();
                        }
                    }
                );
            }

            logStatus("Config carregada.");

        });

    } catch (erro) {

        console.error(erro);

        logStatus(
            "Erro ao carregar: " +
            (erro.message || erro)
        );
    }
}


/* =========================================================
   CONTADOR SMS
========================================================= */

function atualizarContador() {

    const campo =
        document.getElementById("textoMensagem");

    const contador =
        document.getElementById("contadorMensagem");

    if (!campo) {
        return;
    }

    const quantidade =
        campo.value.length;

    if (contador) {

        contador.textContent =
            quantidade + " / 160";
    }

    if (quantidade > 160) {

        campo.style.borderColor = "red";

    } else {

        campo.style.borderColor = "";
    }
}


/* =========================================================
   SALVAR SMS
========================================================= */

async function salvarSMS() {

    logStatus("Iniciando salvamento...");

    try {

        /*
         * CAPTURA DOS CAMPOS
         */

        const empresa =
            obterValor("empresa");

        const qtde =
            obterValor("qtde");

        const supervisor =
            obterValor("supervisor");

        const obs =
            obterValor("obs");

        const mensagemSelect =
            document.getElementById("mensagem");

        const textoMensagem =
            obterValor("textoMensagem");

        let nomeMensagem = "";

        if (mensagemSelect) {
            nomeMensagem =
                mensagemSelect.value.trim();
        }


        /*
         * VALIDAÇÕES
         */

        if (!empresa) {
            logStatus("Informe a empresa.");
            return;
        }

        if (!qtde) {
            logStatus("Informe a quantidade.");
            return;
        }

        const quantidadeNumero =
            Number(qtde);

        if (
            !Number.isInteger(quantidadeNumero) ||
            quantidadeNumero <= 0
        ) {
            logStatus(
                "A quantidade deve ser um número inteiro maior que zero."
            );
            return;
        }

        if (!supervisor) {
            logStatus("Informe o supervisor.");
            return;
        }

        if (!textoMensagem) {
            logStatus("Informe a mensagem.");
            return;
        }

        if (textoMensagem.length > 160) {

            logStatus(
                "A mensagem possui " +
                textoMensagem.length +
                " caracteres. O limite é 160."
            );

            return;
        }


        /*
         * REALIZADO POR
         */

        let realizadoPor =
            obterValor("realizadoPor");

        /*
         * Tenta pegar automaticamente
         * a conta Microsoft logada.
         */

        if (!realizadoPor) {

            try {

                const token =
                    await OfficeRuntime.auth.getAccessToken({
                        allowSignInPrompt: true,
                        allowConsentPrompt: true
                    });

                if (token) {

                    /*
                     * O token existe, mas não precisamos
                     * decodificar aqui.
                     *
                     * Se o campo estiver vazio,
                     * mantemos a possibilidade de
                     * preenchimento manual.
                     */

                    console.log(
                        "Token Microsoft obtido."
                    );
                }

            } catch (erroSSO) {

                console.warn(
                    "Não foi possível obter o usuário automaticamente.",
                    erroSSO
                );
            }
        }


        /*
         * DATA + HORA
         */

        const agora =
            new Date();

        const dataHora =
            formatarDataHora(agora);


        logStatus("1/4 - Preparando dados...");


        /*
         * SALVAR NA ABA SMS
         */

        await Excel.run(async function (context) {

            const sheet =
                context.workbook.worksheets.getItem("SMS");

            /*
             * Descobre a próxima linha usada.
             */

            const usado =
                sheet.getUsedRangeOrNullObject(true);

            usado.load([
                "isNullObject",
                "rowIndex",
                "rowCount"
            ]);

            await context.sync();

            let linhaDestino;

            if (usado.isNullObject) {

                /*
                 * Linha 3
                 * índice zero-based = 2
                 */

                linhaDestino = 2;

            } else {

                /*
                 * Próxima linha depois da usada.
                 */

                linhaDestino =
                    usado.rowIndex +
                    usado.rowCount;

                /*
                 * Nunca salvar acima da linha 3.
                 */

                if (linhaDestino < 2) {
                    linhaDestino = 2;
                }
            }


            logStatus("2/4 - Salvando na aba SMS...");


            /*
             * A:G
             *
             * A Data
             * B Realizado por
             * C Empresa
             * D Qtde
             * E Supervisor
             * F Obs
             * G Mensagem
             */

            const destino =
                sheet.getRangeByIndexes(
                    linhaDestino,
                    0,
                    1,
                    7
                );

            destino.values = [[
                dataHora,
                realizadoPor,
                empresa,
                quantidadeNumero,
                supervisor,
                obs,
                textoMensagem
            ]];

            await context.sync();


            /*
             * Confirma a gravação.
             */

            destino.load("values");

            await context.sync();

            console.log(
                "SMS salvo:",
                destino.values
            );


            logStatus("3/4 - Atualizando Config...");

        });


        /*
         * ATUALIZAR CONFIG
         */

        await adicionarConfig(
            "EMPRESAS",
            empresa
        );

        await adicionarConfig(
            "SUPERVISORES",
            supervisor
        );

        if (realizadoPor) {

            await adicionarConfig(
                "REALIZADO POR",
                realizadoPor
            );
        }


        logStatus(
            "4/4 - SMS salvo com sucesso!"
        );


        /*
         * LIMPA O FORMULÁRIO
         */

        limparFormulario();


    } catch (erro) {

        console.error(
            "ERRO AO SALVAR SMS:",
            erro
        );

        logStatus(
            "Erro ao salvar: " +
            (erro.message || erro)
        );
    }
}


/* =========================================================
   ADICIONAR VALOR NA CONFIG
========================================================= */

async function adicionarConfig(
    coluna,
    valor
) {

    if (!valor) {
        return;
    }

    await Excel.run(async function (context) {

        const config =
            context.workbook.worksheets.getItem("Config");

        /*
         * Define a coluna pela estrutura fixa.
         */

        let numeroColuna;

        if (coluna === "EMPRESAS") {

            numeroColuna = 0;

        } else if (coluna === "SUPERVISORES") {

            numeroColuna = 2;

        } else if (coluna === "REALIZADO POR") {

            numeroColuna = 3;

        } else {

            return;
        }


        /*
         * Lê a coluna usada para verificar duplicidade.
         */

        const usado =
            config.getUsedRangeOrNullObject(true);

        usado.load([
            "isNullObject",
            "rowIndex",
            "rowCount",
            "values"
        ]);

        await context.sync();


        /*
         * Verifica se já existe.
         */

        if (!usado.isNullObject) {

            const valores =
                usado.values || [];

            for (
                let i = 0;
                i < valores.length;
                i++
            ) {

                /*
                 * Converte o índice da coluna
                 * dentro do UsedRange.
                 */

                const indiceColuna =
                    numeroColuna -
                    usado.columnIndex;

                if (
                    indiceColuna >= 0 &&
                    indiceColuna <
                    valores[i].length
                ) {

                    const existente =
                        String(
                            valores[i][indiceColuna] || ""
                        )
                        .trim()
                        .toLowerCase();

                    if (
                        existente ===
                        String(valor)
                            .trim()
                            .toLowerCase()
                    ) {

                        console.log(
                            valor +
                            " já existe em " +
                            coluna
                        );

                        return;
                    }
                }
            }
        }


        /*
         * Próxima linha disponível.
         */

        let linhaDestino;

        if (usado.isNullObject) {

            /*
             * Linha 3
             */

            linhaDestino = 2;

        } else {

            linhaDestino =
                usado.rowIndex +
                usado.rowCount;

            if (linhaDestino < 2) {
                linhaDestino = 2;
            }
        }


        const destino =
            config.getRangeByIndexes(
                linhaDestino,
                numeroColuna,
                1,
                1
            );

        destino.values = [[valor]];

        await context.sync();

        console.log(
            "Config atualizada:",
            coluna,
            valor
        );
    });
}


/* =========================================================
   DATA E HORA
========================================================= */

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
   LIMPAR FORMULÁRIO
========================================================= */

function limparFormulario() {

    const ids = [
        "empresa",
        "qtde",
        "supervisor",
        "obs",
        "textoMensagem"
    ];

    ids.forEach(function (id) {

        const campo =
            document.getElementById(id);

        if (campo) {
            campo.value = "";
        }
    });


    const mensagem =
        document.getElementById("mensagem");

    if (mensagem) {
        mensagem.selectedIndex = 0;
    }


    atualizarContador();


    logStatus("Formulário limpo.");
}
