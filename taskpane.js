/* =========================================================
   REGISTROS DE ENVIOS - TASKPANE
   Excel Add-in
   SSO Microsoft 365 + fallback manual + diagnóstico
   ========================================================= */

let config = {
  empresas: [],
  emailsResposta: [],
  supervisores: [],
  realizadoPor: [],
  nomesMensagem: [],
  textosMensagem: []
};

let configHeaders = {};


/* =========================================================
   INICIALIZAÇÃO
   ========================================================= */

Office.onReady(async (info) => {

  if (info.host !== Office.HostType.Excel) {

    mostrarStatus(
      "Este suplemento foi desenvolvido para o Excel.",
      true
    );

    return;
  }

  try {

    await carregarConfig();

    preencherCampos();

    await identificarUsuario();

    configurarEventos();

  } catch (erro) {

    console.error(
      "Erro na inicialização:",
      erro
    );

    mostrarStatus(
      "Erro ao carregar a configuração: " +
      obterMensagemErro(erro),
      true
    );

  }

});


/* =========================================================
   CARREGAR CONFIG
   ========================================================= */

async function carregarConfig() {

  await Excel.run(async (context) => {

    const planilha =
      context.workbook.worksheets.getItem("Config");

    const intervalo =
      planilha.getUsedRange();

    intervalo.load([
      "values",
      "rowCount",
      "columnCount"
    ]);

    await context.sync();

    const valores =
      intervalo.values;

    if (
      !valores ||
      valores.length === 0
    ) {

      throw new Error(
        "A aba Config está vazia."
      );

    }

    const cabecalhos =
      valores[0].map(valor =>
        String(valor || "").trim()
      );

    configHeaders = {};

    cabecalhos.forEach(
      (cabecalho, indice) => {

        const chave =
          normalizar(cabecalho);

        if (chave) {
          configHeaders[chave] =
            indice;
        }

      }
    );

    const dados =
      valores.slice(1);

    config.empresas =
      obterColuna(
        dados,
        "EMPRESAS"
      );

    config.emailsResposta =
      obterColuna(
        dados,
        "EMAIL RESPOSTA"
      );

    config.supervisores =
      obterColuna(
        dados,
        "SUPERVISORES"
      );

    config.realizadoPor =
      obterColuna(
        dados,
        "REALIZADO POR"
      );

    config.nomesMensagem =
      obterColuna(
        dados,
        "NOME MENSAGEM"
      );

    config.textosMensagem =
      obterColuna(
        dados,
        "TEXTO MENSAGEM"
      );

  });

}


/* =========================================================
   OBTER COLUNA DA CONFIG
   ========================================================= */

function obterColuna(
  dados,
  nomeColuna
) {

  const indice =
    configHeaders[
      normalizar(nomeColuna)
    ];

  if (indice === undefined) {
    return [];
  }

  return dados
    .map(linha =>
      String(
        linha[indice] || ""
      ).trim()
    )
    .filter(valor =>
      valor !== ""
    );

}


/* =========================================================
   PREENCHER CAMPOS
   ========================================================= */

function preencherCampos() {

  preencherDatalist(
    "empresas",
    config.empresas
  );

  preencherDatalist(
    "emailsResposta",
    config.emailsResposta
  );

  preencherDatalist(
    "supervisores",
    config.supervisores
  );

  preencherDatalist(
    "realizadoPorLista",
    config.realizadoPor
  );

  preencherMensagens();

}


/* =========================================================
   PREENCHER DATALIST
   ========================================================= */

function preencherDatalist(
  id,
  valores
) {

  const lista =
    document.getElementById(id);

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


/* =========================================================
   MENSAGENS
   ========================================================= */

function preencherMensagens() {

  const select =
    document.getElementById(
      "mensagem"
    );

  if (!select) {
    return;
  }

  select.innerHTML = "";

  const opcaoInicial =
    document.createElement(
      "option"
    );

  opcaoInicial.value = "";

  opcaoInicial.textContent =
    "Selecione a mensagem";

  select.appendChild(
    opcaoInicial
  );

  config.nomesMensagem.forEach(
    (nome, indice) => {

      const option =
        document.createElement(
          "option"
        );

      option.value = nome;

      option.textContent = nome;

      option.dataset.texto =
        config.textosMensagem[
          indice
        ] || "";

      select.appendChild(
        option
      );

    }
  );

}


/* =========================================================
   IDENTIFICAR USUÁRIO
   ========================================================= */

async function identificarUsuario() {

  const campo =
    document.getElementById(
      "realizadoPor"
    );

  if (!campo) {
    return;
  }

  try {
         const identityApi =
      Office.context.requirements.isSetSupported(
        "IdentityAPI",
        "1.3"
      );

    const getAccessTokenDisponivel =
      typeof Office.auth?.getAccessToken === "function";

    const versaoOffice =
      Office.context.diagnostics?.version || "não informada";

    const hostOffice =
      Office.context.host || "não informado";

    console.log("DIAGNÓSTICO SSO:", {
      identityApi: identityApi,
      getAccessToken: getAccessTokenDisponivel,
      versaoOffice: versaoOffice,
      hostOffice: hostOffice
    });

    mostrarStatus(
      "Diagnóstico SSO — IdentityAPI 1.3: " +
      (identityApi ? "SIM" : "NÃO") +
      " | getAccessToken: " +
      (getAccessTokenDisponivel ? "SIM" : "NÃO")
    );

    if (!identityApi) {
      throw criarErroSSO(
        "IDENTITY_API_NAO_SUPORTADA",
        "O Excel não informou suporte à IdentityAPI 1.3."
      );
    }

    if (!getAccessTokenDisponivel) {
      throw criarErroSSO(
        "GET_ACCESS_TOKEN_NAO_DISPONIVEL",
        "Office.auth.getAccessToken não está disponível."
      );
    }

    mostrarStatus(
      "Identificando usuário Microsoft 365..."
    );

    const token =
      await obterTokenSSO();

    const dadosUsuario =
      decodificarToken(token);

    const nome =
      dadosUsuario.name ||
      dadosUsuario.preferred_username ||
      "";

    const email =
      dadosUsuario.preferred_username ||
      "";

    if (
      !nome &&
      !email
    ) {

      throw criarErroSSO(
        "IDENTIDADE_NAO_ENCONTRADA",
        "O token foi obtido, mas não contém nome ou e-mail do usuário."
      );

    }

    /*
      Usuário identificado automaticamente.
    */

    campo.value =
      nome || email;

    campo.readOnly = true;

    campo.disabled = false;

    campo.style.backgroundColor =
      "#f3f4f6";

    campo.title =
      "Usuário identificado automaticamente pelo Microsoft 365.";

    campo.removeAttribute("list");

    mostrarStatus(
      "Usuário identificado: " +
      (nome || email)
    );

    console.log(
      "SSO OK:",
      {
        nome: nome,
        email: email,
        oid: dadosUsuario.oid,
        tid: dadosUsuario.tid
      }
    );

  } catch (erro) {

    registrarErroSSO(erro);

    ativarFallbackRealizadoPor();

  }

}


/* =========================================================
   OBTER TOKEN SSO
   ========================================================= */

async function obterTokenSSO() {

  try {

    if (
      Office.auth &&
      typeof Office.auth.getAccessToken ===
        "function"
    ) {

      return await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true
      });

    }

  } catch (erro) {

    throw criarErroSSO(
      obterCodigoErro(erro),
      obterMensagemErro(erro),
      erro
    );

  }


  try {

    if (
      typeof OfficeRuntime !== "undefined" &&
      OfficeRuntime.auth &&
      typeof OfficeRuntime.auth.getAccessToken ===
        "function"
    ) {

      return await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true
      });

    }

  } catch (erro) {

    throw criarErroSSO(
      obterCodigoErro(erro),
      obterMensagemErro(erro),
      erro
    );

  }


  throw criarErroSSO(
    "API_NAO_DISPONIVEL",
    "A API de autenticação do Office não está disponível."
  );

}


/* =========================================================
   CRIAR ERRO SSO
   ========================================================= */

function criarErroSSO(
  codigo,
  mensagem,
  original = null
) {

  const erro =
    new Error(mensagem);

  erro.code =
    codigo || "SEM_CODIGO";

  erro.original =
    original;

  erro.isSSOError =
    true;

  return erro;

}


/* =========================================================
   REGISTRAR ERRO SSO
   ========================================================= */

function registrarErroSSO(erro) {

  const codigo =
    obterCodigoErro(erro);

  const mensagem =
    obterMensagemErro(erro);

  console.group(
    "Diagnóstico SSO"
  );

  console.error(
    "Código:",
    codigo
  );

  console.error(
    "Mensagem:",
    mensagem
  );

  if (erro) {

    console.error(
      "Erro completo:",
      erro
    );

  }

  console.groupEnd();


  /*
    Guardamos o diagnóstico para a sessão.
    Assim conseguimos consultar posteriormente
    sem alterar novamente o código.
  */

  try {

    sessionStorage.setItem(
      "registrosEnvios_ssoErro",
      JSON.stringify({
        data:
          new Date().toISOString(),
        codigo:
          codigo,
        mensagem:
          mensagem
      })
    );

  } catch (e) {

    console.warn(
      "Não foi possível salvar diagnóstico:",
      e
    );

  }

}


/* =========================================================
   MOSTRAR DIAGNÓSTICO SSO
   ========================================================= */

function obterDiagnosticoSSO() {

  try {

    const dados =
      sessionStorage.getItem(
        "registrosEnvios_ssoErro"
      );

    if (!dados) {
      return null;
    }

    return JSON.parse(dados);

  } catch (erro) {

    return null;

  }

}


/* =========================================================
   FALLBACK REALIZADO POR
   ========================================================= */

function ativarFallbackRealizadoPor() {

  const campo =
    document.getElementById(
      "realizadoPor"
    );

  if (!campo) {
    return;
  }

  campo.readOnly = false;

  campo.disabled = false;

  campo.style.backgroundColor =
    "#ffffff";

  campo.placeholder =
    "Selecione ou digite o responsável";

  let lista =
    document.getElementById(
      "realizadoPorLista"
    );

  if (!lista) {

    lista =
      document.createElement(
        "datalist"
      );

    lista.id =
      "realizadoPorLista";

    document.body.appendChild(
      lista
    );

  }

  preencherDatalist(
    "realizadoPorLista",
    config.realizadoPor
  );

  campo.setAttribute(
    "list",
    "realizadoPorLista"
  );

  campo.title =
    "SSO não conseguiu identificar o usuário. Selecione um nome da Config ou digite um novo.";

  const diagnostico =
    obterDiagnosticoSSO();

  if (diagnostico) {

    mostrarStatus(
      "Não foi possível identificar automaticamente. Código SSO: " +
      diagnostico.codigo +
      ". Selecione o responsável manualmente."
    );

  } else {

    mostrarStatus(
      "Não foi possível identificar automaticamente. Selecione o responsável manualmente."
    );

  }

}


/* =========================================================
   EVENTOS
   ========================================================= */

function configurarEventos() {

  const formulario =
    document.getElementById(
      "registroForm"
    );

  if (formulario) {

    formulario.addEventListener(
      "submit",
      async function(evento) {

        evento.preventDefault();

        await salvarRegistro();

      }
    );

  }


  const botaoSalvar =
    document.getElementById(
      "salvar"
    );

  if (botaoSalvar) {

    botaoSalvar.addEventListener(
      "click",
      salvarRegistro
    );

  }


  const botaoCancelar =
    document.getElementById(
      "cancelar"
    );

  if (botaoCancelar) {

    botaoCancelar.addEventListener(
      "click",
      limparFormulario
    );

  }


  const campoMensagem =
    document.getElementById(
      "mensagem"
    );

  if (campoMensagem) {

    campoMensagem.addEventListener(
      "change",
      atualizarTextoMensagem
    );

  }

}


/* =========================================================
   MENSAGEM SELECIONADA
   ========================================================= */

function atualizarTextoMensagem() {

  const select =
    document.getElementById(
      "mensagem"
    );

  const nome =
    select
      ? select.value
      : "";

  if (!nome) {
    return;
  }

  const indice =
    config.nomesMensagem.indexOf(
      nome
    );

  if (indice >= 0) {

    console.log(
      "Mensagem selecionada:",
      config.textosMensagem[
        indice
      ]
    );

  }

}


/* =========================================================
   SALVAR REGISTRO
   ========================================================= */

async function salvarRegistro() {

  const botaoSalvar =
    document.getElementById(
      "salvar"
    );

  try {

    if (botaoSalvar) {
      botaoSalvar.disabled = true;
    }

    limparStatus();


    const realizadoPor =
      obterValor(
        "realizadoPor"
      );

    const empresa =
      obterValor(
        "empresa"
      );

    const qtde =
      obterValor(
        "qtde"
      );

    const emailResposta =
      obterValor(
        "emailResposta"
      );

    const supervisor =
      obterValor(
        "supervisor"
      );

    const obs =
      obterValor(
        "obs"
      );

    const historicoExterno =
      obterValor(
        "historicoExterno"
      );

    const nomeMensagem =
      obterValor(
        "mensagem"
      );

    const assunto =
      obterValor(
        "assunto"
      );


    /* =====================================================
       VALIDAÇÕES
       ===================================================== */

    if (!realizadoPor) {

      throw new Error(
        "Informe quem realizou o envio."
      );

    }

    if (!empresa) {

      throw new Error(
        "Informe a Empresa."
      );

    }

    if (!qtde) {

      throw new Error(
        "Informe a quantidade."
      );

    }

    const quantidadeNumero =
      Number(qtde);

    if (
      !Number.isInteger(
        quantidadeNumero
      ) ||
      quantidadeNumero <= 0
    ) {

      throw new Error(
        "A quantidade deve ser um número inteiro maior que zero."
      );

    }

    if (!emailResposta) {

      throw new Error(
        "Informe o Email Resposta."
      );

    }

    if (!supervisor) {

      throw new Error(
        "Informe o Supervisor."
      );

    }

    if (!historicoExterno) {

      throw new Error(
        "Informe o Histórico Externo."
      );

    }

    if (!nomeMensagem) {

      throw new Error(
        "Selecione a Mensagem."
      );

    }

    if (!assunto) {

      throw new Error(
        "Informe o Assunto."
      );

    }


    /* =====================================================
       TEXTO COMPLETO DA MENSAGEM
       ===================================================== */

    const indiceMensagem =
      config.nomesMensagem.indexOf(
        nomeMensagem
      );

    if (indiceMensagem === -1) {

      throw new Error(
        "A mensagem selecionada não foi encontrada na Config."
      );

    }

    const textoMensagem =
      config.textosMensagem[
        indiceMensagem
      ] || "";


    /* =====================================================
       DATA
       ===================================================== */

    const dataAtual =
      new Date().toLocaleDateString(
        "pt-BR"
      );


    /* =====================================================
       GRAVAR NA ABA EMAIL
       ===================================================== */

    await Excel.run(
      async (context) => {

        const planilha =
          context.workbook.worksheets.getItem(
            "Email"
          );

        const usado =
          planilha.getUsedRange();

        usado.load([
          "rowIndex",
          "rowCount"
        ]);

        await context.sync();

        const proximaLinha =
          usado.rowIndex +
          usado.rowCount;


        /*
          Ordem:

          1 Data
          2 Realizado por
          3 Empresa
          4 Qtde
          5 Email Resposta
          6 Supervisor
          7 Obs
          8 Historico Externo
          9 Mensagem
          10 Assunto
        */

        const valores = [[

          dataAtual,
          realizadoPor,
          empresa,
          quantidadeNumero,
          emailResposta,
          supervisor,
          obs,
          historicoExterno,
          textoMensagem,
          assunto

        ]];

        const intervalo =
          planilha.getRangeByIndexes(
            proximaLinha,
            0,
            1,
            10
          );

        intervalo.values =
          valores;

        await context.sync();

      }
    );


    /* =====================================================
       ATUALIZAR CONFIG
       ===================================================== */

    await atualizarConfig(
      empresa,
      emailResposta,
      supervisor,
      realizadoPor
    );


    limparFormulario();


    mostrarStatus(
      "Registro salvo com sucesso!"
    );


  } catch (erro) {

    console.error(
      "Erro ao salvar:",
      erro
    );

    mostrarStatus(
      obterMensagemErro(erro),
      true
    );


  } finally {

    if (botaoSalvar) {
      botaoSalvar.disabled = false;
    }

  }

}


/* =========================================================
   ATUALIZAR CONFIG
   ========================================================= */

async function atualizarConfig(
  empresa,
  emailResposta,
  supervisor,
  realizadoPor
) {

  await Excel.run(
    async (context) => {

      const planilha =
        context.workbook.worksheets.getItem(
          "Config"
        );

      const usado =
        planilha.getUsedRange();

      usado.load([
        "values",
        "rowIndex",
        "rowCount",
        "columnCount"
      ]);

      await context.sync();

      const valores =
        usado.values || [];

      if (
        valores.length === 0
      ) {
        return;
      }

      const linhas =
        valores.slice(1);


      await adicionarNaConfig(
        context,
        planilha,
        usado,
        linhas,
        "EMPRESAS",
        empresa
      );

      await adicionarNaConfig(
        context,
        planilha,
        usado,
        linhas,
        "EMAIL RESPOSTA",
        emailResposta
      );

      await adicionarNaConfig(
        context,
        planilha,
        usado,
        linhas,
        "SUPERVISORES",
        supervisor
      );

      await adicionarNaConfig(
        context,
        planilha,
        usado,
        linhas,
        "REALIZADO POR",
        realizadoPor
      );


      await context.sync();

    }
  );


  /*
    Recarrega as listas.
  */

  await carregarConfig();

  preencherCampos();

}


/* =========================================================
   ADICIONAR NA CONFIG
   ========================================================= */

async function adicionarNaConfig(
  context,
  planilha,
  usado,
  linhas,
  nomeColuna,
  valor
) {

  if (!valor) {
    return;
  }

  const indice =
    configHeaders[
      normalizar(nomeColuna)
    ];

  if (indice === undefined) {
    return;
  }


  const existe =
    linhas.some(
      linha => {

        const existente =
          String(
            linha[indice] || ""
          ).trim();

        return (
          normalizar(existente) ===
          normalizar(valor)
        );

      }
    );


  if (existe) {
    return;
  }


  const novaLinha =
    usado.rowIndex +
    usado.rowCount;

  const celula =
    planilha.getCell(
      novaLinha,
      indice
    );

  celula.values =
    [[valor]];

  usado.rowCount += 1;

}


/* =========================================================
   LIMPAR FORMULÁRIO
   ========================================================= */

function limparFormulario() {

  const campos = [

    "empresa",
    "qtde",
    "emailResposta",
    "supervisor",
    "obs",
    "historicoExterno",
    "assunto"

  ];


  campos.forEach(
    id => {

      const elemento =
        document.getElementById(
          id
        );

      if (elemento) {
        elemento.value = "";
      }

    }
  );


  const mensagem =
    document.getElementById(
      "mensagem"
    );

  if (mensagem) {
    mensagem.value = "";
  }


  /*
    Se o usuário foi identificado
    automaticamente, mantém o nome.

    Se estiver no fallback manual,
    limpa o campo.
  */

  const realizadoPor =
    document.getElementById(
      "realizadoPor"
    );

  if (realizadoPor) {

    if (
      !realizadoPor.readOnly
    ) {

      realizadoPor.value = "";

    }

  }

}


/* =========================================================
   OBTER VALOR
   ========================================================= */

function obterValor(id) {

  const elemento =
    document.getElementById(
      id
    );

  if (!elemento) {
    return "";
  }

  return String(
    elemento.value || ""
  ).trim();

}


/* =========================================================
   NORMALIZAR
   ========================================================= */

function normalizar(valor) {

  return String(
    valor || ""
  )
    .trim()
    .toLowerCase();

}


/* =========================================================
   STATUS
   ========================================================= */

function mostrarStatus(
  mensagem,
  erro = false
) {

  const status =
    document.getElementById(
      "status"
    );

  if (!status) {
    return;
  }

  status.textContent =
    mensagem;

  status.className =
    erro
      ? "status erro"
      : "status sucesso";

}


function limparStatus() {

  const status =
    document.getElementById(
      "status"
    );

  if (!status) {
    return;
  }

  status.textContent = "";

  status.className =
    "status";

}


/* =========================================================
   CÓDIGO DO ERRO
   ========================================================= */

function obterCodigoErro(
  erro
) {

  if (!erro) {
    return "SEM_CODIGO";
  }

  if (erro.code !== undefined) {
    return String(erro.code);
  }

  if (
    erro.errorCode !== undefined
  ) {
    return String(
      erro.errorCode
    );
  }

  if (
    erro.name &&
    erro.name !== "Error"
  ) {
    return erro.name;
  }

  return "SEM_CODIGO";

}


/* =========================================================
   MENSAGEM DO ERRO
   ========================================================= */

function obterMensagemErro(
  erro
) {

  if (!erro) {
    return "Ocorreu um erro.";
  }

  if (
    typeof erro === "string"
  ) {
    return erro;
  }

  if (erro.message) {
    return erro.message;
  }

  return "Ocorreu um erro inesperado.";

}
