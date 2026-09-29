/* global Office, Excel, document, window, sessionStorage, console, atob */

"use strict";

/* =========================================================
   CONFIGURAÇÕES
========================================================= */

const ABA_EMAIL = "Email";
const ABA_CONFIG = "Config";

const COLUNAS_EMAIL = [
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
];

const CONFIG_COLUNAS = {
  empresas: 0,
  emails: 1,
  supervisores: 2,
  realizadosPor: 3,
  nomesMensagem: 4,
  textosMensagem: 5
};

/* =========================================================
   INICIALIZAÇÃO
========================================================= */

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    mostrarStatus("Este suplemento foi desenvolvido para o Excel.");
    return;
  }

  console.log("Excel pronto.");

  configurarEventos();

  try {
    await carregarConfiguracao();
    await identificarUsuario();
  } catch (erro) {
    console.error("Erro na inicialização:", erro);
    mostrarStatus("Erro ao inicializar o suplemento.");
  }
});


/* =========================================================
   EVENTOS
========================================================= */

function configurarEventos() {

  const btnSalvar = document.getElementById("btnSalvar");
  const btnCancelar = document.getElementById("btnCancelar");

  if (btnSalvar) {
    btnSalvar.addEventListener("click", salvarRegistro);
  }

  if (btnCancelar) {
    btnCancelar.addEventListener("click", limparFormulario);
  }

  const mensagemSelect = document.getElementById("mensagem");

  if (mensagemSelect) {
    mensagemSelect.addEventListener("change", preencherTextoMensagem);
  }

  const empresa = document.getElementById("empresa");
  const emailResposta = document.getElementById("emailResposta");
  const supervisor = document.getElementById("supervisor");

  if (empresa) {
    empresa.addEventListener("change", permitirNovoValor);
  }

  if (emailResposta) {
    emailResposta.addEventListener("change", permitirNovoValor);
  }

  if (supervisor) {
    supervisor.addEventListener("change", permitirNovoValor);
  }
}


/* =========================================================
   DIAGNÓSTICO SSO
========================================================= */

function diagnosticarSSO() {

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

  return {
    identityApi,
    getAccessTokenDisponivel,
    versaoOffice,
    hostOffice
  };
}


/* =========================================================
   IDENTIFICAÇÃO DO USUÁRIO
========================================================= */

async function identificarUsuario() {

  try {

    const diagnostico = diagnosticarSSO();

    if (!diagnostico.identityApi) {
      throw criarErroSSO(
        "IDENTITY_API_NAO_SUPORTADA",
        "O Excel não informou suporte à IdentityAPI 1.3."
      );
    }

    if (!diagnostico.getAccessTokenDisponivel) {
      throw criarErroSSO(
        "GET_ACCESS_TOKEN_NAO_DISPONIVEL",
        "Office.auth.getAccessToken não está disponível."
      );
    }

    mostrarStatus("Identificando usuário Microsoft...");

    const token = await obterTokenSSO();

    if (!token) {
      throw criarErroSSO(
        "TOKEN_VAZIO",
        "O Excel não retornou um token SSO."
      );
    }

    console.log("Token SSO recebido.");

    const payload = decodificarToken(token);

    console.log("Payload SSO:", payload);

    const nome =
      payload.name ||
      payload.preferred_username ||
      payload.email ||
      payload.upn ||
      "";

    if (!nome) {
      throw criarErroSSO(
        "USUARIO_NAO_IDENTIFICADO",
        "O token foi recebido, mas não contém nome ou usuário."
      );
    }

    preencherRealizadoPor(nome);

    mostrarStatus(
      "Usuário identificado automaticamente: " + nome
    );

    ocultarFallbackRealizadoPor();

  } catch (erro) {

    registrarErroSSO(erro);

    ativarFallbackRealizadoPor();

    mostrarStatus(
      "Não foi possível identificar automaticamente. " +
      "Código SSO: " +
      obterCodigoErro(erro) +
      ". Selecione o responsável manualmente."
    );
  }
}


/* =========================================================
   OBTER TOKEN SSO
========================================================= */

async function obterTokenSSO() {

  try {

    const token = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true
    });

    return token;

  } catch (erro) {

    console.error("Erro ao obter token SSO:", erro);

    throw criarErroSSO(
      erro?.code ||
      erro?.errorCode ||
      "GET_ACCESS_TOKEN_ERRO",
      erro?.message ||
      "Não foi possível obter o token SSO."
    );
  }
}


/* =========================================================
   DECODIFICAR JWT
========================================================= */

function decodificarToken(token) {

  if (!token || typeof token !== "string") {
    throw new Error("Token SSO inválido ou vazio.");
  }

  const partes = token.split(".");

  if (partes.length !== 3) {
    throw new Error(
      "Token SSO não possui formato JWT válido."
    );
  }

  const base64Url = partes[1];

  const base64 = base64Url
    .replace(/-/g, "+")
    .replace(/_/g, "/");

  const padding =
    "=".repeat(
      (4 - (base64.length % 4)) % 4
    );

  const binary = atob(base64 + padding);

  const bytes = Uint8Array.from(
    binary,
    char => char.charCodeAt(0)
  );

  const texto =
    new TextDecoder("utf-8").decode(bytes);

  return JSON.parse(texto);
}


/* =========================================================
   ERROS SSO
========================================================= */

function criarErroSSO(codigo, mensagem) {

  const erro = new Error(mensagem);

  erro.code = codigo;
  erro.errorCode = codigo;

  return erro;
}


function obterCodigoErro(erro) {

  if (!erro) {
    return "DESCONHECIDO";
  }

  return (
    erro.code ||
    erro.errorCode ||
    erro.name ||
    "DESCONHECIDO"
  );
}


function registrarErroSSO(erro) {

  console.error(
    "Diagnóstico SSO"
  );

  console.error(
    "Código:",
    obterCodigoErro(erro)
  );

  console.error(
    "Mensagem:",
    erro?.message
  );

  console.error(
    "Erro completo:",
    erro
  );

  try {

    sessionStorage.setItem(
      "ultimoErroSSO",
      JSON.stringify({
        codigo: obterCodigoErro(erro),
        mensagem: erro?.message || "",
        data: new Date().toISOString()
      })
    );

  } catch (e) {
    console.warn(
      "Não foi possível salvar erro SSO.",
      e
    );
  }
}


/* =========================================================
   REALIZADO POR
========================================================= */

function preencherRealizadoPor(nome) {

  const campo =
    document.getElementById("realizadoPor");

  if (!campo) {
    return;
  }

  if (
    campo.tagName === "SELECT"
  ) {

    let encontrou = false;

    for (const option of campo.options) {

      if (
        option.value.trim().toLowerCase() ===
        nome.trim().toLowerCase()
      ) {

        campo.value = option.value;
        encontrou = true;
        break;
      }
    }

    if (!encontrou) {

      const option =
        document.createElement("option");

      option.value = nome;
      option.textContent = nome;

      campo.appendChild(option);

      campo.value = nome;
    }

  } else {

    campo.value = nome;
  }
}


function ativarFallbackRealizadoPor() {

  const campo =
    document.getElementById("realizadoPor");

  if (!campo) {
    return;
  }

  campo.disabled = false;

  campo.removeAttribute("readonly");
}


function ocultarFallbackRealizadoPor() {

  const campo =
    document.getElementById("realizadoPor");

  if (!campo) {
    return;
  }

  campo.disabled = false;
}


/* =========================================================
   CONFIG
========================================================= */

async function carregarConfiguracao() {

  await Excel.run(async (context) => {

    const sheet =
      context.workbook.worksheets.getItem(ABA_CONFIG);

    const usedRange =
      sheet.getUsedRangeOrNullObject();

    usedRange.load([
      "values",
      "rowCount",
      "columnCount"
    ]);

    await context.sync();

    if (usedRange.isNullObject) {
      return;
    }

    const valores = usedRange.values;

    if (!valores || valores.length < 1) {
      return;
    }

    const dados = valores.slice(1);

    preencherSelect(
      "empresa",
      dados.map(linha => linha[CONFIG_COLUNAS.empresas])
    );

    preencherSelect(
      "emailResposta",
      dados.map(linha => linha[CONFIG_COLUNAS.emails])
    );

    preencherSelect(
      "supervisor",
      dados.map(linha => linha[CONFIG_COLUNAS.supervisores])
    );

    preencherSelect(
      "realizadoPor",
      dados.map(linha => linha[CONFIG_COLUNAS.realizadosPor])
    );

    preencherMensagens(dados);
  });
}


/* =========================================================
   PREENCHER SELECTS
========================================================= */

function preencherSelect(id, valores) {

  const select =
    document.getElementById(id);

  if (!select) {
    return;
  }

  const atual =
    select.value;

  select.innerHTML =
    '<option value="">Selecione...</option>';

  const unicos =
    [...new Set(
      valores
        .map(v => String(v ?? "").trim())
        .filter(v => v !== "")
    )];

  for (const valor of unicos) {

    const option =
      document.createElement("option");

    option.value = valor;
    option.textContent = valor;

    select.appendChild(option);
  }

  if (atual) {
    select.value = atual;
  }
}


/* =========================================================
   MENSAGENS
========================================================= */

let mensagensConfig = [];


function preencherMensagens(dados) {

  mensagensConfig = [];

  const select =
    document.getElementById("mensagem");

  if (!select) {
    return;
  }

  select.innerHTML =
    '<option value="">Selecione...</option>';

  for (const linha of dados) {

    const nome =
      String(
        linha[CONFIG_COLUNAS.nomesMensagem] ?? ""
      ).trim();

    const texto =
      String(
        linha[CONFIG_COLUNAS.textosMensagem] ?? ""
      ).trim();

    if (!nome) {
      continue;
    }

    mensagensConfig.push({
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


function preencherTextoMensagem() {

  const select =
    document.getElementById("mensagem");

  const campoTexto =
    document.getElementById("textoMensagem");

  if (!select || !campoTexto) {
    return;
  }

  const selecionada =
    mensagensConfig.find(
      mensagem =>
        mensagem.nome === select.value
    );

  if (selecionada) {

    campoTexto.value =
      selecionada.texto;

  } else {

    campoTexto.value = "";
  }
}


/* =========================================================
   NOVOS VALORES
========================================================= */

function permitirNovoValor(evento) {

  const select =
    evento.target;

  if (!select.value) {
    return;
  }

  const existe =
    [...select.options].some(
      option =>
        option.value.trim().toLowerCase() ===
        select.value.trim().toLowerCase()
    );

  if (!existe) {

    const option =
      document.createElement("option");

    option.value =
      select.value;

    option.textContent =
      select.value;

    select.appendChild(option);
  }
}


/* =========================================================
   SALVAR REGISTRO
========================================================= */

async function salvarRegistro() {

  try {

    const dados =
      coletarFormulario();

    const validacao =
      validarFormulario(dados);

    if (!validacao.valido) {

      mostrarStatus(
        validacao.mensagem
      );

      alert(
        validacao.mensagem
      );

      return;
    }

    mostrarStatus(
      "Salvando registro..."
    );

    await adicionarRegistroEmail(dados);

    await atualizarConfig(dados);

    mostrarStatus(
      "Registro salvo com sucesso!"
    );

    alert(
      "Registro salvo com sucesso."
    );

    limparFormulario();

    await carregarConfiguracao();

  } catch (erro) {

    console.error(
      "Erro ao salvar:",
      erro
    );

    mostrarStatus(
      "Erro ao salvar o registro."
    );

    alert(
      "Erro ao salvar o registro:\n\n" +
      (erro?.message || erro)
    );
  }
}


/* =========================================================
   COLETAR FORMULÁRIO
========================================================= */

function coletarFormulario() {

  const valor =
    id => {

      const elemento =
        document.getElementById(id);

      return elemento
        ? elemento.value.trim()
        : "";
    };

  const mensagemNome =
    valor("mensagem");

  const mensagemEncontrada =
    mensagensConfig.find(
      mensagem =>
        mensagem.nome === mensagemNome
    );

  return {

    data: new Date(),

    realizadoPor:
      valor("realizadoPor"),

    empresa:
      valor("empresa"),

    qtde:
      valor("qtde"),

    emailResposta:
      valor("emailResposta"),

    supervisor:
      valor("supervisor"),

    obs:
      valor("obs"),

    historicoExterno:
      valor("historicoExterno"),

    mensagem:
      mensagemEncontrada
        ? mensagemEncontrada.texto
        : valor("textoMensagem"),

    assunto:
      valor("assunto")
  };
}


/* =========================================================
   VALIDAR FORMULÁRIO
========================================================= */

function validarFormulario(dados) {

  if (!dados.empresa) {

    return {
      valido: false,
      mensagem: "Informe a Empresa."
    };
  }

  if (!dados.qtde) {

    return {
      valido: false,
      mensagem: "Informe a Qtde."
    };
  }

  const quantidade =
    Number(dados.qtde);

  if (
    !Number.isInteger(quantidade) ||
    quantidade <= 0
  ) {

    return {
      valido: false,
      mensagem:
        "A Qtde deve ser um número inteiro maior que zero."
    };
  }

  if (!dados.emailResposta) {

    return {
      valido: false,
      mensagem:
        "Informe o Email Resposta."
    };
  }

  if (!dados.supervisor) {

    return {
      valido: false,
      mensagem:
        "Informe o Supervisor."
    };
  }

  if (!dados.historicoExterno) {

    return {
      valido: false,
      mensagem:
        "Informe o Histórico Externo."
    };
  }

  if (!dados.mensagem) {

    return {
      valido: false,
      mensagem:
        "Selecione uma Mensagem."
    };
  }

  if (!dados.assunto) {

    return {
      valido: false,
      mensagem:
        "Informe o Assunto."
    };
  }

  if (!dados.realizadoPor) {

    return {
      valido: false,
      mensagem:
        "Informe o responsável."
    };
  }

  return {
    valido: true,
    mensagem: ""
  };
}


/* =========================================================
   ADICIONAR NA ABA EMAIL
========================================================= */

async function adicionarRegistroEmail(dados) {

  await Excel.run(async (context) => {

    const sheet =
      context.workbook.worksheets.getItem(
        ABA_EMAIL
      );

    const usedRange =
      sheet.getUsedRangeOrNullObject();

    usedRange.load([
      "rowCount",
      "columnCount"
    ]);

    await context.sync();

    let proximaLinha = 0;

    if (usedRange.isNullObject) {

      proximaLinha = 0;

    } else {

      proximaLinha =
        usedRange.rowCount;
    }

    if (proximaLinha === 0) {

      const cabecalho =
        [COLUNAS_EMAIL];

      sheet
        .getRangeByIndexes(
          0,
          0,
          1,
          COLUNAS_EMAIL.length
        )
        .values =
        cabecalho;

      proximaLinha = 1;
    }

    const linha = [[

      formatarDataExcel(dados.data),

      dados.realizadoPor,

      dados.empresa,

      Number(dados.qtde),

      dados.emailResposta,

      dados.supervisor,

      dados.obs,

      dados.historicoExterno,

      dados.mensagem,

      dados.assunto
    ]];

    sheet
      .getRangeByIndexes(
        proximaLinha,
        0,
        1,
        COLUNAS_EMAIL.length
      )
      .values = linha;

    await context.sync();
  });
}


/* =========================================================
   ATUALIZAR CONFIG
========================================================= */

async function atualizarConfig(dados) {

  await Excel.run(async (context) => {

    const sheet =
      context.workbook.worksheets.getItem(
        ABA_CONFIG
      );

    const usedRange =
      sheet.getUsedRangeOrNullObject();

    usedRange.load([
      "values",
      "rowCount",
      "columnCount"
    ]);

    await context.sync();

    let valores = [];

    if (!usedRange.isNullObject) {
      valores = usedRange.values;
    }

    if (valores.length === 0) {

      sheet
        .getRange("A1:F1")
        .values = [[
          "EMPRESAS",
          "EMAIL RESPOSTA",
          "SUPERVISORES",
          "REALIZADO POR",
          "NOME MENSAGEM",
          "TEXTO MENSAGEM"
        ]];

      valores = [[
        "EMPRESAS",
        "EMAIL RESPOSTA",
        "SUPERVISORES",
        "REALIZADO POR",
        "NOME MENSAGEM",
        "TEXTO MENSAGEM"
      ]];
    }

    const novosValores = [

      {
        coluna: CONFIG_COLUNAS.empresas,
        valor: dados.empresa
      },

      {
        coluna: CONFIG_COLUNAS.emails,
        valor: dados.emailResposta
      },

      {
        coluna: CONFIG_COLUNAS.supervisores,
        valor: dados.supervisor
      },

      {
        coluna: CONFIG_COLUNAS.realizadosPor,
        valor: dados.realizadoPor
      }
    ];

    for (const item of novosValores) {

      if (!item.valor) {
        continue;
      }

      const existe =
        valores
          .slice(1)
          .some(
            linha =>
              String(
                linha[item.coluna] ?? ""
              )
                .trim()
                .toLowerCase() ===
              item.valor
                .trim()
                .toLowerCase()
          );

      if (!existe) {

        const proximaLinha =
          valores.length;

        sheet
          .getCell(
            proximaLinha,
            item.coluna
          )
          .values = [[
            item.valor
          ]];

        valores.push(
          new Array(6).fill("")
        );
      }
    }

    await context.sync();
  });
}


/* =========================================================
   LIMPAR FORMULÁRIO
========================================================= */

function limparFormulario() {

  const ids = [
    "empresa",
    "qtde",
    "emailResposta",
    "supervisor",
    "obs",
    "historicoExterno",
    "mensagem",
    "textoMensagem",
    "assunto"
  ];

  for (const id of ids) {

    const elemento =
      document.getElementById(id);

    if (!elemento) {
      continue;
    }

    if (
      elemento.tagName === "SELECT"
    ) {

      elemento.selectedIndex = 0;

    } else {

      elemento.value = "";
    }
  }

  mostrarStatus(
    "Novo registro."
  );
}


/* =========================================================
   DATA
========================================================= */

function formatarDataExcel(data) {

  const ano =
    data.getFullYear();

  const mes =
    String(
      data.getMonth() + 1
    ).padStart(2, "0");

  const dia =
    String(
      data.getDate()
    ).padStart(2, "0");

  return `${ano}-${mes}-${dia}`;
}


/* =========================================================
   STATUS
========================================================= */

function mostrarStatus(mensagem) {

  console.log(
    "STATUS:",
    mensagem
  );

  const elementos = [
    "status",
    "mensagemStatus"
  ];

  for (const id of elementos) {

    const elemento =
      document.getElementById(id);

    if (elemento) {

      elemento.textContent =
        mensagem;
    }
  }
}
