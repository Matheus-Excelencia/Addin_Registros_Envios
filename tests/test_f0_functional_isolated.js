const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const scripts = {
  email: "app/scripts/email.js",
  sms: "app/scripts/sms.js",
  history: "app/scripts/historico.js",
  config: "app/scripts/config.js",
};

function loadHelpers(name, extra = "") {
  const source = fs.readFileSync(path.join(root, scripts[name]), "utf8");
  const sandbox = {
    Office: { onReady() {} },
    window: { addEventListener() {} },
    document: { getElementById() { return null; } },
    console,
    Date,
    Math,
    Set,
  };
  vm.createContext(sandbox);
  vm.runInContext(source + "\n" + extra, sandbox, { filename: scripts[name] });
  return sandbox;
}

test("header normalization preserves accents and normalizes whitespace/case", () => {
  for (const name of ["email", "sms", "history", "config"]) {
    const sandbox = loadHelpers(name, "globalThis.__test = { normalizar: typeof normalizar === 'function' ? normalizar : n };");
    const normalize = sandbox.__test.normalizar;
    assert.equal(normalize("  Histórico   Externo "), "HISTÓRICO EXTERNO", name);
    assert.equal(normalize("  Obs.  "), "OBS.", name);
  }
});

test("duplicate normalized headers fail instead of selecting a column", () => {
  for (const [name, fnName] of [["email", "encontrarCabecalho"], ["sms", "encontrarCabecalho"], ["history", "encontrarCabecalho"], ["config", "head"]]) {
    const sandbox = loadHelpers(name, `globalThis.__test = { parser: ${fnName} };`);
    const parse = sandbox.__test.parser;
    const row = name === "history"
      ? [["DATA", " data ", "EMPRESA"]]
      : name === "config"
        ? [["EMPRESAS", " EMPRESAS ", "SUPERVISORES"]]
        : [["EMPRESAS", " EMPRESAS ", "SUPERVISORES"]];
    assert.throws(() => parse(row), /Cabeçalho duplicado ou ambíguo/, name);
  }
});

test("aliases resolving to different columns are rejected", () => {
  for (const [name, fnName] of [["email", "obterIndice"], ["sms", "obterIndice"], ["history", "indice"], ["config", "idx"]]) {
    const sandbox = loadHelpers(name, `globalThis.__test = { resolve: ${fnName} };`);
    const resolve = sandbox.__test.resolve;
    assert.throws(() => resolve({ "EMPRESA": 0, "EMPRESAS": 1 }, ["EMPRESA", "EMPRESAS"]), /Aliases ambíguos para o mesmo campo/, name);
  }
});

test("Email and SMS block missing ID REGISTRO without changing workbook", async () => {
  const headers = {
    email: ["DATA", "REALIZADO POR", "EMPRESA", "QTDE", "EMAIL RESPOSTA", "SUPERVISOR", "OBS", "HISTORICO EXTERNO", "MENSAGEM", "ASSUNTO"],
    sms: ["DATA", "REALIZADO POR", "EMPRESA", "QTDE", "SUPERVISOR", "OBS.", "HISTORICO EXTERNO", "MENSAGEM"],
  };
  for (const [name, functionName, sheetName] of [["email", "adicionarRegistroEmail", "Email"], ["sms", "adicionarRegistroSMS", "SMS"]]) {
    const sandbox = loadHelpers(name, `globalThis.__test = { add: ${functionName} };`);
    const writes = [];
    const used = {
      values: [headers[name]],
      isNullObject: false,
      rowIndex: 0,
      columnIndex: 0,
      load() {},
    };
    const context = {
      workbook: { worksheets: { getItem(requested) {
        assert.equal(requested, sheetName);
        return {
          getUsedRangeOrNullObject() { return used; },
          getRangeByIndexes() { writes.push("getRangeByIndexes"); throw new Error("unexpected write"); },
          getRange() { writes.push("getRange"); throw new Error("unexpected write"); },
        };
      } } },
      sync: async () => {},
    };
    await assert.rejects(sandbox.__test.add(context, {}), /ID REGISTRO está ausente/, name);
    assert.deepEqual(writes, [], name);
  }
});

test("uncertain-write paths do not loop/retry automatically", () => {
  for (const name of ["email", "sms"]) {
    const source = fs.readFileSync(path.join(root, scripts[name]), "utf8");
    assert.match(source, /for\s*\(let tentativa = 1; tentativa <= 1; tentativa\+\+\)/, name);
    assert.doesNotMatch(source, /tentativa <= 3/, name);
    assert.match(source, /Não foi possível confirmar a gravação/, name);
  }
});
