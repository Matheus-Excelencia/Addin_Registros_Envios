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


/*
 * SIMULAÇÃO DO PROTOCOLO — não executa o código real de persistência Office.js.
 * Serve para definir o comportamento esperado de falha incerta/reconciliação.
 */
function createUncertainWriteSimulator({ failBeforePersist = false, failAfterPersist = false } = {}) {
  const rows = [];
  let writeAttempts = 0;
  return {
    rows,
    get writeAttempts() { return writeAttempts; },
    async write(record) {
      writeAttempts += 1;
      if (failBeforePersist) throw new Error("simulated failure before persistence");
      rows.push({ ...record });
      if (failAfterPersist) throw new Error("simulated confirmation failure after persistence");
      return { confirmed: true, id: record.id };
    },
    findById(id) {
      return rows.filter(row => String(row.id) === String(id));
    },
  };
}

async function writeWithReconciliation(simulator, record) {
  try {
    const result = await simulator.write(record);
    return { state: "confirmed", id: result.id };
  } catch (error) {
    const matches = simulator.findById(record.id);
    if (matches.length === 1) return { state: "reconciled-existing", id: record.id };
    if (matches.length > 1) return { state: "duplicate-conflict", id: record.id };
    return { state: "unresolved-no-match", id: record.id, reason: error.message };
  }
}

test("SIMULATION ONLY: failure before persistence leaves no record and no blind retry", async () => {
  const sim = createUncertainWriteSimulator({ failBeforePersist: true });
  const result = await writeWithReconciliation(sim, { id: "F0-SIM-BEFORE-001", value: "x" });
  assert.equal(result.state, "unresolved-no-match");
  assert.equal(sim.rows.length, 0);
  assert.equal(sim.writeAttempts, 1);
});

test("SIMULATION ONLY: confirmation failure after persistence reconciles by ID without retry", async () => {
  const sim = createUncertainWriteSimulator({ failAfterPersist: true });
  const result = await writeWithReconciliation(sim, { id: "F0-SIM-AFTER-001", value: "x" });
  assert.equal(result.state, "reconciled-existing");
  assert.equal(sim.findById("F0-SIM-AFTER-001").length, 1);
  assert.equal(sim.writeAttempts, 1);
});

test("SIMULATION ONLY: reconciliation detects duplicate IDs instead of writing again", async () => {
  const sim = createUncertainWriteSimulator({ failAfterPersist: true });
  sim.rows.push({ id: "F0-SIM-DUP-001", value: "pre-existing" });
  const result = await writeWithReconciliation(sim, { id: "F0-SIM-DUP-001", value: "x" });
  assert.equal(result.state, "duplicate-conflict");
  assert.equal(sim.writeAttempts, 1);
});


/*
 * TESTE DO CÓDIGO REAL com mock da superfície Office.js.
 * Não conecta ao Excel real; executa adicionarRegistroEmail/SMS carregadas do arquivo.
 */
function createOfficeWriteMock(headers, { failMode = "none" } = {}) {
  const rows = [headers.slice()];
  let syncCount = 0;
  let writeAttempts = 0;
  let pendingWrite = null;
  const usedRange = () => ({
    get values() { return rows.map(row => row.slice()); },
    isNullObject: false,
    rowIndex: 0,
    columnIndex: 0,
    load() {},
  });
  const sheet = {
    getUsedRangeOrNullObject() { return usedRange(); },
    getRangeByIndexes(rowIndex, columnIndex, rowCount, columnCount) {
      const range = {
        _values: null,
        set values(value) {
          writeAttempts += 1;
          pendingWrite = { rowIndex, columnIndex, values: value.map(row => row.slice()) };
          range._values = value;
        },
        get values() { return range._values; },
        getCell(rowOffset, columnOffset) {
          return {
            format: { set wrapText(_value) {} },
            set numberFormat(_value) {},
            load() {},
            get values() {
              const row = rows[rowIndex + rowOffset] || [];
              return [[row[columnIndex + columnOffset]]];
            },
          };
        },
      };
      return range;
    },
  };
  const context = {
    workbook: { worksheets: { getItem(name) {
      assert.ok(name === "Email" || name === "SMS");
      return sheet;
    } } },
    async sync() {
      syncCount += 1;
      // First sync loads the header range; second sync refreshes the target used range.
      if (syncCount < 3 || !pendingWrite) return;
      const queued = pendingWrite;
      pendingWrite = null;
      if (failMode === "before-persist") {
        throw new Error("mock sync failed before persistence");
      }
      for (let r = 0; r < queued.values.length; r++) {
        const targetRow = queued.rowIndex + r;
        while (rows.length <= targetRow) rows.push([]);
        for (let c = 0; c < queued.values[r].length; c++) {
          rows[targetRow][queued.columnIndex + c] = queued.values[r][c];
        }
      }
      if (failMode === "after-persist") {
        throw new Error("mock confirmation failed after persistence");
      }
    },
  };
  return { context, rows, get writeAttempts() { return writeAttempts; } };
}

const completeHeaders = {
  email: ["DATA", "REALIZADO POR", "EMPRESA", "QTDE", "EMAIL RESPOSTA", "SUPERVISOR", "OBS", "HISTORICO EXTERNO", "MENSAGEM", "ASSUNTO", "ID REGISTRO"],
  sms: ["DATA", "REALIZADO POR", "EMPRESA", "QTDE", "SUPERVISOR", "OBS.", "HISTORICO EXTERNO", "MENSAGEM", "ID REGISTRO"],
};
const validRecord = {
  realizadoPor: "Teste", empresa: "Empresa Teste", qtde: 1,
  emailResposta: "teste@example.com", supervisor: "Supervisão",
  obs: "", historicoExterno: "EXT-1", textoMensagem: "Mensagem teste",
  assunto: "Assunto teste", idRegistro: "F0-MOCK-ID-001",
};

test("REAL FUNCTION + Office.js mock: sync failure before persistence leaves no row and does not retry", async () => {
  for (const [name, fnName] of [["email", "adicionarRegistroEmail"], ["sms", "adicionarRegistroSMS"]]) {
    const sandbox = loadHelpers(name, `globalThis.__test = { add: ${fnName} };`);
    const mock = createOfficeWriteMock(completeHeaders[name], { failMode: "before-persist" });
    await assert.rejects(sandbox.__test.add(mock.context, validRecord), /mock sync failed before persistence/, name);
    assert.equal(mock.rows.length, 1, name);
    assert.equal(mock.writeAttempts, 1, name);
  }
});

test("REAL FUNCTION + Office.js mock: confirmation failure after persistence leaves one row but code does not reconcile", async () => {
  for (const [name, fnName] of [["email", "adicionarRegistroEmail"], ["sms", "adicionarRegistroSMS"]]) {
    const sandbox = loadHelpers(name, `globalThis.__test = { add: ${fnName} };`);
    const mock = createOfficeWriteMock(completeHeaders[name], { failMode: "after-persist" });
    await assert.rejects(sandbox.__test.add(mock.context, validRecord), /mock confirmation failed after persistence/, name);
    assert.equal(mock.rows.length, 2, name);
    assert.equal(mock.rows[1][completeHeaders[name].indexOf("ID REGISTRO")], validRecord.idRegistro, name);
    assert.equal(mock.writeAttempts, 1, name);
  }
});

test("REAL FUNCTION + Office.js mock: successful write confirms ID exactly once", async () => {
  for (const [name, fnName] of [["email", "adicionarRegistroEmail"], ["sms", "adicionarRegistroSMS"]]) {
    const sandbox = loadHelpers(name, `globalThis.__test = { add: ${fnName} };`);
    const mock = createOfficeWriteMock(completeHeaders[name]);
    await sandbox.__test.add(mock.context, validRecord);
    assert.equal(mock.rows.length, 2, name);
    assert.equal(mock.rows[1][completeHeaders[name].indexOf("ID REGISTRO")], validRecord.idRegistro, name);
    assert.equal(mock.writeAttempts, 1, name);
  }
});
