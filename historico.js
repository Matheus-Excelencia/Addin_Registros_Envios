let todos = [], registros = [], pronto = false;

function iniciarHistorico() {
    if (pronto) return;
    pronto = true;

    const v = document.getElementById("btnVoltar");
    const p = document.getElementById("btnPesquisar");
    const l = document.getElementById("btnLimpar");
    const o = document.getElementById("ordenacao");

    if (v) v.addEventListener("click", () => window.location.href = "taskpane.html");
    if (p) p.addEventListener("click", pesquisar);
    if (l) l.addEventListener("click", limpar);
    if (o) o.addEventListener("change", render);

    carregar();
}

if (window.Office && Office.onReady) {
    Office.onReady(() => iniciarHistorico());
} else {
    window.addEventListener("DOMContentLoaded", iniciarHistorico);
}

function normalizar(v) {
    return String(v ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
}

function indice(mapa, nomes) {
    for (const nome of nomes) {
        const chave = normalizar(nome);
        if (mapa[chave] !== undefined) return mapa[chave];
    }
    return -1;
}

function encontrarCabecalho(valores, obrigatorios) {
    for (let linha = 0; linha < Math.min(valores.length, 30); linha++) {
        const mapa = {};
        (valores[linha] || []).forEach((valor, coluna) => {
            if (String(valor ?? "").trim()) mapa[normalizar(valor)] = coluna;
        });
        if (obrigatorios.every(nomes => indice(mapa, nomes) !== -1)) {
            return { linha, mapa };
        }
    }
    return null;
}

function valor(registro, mapa, nomes) {
    const i = indice(mapa, nomes);
    return i < 0 ? "" : registro[i] ?? "";
}

function escapar(v) {
    return String(v ?? "").replace(/[&<>"']/g, c => ({
        "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[c]));
}

function dataTexto(v) {
    if (typeof v === "number") {
        return new Date(Date.UTC(1899, 11, 30) + v * 86400000)
            .toLocaleString("pt-BR", { timeZone: "UTC" });
    }
    return String(v || "—");
}

function dataNumero(v) {
    if (typeof v === "number") return Date.UTC(1899, 11, 30) + v * 86400000;
    const texto = String(v || "").trim();
    const br = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (br) return new Date(+br[3], +br[2] - 1, +br[1], +(br[4] || 0), +(br[5] || 0), +(br[6] || 0)).getTime();
    const d = new Date(texto);
    return isNaN(d) ? 0 : d.getTime();
}

async function carregar() {
    try {
        todos = [];

        await Excel.run(async context => {
            const nomes = ["Email", "SMS"];
            const ranges = nomes.map(nome => {
                const sheet = context.workbook.worksheets.getItem(nome);
                const range = sheet.getUsedRangeOrNullObject(true);
                range.load(["values", "isNullObject", "rowIndex"]);
                return { nome, range };
            });

            const config = context.workbook.worksheets.getItem("Config");
            const configRange = config.getUsedRangeOrNullObject(true);
            configRange.load(["values", "isNullObject"]);
            await context.sync();

            for (const item of ranges) {
                if (item.range.isNullObject) continue;
                const valores = item.range.values || [];
                const cab = encontrarCabecalho(valores, [["EMPRESA"], ["DATA"]]);
                if (!cab) continue;

                for (let i = cab.linha + 1; i < valores.length; i++) {
                    const r = valores[i] || [];
                    const empresa = String(valor(r, cab.mapa, ["EMPRESA"])).trim();
                    if (!empresa) continue;

                    todos.push({
                        tipo: item.nome,
                        linhaPlanilha: item.range.rowIndex + i,
                        data: valor(r, cab.mapa, ["DATA"]),
                        empresa,
                        realizadoPor: String(valor(r, cab.mapa, ["REALIZADO POR"])).trim(),
                        supervisor: String(valor(r, cab.mapa, ["SUPERVISOR"])).trim(),
                        qtde: String(valor(r, cab.mapa, ["QTDE"])).trim(),
                        historico: String(valor(r, cab.mapa, ["HISTORICO EXTERNO", "HISTÓRICO EXTERNO"])).trim(),
                        mensagem: String(valor(r, cab.mapa, ["MENSAGEM"])).trim(),
                        assunto: String(valor(r, cab.mapa, ["ASSUNTO"])).trim(),
                        id: String(valor(r, cab.mapa, ["ID REGISTRO"])).trim(),
                        emailResposta: String(valor(r, cab.mapa, ["EMAIL RESPOSTA", "E-MAIL RESPOSTA"])).trim()
                    });
                }
            }

            if (!configRange.isNullObject) preencherListasConfig(configRange.values || []);
        });

        pesquisar();
        const status = document.getElementById("status");
        if (status) status.style.display = "none";
    } catch (e) {
        erro("Não foi possível carregar a consulta. " + (e?.message || "Verifique as abas Email, SMS e Config."));
    }
}

function preencherListasConfig(valores) {
    const cab = encontrarCabecalho(valores, [
        ["EMPRESAS", "EMPRESA"],
        ["SUPERVISORES", "SUPERVISOR"]
    ]);
    if (!cab) {
        preencherListasFallback();
        return;
    }

    const empresas = [], supervisores = [], realizados = [];
    const ie = indice(cab.mapa, ["EMPRESAS", "EMPRESA"]);
    const is = indice(cab.mapa, ["SUPERVISORES", "SUPERVISOR"]);
    const ir = indice(cab.mapa, ["REALIZADO POR", "REALIZADO_POR", "REALIZADOPOR"]);

    for (let i = cab.linha + 1; i < valores.length; i++) {
        if (ie !== -1 && valores[i][ie]) empresas.push(String(valores[i][ie]).trim());
        if (is !== -1 && valores[i][is]) supervisores.push(String(valores[i][is]).trim());
        if (ir !== -1 && valores[i][ir]) realizados.push(String(valores[i][ir]).trim());
    }

    preencherSelect("empresa", empresas, "Todas");
    preencherSelect("supervisor", supervisores, "Todos");
    preencherSelect("realizadoPor", realizados, "Todos");
}

function preencherListasFallback() {
    preencherSelect("empresa", todos.map(r => r.empresa), "Todas");
    preencherSelect("supervisor", todos.map(r => r.supervisor), "Todos");
    preencherSelect("realizadoPor", todos.map(r => r.realizadoPor), "Todos");
}

function preencherSelect(id, valores, padrao) {
    const select = document.getElementById(id);
    if (!select) return;
    const atual = select.value;
    select.innerHTML = "";

    const primeira = document.createElement("option");
    primeira.value = "";
    primeira.textContent = padrao;
    select.appendChild(primeira);

    [...new Set(valores.map(v => String(v).trim()).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, "pt-BR"))
        .forEach(v => {
            const option = document.createElement("option");
            option.value = v;
            option.textContent = v;
            select.appendChild(option);
        });

    if ([...select.options].some(o => o.value === atual)) select.value = atual;
}

function pesquisar() {
    const tipo = document.getElementById("tipo")?.value || "";
    const inicio = document.getElementById("dataInicio")?.value || "";
    const fim = document.getElementById("dataFim")?.value || "";
    const empresa = normalizar(document.getElementById("empresa")?.value || "");
    const realizado = normalizar(document.getElementById("realizadoPor")?.value || "");
    const supervisor = normalizar(document.getElementById("supervisor")?.value || "");

    const inicioMs = inicio ? new Date(inicio + "T00:00:00").getTime() : null;
    const fimMs = fim ? new Date(fim + "T23:59:59").getTime() : null;

    registros = todos.filter(r => {
        const t = dataNumero(r.data);
        return (!tipo || r.tipo === tipo) &&
            (!empresa || normalizar(r.empresa) === empresa) &&
            (!realizado || normalizar(r.realizadoPor) === realizado) &&
            (!supervisor || normalizar(r.supervisor) === supervisor) &&
            (inicioMs === null || t >= inicioMs) &&
            (fimMs === null || t <= fimMs);
    });

    render();
}

function ordenar(lista) {
    const ordem = document.getElementById("ordenacao")?.value || "recente";
    return lista.sort((a, b) => {
        if (ordem === "antiga") return dataNumero(a.data) - dataNumero(b.data);
        if (ordem === "empresa") return normalizar(a.empresa).localeCompare(normalizar(b.empresa), "pt-BR") || dataNumero(b.data) - dataNumero(a.data);
        if (ordem === "realizado") return normalizar(a.realizadoPor).localeCompare(normalizar(b.realizadoPor), "pt-BR") || dataNumero(b.data) - dataNumero(a.data);
        if (ordem === "supervisor") return normalizar(a.supervisor).localeCompare(normalizar(b.supervisor), "pt-BR") || dataNumero(b.data) - dataNumero(a.data);
        return dataNumero(b.data) - dataNumero(a.data);
    });
}

function render() {
    const lista = document.getElementById("lista");
    const contador = document.getElementById("contador");
    if (!lista || !contador) return;

    const ordenados = ordenar([...registros]);
    contador.textContent = ordenados.length + (ordenados.length === 1 ? " registro" : " registros");
    lista.innerHTML = "";

    ordenados.slice(0, 300).forEach((r, i) => {
        const card = document.createElement("div");
        card.className = "registro";
        card.innerHTML =
            '<div class="registro-cab"><span class="registro-tipo">' +
            escapar(r.tipo === "Email" ? "E-mail" : "SMS") + " · " + escapar(r.empresa) +
            '</span><span class="registro-data">' + escapar(dataTexto(r.data)) + '</span></div>' +
            '<div class="registro-grid">' +
            '<div><small>ID</small>' + escapar(r.id || "—") + '</div>' +
            '<div><small>Realizado por</small>' + escapar(r.realizadoPor || "—") + '</div>' +
            '<div><small>Supervisor</small>' + escapar(r.supervisor || "—") + '</div>' +
            '<div><small>Qtde</small>' + escapar(r.qtde || "—") + '</div>' +
            '<div><small>Assunto</small>' + escapar(r.assunto || "—") + '</div>' +
            '<div class="campo-mensagem"><small>Mensagem</small><textarea class="mensagem-historico" readonly>' + escapar(r.mensagem || "—") + '</textarea></div>' +
            '<div><small>Histórico</small>' + escapar(r.historico || "—") + '</div>' +
            '</div><div class="registro-acoes"><button type="button" class="btn-ir" data-index="' + i + '">Ir para registro</button><button type="button" class="btn-duplicar" data-index="' + i + '">Duplicar</button></div>';
        lista.appendChild(card);
    });

    lista.querySelectorAll(".btn-ir").forEach(button => {
        button.addEventListener("click", () => irParaRegistro(ordenados[Number(button.dataset.index)]));
    });

    lista.querySelectorAll(".btn-duplicar").forEach(button => {
        button.addEventListener("click", () => duplicar(ordenados[Number(button.dataset.index)]));
    });

    if (!ordenados.length) lista.innerHTML = '<div class="registro">Nenhum registro encontrado.</div>';
}

async function irParaRegistro(r) {
    if (!r || !Number.isInteger(r.linhaPlanilha)) {
        erro("Não foi possível localizar a linha deste registro.");
        return;
    }

    try {
        await Excel.run(async context => {
            const sheet = context.workbook.worksheets.getItem(r.tipo);
            const celula = sheet.getRangeByIndexes(r.linhaPlanilha, 0, 1, 1);
            sheet.activate();
            celula.select();
            await context.sync();
        });
    } catch (e) {
        erro("Não foi possível ir para o registro. " + (e?.message || ""));
    }
}

function duplicar(r) {
    localStorage.setItem("registroDuplicado", JSON.stringify({
        tipo: r.tipo === "Email" ? "email" : "sms",
        realizadoPor: r.realizadoPor,
        empresa: r.empresa,
        qtde: r.qtde,
        supervisor: r.supervisor,
        historicoExterno: r.historico,
        emailResposta: r.emailResposta,
        assunto: r.assunto,
        mensagem: r.mensagem
    }));
    window.location.href = r.tipo === "Email" ? "email.html" : "sms.html";
}

function limpar() {
    document.getElementById("tipo").value = "";
    document.getElementById("dataInicio").value = "";
    document.getElementById("dataFim").value = "";
    document.getElementById("empresa").value = "";
    document.getElementById("realizadoPor").value = "";
    document.getElementById("supervisor").value = "";
    const ordenacao = document.getElementById("ordenacao");
    if (ordenacao) ordenacao.value = "recente";
    pesquisar();
}

function erro(mensagem) {
    const status = document.getElementById("status");
    if (!status) return;
    status.textContent = mensagem;
    status.className = "status erro";
    status.style.display = "block";
}