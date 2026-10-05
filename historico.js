let todos = [], registros = [], pronto = false;

function iniciarHistorico() {
    if (pronto) return;
    pronto = true;

    const v = document.getElementById("btnVoltar");
    const p = document.getElementById("btnPesquisar");
    const l = document.getElementById("btnLimpar");

    if (v) v.onclick = () => { window.location.href = "taskpane.html"; };
    if (p) p.onclick = () => pesquisar();
    if (l) l.onclick = () => limpar();

    const status = document.getElementById("status");
    if (status) {
        status.textContent = "Carregando registros...";
        status.className = "status";
        status.style.display = "block";
    }

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

function encontrarCabecalho(valores) {
    for (let linha = 0; linha < Math.min(valores.length, 30); linha++) {
        const mapa = {};
        (valores[linha] || []).forEach((valor, coluna) => {
            if (String(valor ?? "").trim()) mapa[normalizar(valor)] = coluna;
        });

        if (indice(mapa, ["EMPRESA"]) >= 0 && indice(mapa, ["DATA"]) >= 0) {
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
    if (!texto) return 0;

    const br = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (br) {
        return new Date(
            Number(br[3]),
            Number(br[2]) - 1,
            Number(br[1]),
            Number(br[4] || 0),
            Number(br[5] || 0),
            Number(br[6] || 0)
        ).getTime();
    }

    const d = new Date(text);
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
                range.load("values,isNullObject");
                return { nome, range };
            });

            await context.sync();

            for (const item of ranges) {
                if (item.range.isNullObject) continue;

                const valores = item.range.values || [];
                const cabecalho = encontrarCabecalho(valores);
                if (!cabecalho) continue;

                for (let i = cabecalho.linha + 1; i < valores.length; i++) {
                    const registro = valores[i] || [];
                    const empresa = String(valor(registro, cabecalho.mapa, ["EMPRESA"])).trim();
                    if (!empresa) continue;

                    todos.push({
                        tipo: item.nome,
                        data: valor(registro, cabecalho.mapa, ["DATA"]),
                        empresa,
                        realizadoPor: String(valor(registro, cabecalho.mapa, ["REALIZADO POR"])).trim(),
                        supervisor: String(valor(registro, cabecalho.mapa, ["SUPERVISOR"])).trim(),
                        qtde: String(valor(registro, cabecalho.mapa, ["QTDE"])).trim(),
                        historico: String(valor(registro, cabecalho.mapa, ["HISTORICO EXTERNO", "HISTÓRICO EXTERNO"])).trim(),
                        mensagem: String(valor(registro, cabecalho.mapa, ["MENSAGEM"])).trim(),
                        assunto: String(valor(registro, cabecalho.mapa, ["ASSUNTO"])).trim(),
                        id: String(valor(registro, cabecalho.mapa, ["ID REGISTRO"])).trim(),
                        emailResposta: String(valor(registro, cabecalho.mapa, ["EMAIL RESPOSTA", "E-MAIL RESPOSTA"])).trim()
                    });
                }
            }
        });

        preencherListas();
        pesquisar();

        const status = document.getElementById("status");
        if (status) status.style.display = "none";
    } catch (e) {
        erro("Não foi possível carregar os registros. " + (e?.message || "Verifique se as planilhas Email e SMS estão disponíveis."));
    }
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

    [...new Set(valores.filter(v => String(v).trim()))]
        .sort((a, b) => String(a).localeCompare(String(b), "pt-BR"))
        .forEach(v => {
            const option = document.createElement("option");
            option.value = v;
            option.textContent = v;
            select.appendChild(option);
        });

    if ([...select.options].some(o => o.value === atual)) select.value = atual;
}

function preencherListas() {
    preencherSelect("empresa", todos.map(r => r.empresa), "Todas");
    preencherSelect("realizadoPor", todos.map(r => r.realizadoPor), "Todos");
    preencherSelect("supervisor", todos.map(r => r.supervisor), "Todos");
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
    }).sort((a, b) => dataNumero(b.data) - dataNumero(a.data));

    render();
}

function render() {
    const lista = document.getElementById("lista");
    const contador = document.getElementById("contador");
    if (!lista || !contador) return;

    contador.textContent = registros.length + (registros.length === 1 ? " registro" : " registros");
    lista.innerHTML = "";

    registros.slice(0, 300).forEach((r, i) => {
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
            '<div><small>Mensagem</small>' + escapar(r.mensagem || "—") + '</div>' +
            '<div><small>Histórico</small>' + escapar(r.historico || "—") + '</div>' +
            '</div><div class="registro-acoes">' +
            '<button type="button" class="btn-duplicar" data-i="' + i + '">Duplicar</button></div>';

        lista.appendChild(card);
    });

    lista.querySelectorAll(".btn-duplicar").forEach(button => {
        button.onclick = () => duplicar(registros[Number(button.dataset.i)]);
    });

    if (!registros.length) {
        lista.innerHTML = '<div class="registro">Nenhum registro encontrado.</div>';
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
    pesquisar();
}

function erro(mensagem) {
    const status = document.getElementById("status");
    if (!status) return;
    status.textContent = mensagem;
    status.className = "status erro";
    status.style.display = "block";
}