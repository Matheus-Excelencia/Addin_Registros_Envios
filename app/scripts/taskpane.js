Office.onReady(function(){
    const go = u => window.location.assign(u);
    const n = document.getElementById("btnNovo");
    const h = document.getElementById("btnHistorico");
    const r = document.getElementById("btnResumo");
    const c = document.getElementById("btnConfig");

    if (n) n.addEventListener("click", () => go("novo.html"));
    if (h) h.addEventListener("click", () => go("historico.html"));
    if (r) r.addEventListener("click", () => go("resumo.html"));
    if (c) c.addEventListener("click", () => go("config.html"));
});