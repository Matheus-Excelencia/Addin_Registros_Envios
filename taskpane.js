Office.onReady(function () {

    const btnContinuar = document.getElementById("btnContinuar");

    btnContinuar.addEventListener("click", function () {

        const tipo = document.getElementById("tipoRegistro").value;

        if (tipo === "email") {
            window.location.href = "email.html";
        }

        if (tipo === "sms") {
            window.location.href = "sms.html";
        }

    });

});
