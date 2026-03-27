(function () {
  "use strict";

  window.CTConfig = {
    appName: "Monitoramento de Entregas SP",
    companyName: "J&T Express",
    version: "2.0.0",
    pages: {
      dashboard: "../index.html",
      insucessos: "../pages/insucessos.html"
    },
    storageKeys: {
      monitoramento: "monitoramento_local_state_v2",
      insucessos: "insucessos_local_state_v2"
    },
    dependencies: {
      dashboard: ["XLSX", "html2canvas", "Chart"],
      insucessos: ["XLSX", "Chart"]
    }
  };
})();
