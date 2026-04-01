(() => {
  function syncPage() {
    const app = document.querySelector("esp-app");
    const appRoot = app && app.shadowRoot;
    if (!appRoot) {
      return;
    }

    const table = appRoot.querySelector("esp-entity-table");
    const tableRoot = table && table.shadowRoot;
    if (!tableRoot) {
      return;
    }

    const rows = Array.from(tableRoot.querySelectorAll("tbody tr"));
    for (const row of rows) {
      const cells = row.querySelectorAll("td");
      const name = (cells[0]?.textContent || "").trim();
      if (name === "Effective Satellite Name") {
        row.style.display = "none";
      }
    }
  }

  window.addEventListener("load", () => {
    syncPage();
    window.setInterval(syncPage, 1000);
  });
})();
