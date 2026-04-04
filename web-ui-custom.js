(() => {
  function ensureSetupBanner(appRoot) {
    if (appRoot.querySelector("[data-hubvoice-setup-banner]")) {
      return;
    }

    const banner = document.createElement("div");
    banner.setAttribute("data-hubvoice-setup-banner", "true");
    banner.style.margin = "12px 16px";
    banner.style.padding = "12px 14px";
    banner.style.borderRadius = "12px";
    banner.style.background = "#0e3a57";
    banner.style.color = "#ffffff";
    banner.style.fontSize = "14px";
    banner.style.lineHeight = "1.4";
    banner.innerHTML = "<strong>Wi-Fi setup</strong><br>When connected to the hub setup Wi-Fi, enter <strong>WiFi Setup SSID</strong> and <strong>WiFi Setup Password</strong>, then press <strong>Apply WiFi Credentials</strong>.";

    const layout = appRoot.querySelector("ha-drawer partial-panel-resolver, partial-panel-resolver");
    if (layout && layout.parentElement) {
      layout.parentElement.insertBefore(banner, layout);
      return;
    }

    appRoot.prepend(banner);
  }

  function syncPage() {
    const app = document.querySelector("esp-app");
    const appRoot = app && app.shadowRoot;
    if (!appRoot) {
      return;
    }

    ensureSetupBanner(appRoot);

    const table = appRoot.querySelector("esp-entity-table");
    const tableRoot = table && table.shadowRoot;
    if (!tableRoot) {
      return;
    }

    const rows = Array.from(tableRoot.querySelectorAll("tbody tr"));
    const wifiRows = [];
    for (const row of rows) {
      const cells = row.querySelectorAll("td");
      const name = (cells[0]?.textContent || "").trim();
      if (name === "Effective Satellite Name") {
        row.style.display = "none";
      }
      if (["WiFi Setup SSID", "WiFi Setup Password", "WiFi Setup Status", "Apply WiFi Credentials"].includes(name)) {
        wifiRows.push(row);
      }
    }

    const tbody = tableRoot.querySelector("tbody");
    if (tbody) {
      for (const row of wifiRows.reverse()) {
        tbody.prepend(row);
      }
    }
  }

  window.addEventListener("load", () => {
    syncPage();
    window.setInterval(syncPage, 1000);
  });
})();
