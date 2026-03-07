// Popup script — shows whether the extension is active on the current tab
(function () {
  const statusEl = document.getElementById("status");

  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    const tab = tabs[0];
    if (!tab || !tab.url) {
      statusEl.className = "status inactive";
      statusEl.textContent = "Cannot read page URL.";
      return;
    }

    const isAdo =
      tab.url.includes("dev.azure.com") ||
      tab.url.includes("visualstudio.com");

    if (isAdo) {
      const isYaml = /\.ya?ml/i.test(tab.url);

      if (isYaml) {
        statusEl.className = "status active";
        statusEl.textContent =
          "Active look for the TF button in the top-right corner.";
      } else {
        statusEl.className = "status inactive";
        statusEl.textContent =
          "On ADO, but not viewing a YAML file. Navigate to a pipeline YAML to activate.";
      }
    } else {
      statusEl.className = "status inactive";
      statusEl.textContent =
        "Not on Azure DevOps. Navigate to an ADO YAML pipeline to use this extension.";
    }
  });
})();
