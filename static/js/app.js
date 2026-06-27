document.addEventListener("submit", (event) => {
  const form = event.target;
  const message = form.getAttribute("data-confirm");
  if (message && !window.confirm(message)) {
    event.preventDefault();
  }
});

const progressPanel = document.querySelector("[data-job-id]");

if (progressPanel) {
  const jobId = progressPanel.dataset.jobId;
  const progressStatus = document.querySelector("#progressStatus");
  const progressPercent = document.querySelector("#progressPercent");
  const progressFill = document.querySelector("#progressFill");
  const processedCount = document.querySelector("#processedCount");
  const totalCount = document.querySelector("#totalCount");
  const importedCount = document.querySelector("#importedCount");
  const skippedCount = document.querySelector("#skippedCount");
  const currentProduct = document.querySelector("#currentProduct");
  const progressError = document.querySelector("#progressError");
  const catalogLink = document.querySelector("#catalogLink");

  async function refreshProgress() {
    const response = await fetch(`/api/imports/${jobId}`, { cache: "no-store" });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Progress could not be loaded.");
    }

    progressStatus.textContent = data.status === "complete"
      ? "Complete"
      : data.status === "error"
        ? "Stopped"
        : "Importing";
    progressPercent.textContent = `${data.percent}%`;
    progressFill.style.width = `${data.percent}%`;
    processedCount.textContent = data.processed_count;
    totalCount.textContent = data.total_count || "...";
    importedCount.textContent = data.imported_count;
    skippedCount.textContent = data.skipped_count;
    currentProduct.textContent = data.current_product || "Preparing next product...";

    if (data.status === "complete") {
      catalogLink.hidden = false;
      return;
    }

    if (data.status === "error") {
      progressError.hidden = false;
      progressError.textContent = data.error_message || "Import stopped because of an error.";
      return;
    }

    window.setTimeout(refreshProgress, 1500);
  }

  refreshProgress().catch((error) => {
    progressError.hidden = false;
    progressError.textContent = error.message;
  });
}
