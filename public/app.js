const yearSelect = document.querySelector("#year");
const filesElement = document.querySelector("#files");
const uploadsElement = document.querySelector("#uploads");
const uploadInput = document.querySelector("#upload-files");
const statusElement = document.querySelector("#status");
const rateTypeSelect = document.querySelector("#rate-type");
const rateTypeForm = document.querySelector("#rate-form-type");
const isinStatusElement = document.querySelector("#isin-status");
let metadata;
let uploadedFiles = [];

function setStatus(element, message, state = "") {
  element.textContent = message;
  element.className = `status ${state}`.trim();
}

async function fetchJson(url, options) {
  const response = await fetch(url, options);
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || "Požadavek se nezdařil.");
  return result;
}

function renderFiles() {
  filesElement.replaceChildren();
  if (!metadata.files.length) {
    filesElement.textContent = "Ve složce data nejsou žádné zdrojové XLSX soubory.";
    return;
  }

  for (const file of metadata.files) {
    const label = document.createElement("label");
    label.className = "file-option";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = file.name;
    checkbox.checked = uploadedFiles.length === 0;
    const content = document.createElement("span");
    content.className = "file-meta";
    const name = document.createElement("span");
    name.className = "file-name";
    name.textContent = file.name;
    const detail = document.createElement("span");
    detail.className = "file-detail";
    detail.textContent = `${file.currency} · účet ${file.account || "nezjištěn"}`;
    content.append(name, detail);
    label.append(checkbox, content);
    filesElement.append(label);
  }
}

function renderUploads() {
  uploadsElement.replaceChildren();
  if (!uploadedFiles.length) {
    uploadsElement.textContent = "Zatím nejsou vybrané žádné soubory k nahrání.";
    return;
  }

  uploadedFiles.forEach((file, index) => {
    const label = document.createElement("label");
    label.className = "file-option";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = true;
    checkbox.dataset.uploadIndex = String(index);
    const content = document.createElement("span");
    content.className = "file-meta";
    const name = document.createElement("span");
    name.className = "file-name";
    name.textContent = file.name;
    const currency = file.name.match(/^([A-Za-z0-9]+)_/)?.[1]?.toUpperCase() || "měna nezjištěna";
    const detail = document.createElement("span");
    detail.className = "file-detail";
    detail.textContent = `${currency} · ${Math.ceil(file.size / 1024 / 1024)} MB · pouze v paměti`;
    content.append(name, detail);
    label.append(checkbox, content);
    uploadsElement.append(label);
  });
}

async function loadMetadata() {
  try {
    metadata = await fetchJson("/api/metadata");
    const oldYear = yearSelect.value;
    yearSelect.replaceChildren();
    for (const year of metadata.years) {
      const option = document.createElement("option");
      option.value = year;
      option.textContent = year;
      yearSelect.append(option);
    }
    if (metadata.years.includes(oldYear)) yearSelect.value = oldYear;
    else if (metadata.years.includes(String(new Date().getFullYear()))) {
      yearSelect.value = String(new Date().getFullYear());
    }
    document.querySelector("#rate-year").value = yearSelect.value;
    renderFiles();
    setStatus(
      isinStatusElement,
      metadata.eodhdConfigured
        ? "EODHD klíč je uložen pouze v paměti serveru."
        : "EODHD klíč není nastaven; použije se jen lokální cache.",
      metadata.eodhdConfigured ? "success" : "",
    );
  } catch (error) {
    setStatus(statusElement, error.message, "error");
  }
}

document.querySelector("#refresh").addEventListener("click", loadMetadata);
yearSelect.addEventListener("change", () => {
  document.querySelector("#rate-year").value = yearSelect.value;
});

uploadInput.addEventListener("change", () => {
  uploadedFiles = [...uploadInput.files];
  if (uploadedFiles.length) {
    for (const checkbox of filesElement.querySelectorAll("input[type=checkbox]")) {
      checkbox.checked = false;
    }
  }
  renderUploads();
});

document.querySelector("#generate").addEventListener("click", async (event) => {
  const button = event.currentTarget;
  const selectedFiles = [...filesElement.querySelectorAll("input:checked")].map((input) => input.value);
  const selectedUploads = [...uploadsElement.querySelectorAll("input:checked")]
    .map((input) => uploadedFiles[Number(input.dataset.uploadIndex)])
    .filter(Boolean);
  if (!selectedFiles.length && !selectedUploads.length) {
    setStatus(statusElement, "Vyberte nebo nahrajte alespoň jeden zdrojový soubor.", "error");
    return;
  }

  button.disabled = true;
  setStatus(statusElement, "Připravuji report…");
  try {
    const formData = new FormData();
    formData.append("year", yearSelect.value);
    formData.append("rateType", rateTypeSelect.value);
    formData.append("clientName", document.querySelector("#client-name").value);
    formData.append("storedFiles", JSON.stringify(selectedFiles));
    for (const file of selectedUploads) formData.append("uploads", file, file.name);
    const response = await fetch("/api/report", {
      method: "POST",
      body: formData,
    });
    if (!response.ok) {
      const result = await response.json();
      throw new Error(result.error || "Report se nepodařilo vytvořit.");
    }
    const blob = await response.blob();
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `Fiscal-Report-${yearSelect.value}.xlsx`;
    link.click();
    URL.revokeObjectURL(link.href);
    setStatus(statusElement, "Report byl vytvořen a stažen.", "success");
  } catch (error) {
    setStatus(statusElement, error.message, "error");
  } finally {
    button.disabled = false;
  }
});

document.querySelector("#isin-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  const keyInput = document.querySelector("#eodhd-key");
  const apiKey = keyInput.value.trim();
  if (!apiKey) {
    setStatus(isinStatusElement, "Zadejte EODHD API klíč.", "error");
    return;
  }
  try {
    const result = await fetchJson("/api/settings/eodhd", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ apiKey }),
    });
    keyInput.value = "";
    setStatus(isinStatusElement, "Klíč je uložen pouze v paměti serveru.", "success");
    if (metadata) metadata.eodhdConfigured = result.configured;
  } catch (error) {
    setStatus(isinStatusElement, error.message, "error");
  }
});

document.querySelector("#clear-eodhd-key").addEventListener("click", async () => {
  try {
    const result = await fetchJson("/api/settings/eodhd", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ apiKey: "" }),
    });
    if (metadata) metadata.eodhdConfigured = result.configured;
    setStatus(isinStatusElement, "EODHD klíč byl odebrán z paměti serveru.", "success");
  } catch (error) {
    setStatus(isinStatusElement, error.message, "error");
  }
});

function toggleDateField() {
  const dateField = document.querySelector("#date-field");
  const dateInput = document.querySelector("#rate-date");
  const isCnb = rateTypeForm.value === "CNB";
  dateField.classList.toggle("hidden", !isCnb);
  dateInput.required = isCnb;
}
rateTypeForm.addEventListener("change", toggleDateField);
toggleDateField();

document.querySelector("#rate-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  const form = event.currentTarget;
  const status = document.querySelector("#rate-status");
  const year = document.querySelector("#rate-year").value;
  try {
    await fetchJson("/api/rates", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: rateTypeForm.value,
        year,
        currency: document.querySelector("#currency").value,
        date: document.querySelector("#rate-date").value,
        rate: document.querySelector("#rate-value").value,
      }),
    });
    setStatus(status, "Kurz byl uložen.", "success");
    form.reset();
    document.querySelector("#rate-year").value = year;
    toggleDateField();
    await loadMetadata();
  } catch (error) {
    setStatus(status, error.message, "error");
  }
});

renderUploads();
loadMetadata();
