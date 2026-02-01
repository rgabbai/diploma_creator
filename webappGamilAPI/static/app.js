const htmlContent = document.getElementById("html-content");
const textContent = document.getElementById("text-content");
const htmlFile = document.getElementById("html-file");
const textFile = document.getElementById("text-file");
const htmlPanel = document.getElementById("html-panel");
const textPanel = document.getElementById("text-panel");
const formatInputs = document.querySelectorAll("input[name='letter_format']");
const preview = document.getElementById("preview");
const log = document.getElementById("log");
const forceRtl = document.getElementById("force-rtl");
const form = document.getElementById("send-form");
const emojiBars = document.querySelectorAll(".emoji-bar");
const previewCsvButton = document.getElementById("preview-csv");
const selectEligibleButton = document.getElementById("select-eligible");
const selectAllButton = document.getElementById("select-all");
const clearSelectionButton = document.getElementById("clear-selection");
const markNewButton = document.getElementById("mark-new");
const clearNewButton = document.getElementById("clear-new");
const markSendButton = document.getElementById("mark-send");
const clearSendButton = document.getElementById("clear-send");
const downloadCsvButton = document.getElementById("download-csv");
const saveToOutputButton = document.getElementById("save-to-output");
const saveToPathButton = document.getElementById("save-to-path");
const savePathInput = document.getElementById("save-path");
const csvRows = document.getElementById("csv-rows");
const csvStats = document.getElementById("csv-stats");
const csvInput = form.querySelector("input[name='csv_file']");
const pdfInput = form.querySelector("input[name='pdf_template']");
const testEmailInput = form.querySelector("input[name='test_email']");
const testNameInput = form.querySelector("input[name='test_name']");
const pdfPreview = document.getElementById("pdf-preview");
const logoInput = form.querySelector("input[name='logo_file']");
const oauthStatus = document.getElementById("oauth-status");
const oauthConnect = document.getElementById("oauth-connect");
const oauthLogout = document.getElementById("oauth-logout");
const oauthCheckSetup = document.getElementById("oauth-check-setup");
const fromEmailInput = document.getElementById("from-email");
const nameXOffsetInput = document.getElementById("name-x-offset");
const nameYOffsetInput = document.getElementById("name-y-offset");
const pdfXMinusButton = document.getElementById("pdf-x-minus");
const pdfXPlusButton = document.getElementById("pdf-x-plus");
const pdfYMinusButton = document.getElementById("pdf-y-minus");
const pdfYPlusButton = document.getElementById("pdf-y-plus");
const pdfXValue = document.getElementById("pdf-x-value");
const pdfYValue = document.getElementById("pdf-y-value");

let lastCsvRows = [];
let lastCsvHeaders = [];
let lastCsvFilename = "updated_list.csv";
const NAME_OFFSET_STEP = 5;

const escapeHtml = (value) =>
  value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");

const textToHtml = (value) => {
  const cleaned = value.replace(/\r\n/g, "\n").trim();
  if (!cleaned) return "";
  const paragraphs = cleaned.split(/\n\s*\n/);
  const body = paragraphs
    .map((para) => {
      const lines = para.split("\n").map(escapeHtml).join("<br>");
      return `<p style=\"line-height: 1.6;\">${lines}</p>`;
    })
    .join("");
  return `<!DOCTYPE html><html lang=\"he\"><head><meta charset=\"UTF-8\"></head><body style=\"direction: rtl; text-align: right; font-family: Arial, sans-serif; color: #333; margin: 20px;\">${body}</body></html>`;
};

const getFormat = () => {
  const selected = document.querySelector("input[name='letter_format']:checked");
  return selected ? selected.value : "html";
};

const updatePreview = () => {
  const format = getFormat();
  let content = format === "text" ? textToHtml(textContent.value || "") : htmlContent.value || "";
  content = applyLogoPreview(content);
  const wrapperStart = forceRtl.checked ? "<div dir=\"rtl\">" : "";
  const wrapperEnd = forceRtl.checked ? "</div>" : "";
  preview.srcdoc = `${wrapperStart}${content}${wrapperEnd}`;
};

const syncFormatPanels = () => {
  const format = getFormat();
  htmlPanel.classList.toggle("hidden", format !== "html");
  textPanel.classList.toggle("hidden", format !== "text");
};

let pdfPreviewTimer = null;
const updatePdfPreview = async () => {
  if (!pdfPreview) return;
  if (!pdfInput || !pdfInput.files.length) {
    pdfPreview.removeAttribute("src");
    return;
  }
  if (!testNameInput || !testNameInput.value.trim()) {
    pdfPreview.removeAttribute("src");
    return;
  }
  const data = new FormData();
  data.set("pdf_template", pdfInput.files[0]);
  data.set("test_name", testNameInput.value.trim());
  if (nameXOffsetInput) {
    data.set("name_x_offset", nameXOffsetInput.value || "0");
  }
  if (nameYOffsetInput) {
    data.set("name_y_offset", nameYOffsetInput.value || "0");
  }

  log.textContent = "Generating PDF preview...";
  try {
    const response = await fetch("/api/preview-pdf", { method: "POST", body: data });
    const payload = await response.json();
    if (!response.ok || !payload.ok) {
      log.textContent = payload.error || "Failed to preview PDF";
      return;
    }
    const cacheBust = `t=${Date.now()}`;
    pdfPreview.src = `${payload.pdf_url}?${cacheBust}#toolbar=0&navpanes=0&scrollbar=0`;
    log.textContent = "PDF preview ready.";
  } catch (err) {
    log.textContent = err.message;
  }
};

const updateOauthStatus = async () => {
  if (!oauthStatus) return;
  try {
    const response = await fetch("/oauth/status");
    const payload = await response.json();
    oauthStatus.value = payload.ok ? "Connected" : "Not connected";
    if (payload.email && fromEmailInput) {
      fromEmailInput.value = payload.email;
    }
  } catch (err) {
    oauthStatus.value = "Not connected";
  }
};

htmlContent.addEventListener("input", updatePreview);
textContent.addEventListener("input", updatePreview);
forceRtl.addEventListener("change", updatePreview);

htmlFile.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    htmlContent.value = e.target.result;
    updatePreview();
  };
  reader.readAsText(file);
});

textFile.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    textContent.value = e.target.result;
    updatePreview();
  };
  reader.readAsText(file);
});

let logoDataUrl = "";
const applyLogoPreview = (content) => {
  if (!logoDataUrl) return content;
  if (content.includes("cid:logo_cid")) {
    return content.replace(/cid:logo_cid/g, logoDataUrl);
  }
  const imgTag = `<img src="${logoDataUrl}" style="max-width: 220px; height: auto; display: block; margin-bottom: 12px;">`;
  const bodyCloseMatch = content.match(/<\/body>/i);
  if (bodyCloseMatch) {
    const idx = content.toLowerCase().lastIndexOf("</body>");
    return content.slice(0, idx) + imgTag + content.slice(idx);
  }
  return content + imgTag;
};

if (logoInput) {
  logoInput.addEventListener("change", () => {
    const file = logoInput.files && logoInput.files[0];
    if (!file) {
      logoDataUrl = "";
      updatePreview();
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      logoDataUrl = e.target.result || "";
      updatePreview();
    };
    reader.readAsDataURL(file);
  });
}

if (pdfInput) {
  pdfInput.addEventListener("change", () => {
    updatePdfPreview();
  });
}

if (testNameInput) {
  testNameInput.addEventListener("input", () => {
    if (pdfPreviewTimer) {
      clearTimeout(pdfPreviewTimer);
    }
    pdfPreviewTimer = setTimeout(updatePdfPreview, 500);
  });
}

const readOffsetValue = (input) => {
  const parsed = Number.parseInt(input?.value ?? "0", 10);
  return Number.isFinite(parsed) ? parsed : 0;
};

const syncOffsetDisplay = () => {
  if (pdfXValue && nameXOffsetInput) {
    pdfXValue.textContent = `${readOffsetValue(nameXOffsetInput)}`;
  }
  if (pdfYValue && nameYOffsetInput) {
    pdfYValue.textContent = `${readOffsetValue(nameYOffsetInput)}`;
  }
};

const updateNameOffsets = (deltaX, deltaY) => {
  if (nameXOffsetInput) {
    const nextX = readOffsetValue(nameXOffsetInput) + deltaX;
    nameXOffsetInput.value = `${nextX}`;
  }
  if (nameYOffsetInput) {
    const nextY = readOffsetValue(nameYOffsetInput) + deltaY;
    nameYOffsetInput.value = `${nextY}`;
  }
  syncOffsetDisplay();
  updatePdfPreview();
};

if (pdfXMinusButton) {
  pdfXMinusButton.addEventListener("click", () => updateNameOffsets(-NAME_OFFSET_STEP, 0));
}
if (pdfXPlusButton) {
  pdfXPlusButton.addEventListener("click", () => updateNameOffsets(NAME_OFFSET_STEP, 0));
}
if (pdfYMinusButton) {
  pdfYMinusButton.addEventListener("click", () => updateNameOffsets(0, -NAME_OFFSET_STEP));
}
if (pdfYPlusButton) {
  pdfYPlusButton.addEventListener("click", () => updateNameOffsets(0, NAME_OFFSET_STEP));
}

formatInputs.forEach((input) => {
  input.addEventListener("change", () => {
    syncFormatPanels();
    updatePreview();
  });
});

const applyCsvOverride = (data) => {
  if (!csvInput || !csvInput.files.length || !lastCsvRows.length) {
    return;
  }
  const csvContent = buildCsvContent();
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  data.set("csv_file", blob, lastCsvFilename || "updated_list.csv");
};

const sendFormData = async (endpoint, options = {}) => {
  const data = new FormData(form);
  data.set("html_content", htmlContent.value || "");
  data.set("text_content", textContent.value || "");
  data.set("selected_indices", getSelectedIndices().join(","));
  applyCsvOverride(data);

  if (options.requirePdf && (!pdfInput || !pdfInput.files.length)) {
    log.textContent = "Missing PDF template. Please load a PDF file first.";
    return;
  }
  if (options.requireTestEmail && (!testEmailInput || !testEmailInput.value.trim())) {
    log.textContent = "Missing test email.";
    return;
  }

  log.textContent = "Sending...";
  try {
    const response = await fetch(endpoint, { method: "POST", body: data });
    const payload = await response.json();
    if (!response.ok) {
      log.textContent = payload.error || "Request failed";
      return;
    }
    if (payload.sent_indices) {
      markRowsByIndex(payload.sent_indices, "send", "1");
      renderCsvRows(lastCsvRows);
    }
    if (payload.log_lines && Array.isArray(payload.log_lines) && payload.log_lines.length) {
      let output = payload.log_lines.join("\n");
      if (payload.summary) {
        output += `\n\nsummary: sent=${payload.summary.sent || 0} skipped=${payload.summary.skipped || 0}`;
      }
      log.textContent = output;
    } else {
      log.textContent = JSON.stringify(payload, null, 2);
    }
  } catch (err) {
    log.textContent = err.message;
  }
};

const sendBatchStream = async () => {
  if (!csvInput || !csvInput.files.length) {
    log.textContent = "Missing CSV file. Please load a CSV file first.";
    return;
  }
  if (!pdfInput || !pdfInput.files.length) {
    log.textContent = "Missing PDF template. Please load a PDF file first.";
    return;
  }

  const data = new FormData(form);
  data.set("html_content", htmlContent.value || "");
  data.set("text_content", textContent.value || "");
  data.set("selected_indices", getSelectedIndices().join(","));
  applyCsvOverride(data);

  log.textContent = "Sending batch...";
  try {
    const response = await fetch("/api/send-stream", { method: "POST", body: data });
    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      log.textContent = payload.error || "Request failed";
      return;
    }
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    const lines = [];
    let sentIndices = [];

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const parts = buffer.split("\n");
      buffer = parts.pop() || "";
      parts.forEach((line) => {
        if (!line) return;
        if (line.startsWith("json:")) {
          try {
            const payload = JSON.parse(line.slice(5));
            sentIndices = payload.sent_indices || sentIndices;
          } catch (err) {
            lines.push(line);
          }
        } else {
          lines.push(line);
        }
        log.textContent = lines.join("\n");
      });
    }

    if (sentIndices.length) {
      markRowsByIndex(sentIndices, "send", "1");
      renderCsvRows(lastCsvRows);
    }
  } catch (err) {
    log.textContent = err.message;
  }
};

form.addEventListener("click", (event) => {
  const action = event.target.getAttribute("data-action");
  if (!action) return;
  if (action === "test") {
    sendFormData("/api/test-send", { requirePdf: true, requireTestEmail: true });
  }
  if (action === "batch") {
    const ok = window.confirm("Are you sure - Did you test it before");
    if (!ok) {
      return;
    }
    sendBatchStream();
  }
});

updatePreview();
syncFormatPanels();
updateOauthStatus();
syncOffsetDisplay();

if (oauthConnect) {
  oauthConnect.addEventListener("click", () => {
    window.open("/oauth/start", "_blank");
    let attempts = 0;
    const poll = setInterval(async () => {
      attempts += 1;
      await updateOauthStatus();
      if (oauthStatus && oauthStatus.value === "Connected") {
        clearInterval(poll);
      }
      if (attempts >= 20) {
        clearInterval(poll);
      }
    }, 1000);
  });
}

if (oauthLogout) {
  oauthLogout.addEventListener("click", async () => {
    try {
      await fetch("/oauth/logout", { method: "POST" });
    } finally {
      updateOauthStatus();
    }
  });
}

if (oauthCheckSetup) {
  oauthCheckSetup.addEventListener("click", async () => {
    log.textContent = "Checking setup...";
    try {
      const response = await fetch("/oauth/check-setup");
      const payload = await response.json();
      if (payload.ok) {
        log.textContent = `Setup OK. Client secret found at: ${payload.client_secret_path}`;
      } else {
        log.textContent = `Setup missing. Expected client secret at: ${payload.client_secret_path}`;
      }
    } catch (err) {
      log.textContent = err.message;
    }
  });
}

emojiBars.forEach((bar) => {
  bar.addEventListener("click", (event) => {
    const button = event.target.closest(".emoji-btn");
    if (!button) return;
    const targetId = bar.dataset.target;
    const target = document.getElementById(targetId);
    if (!target) return;
    const emoji = button.dataset.emoji || "";
    const start = target.selectionStart || 0;
    const end = target.selectionEnd || 0;
    const value = target.value || "";
    target.value = `${value.slice(0, start)}${emoji}${value.slice(end)}`;
    const cursor = start + emoji.length;
    target.setSelectionRange(cursor, cursor);
    target.focus();
    updatePreview();
  });
});

const renderCsvRows = (rows) => {
  csvRows.innerHTML = "";
  if (!rows.length) {
    csvStats.textContent = "No CSV rows";
    return;
  }
  const eligibleCount = rows.filter((row) => row.eligible).length;
  csvStats.textContent = `${rows.length} rows, ${eligibleCount} eligible`;

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    if (row.send === "1") {
      tr.classList.add("csv-row-sent");
    } else if (row.eligible) {
      tr.classList.add("csv-row-eligible");
    }

    const checked = row.selected ?? row.eligible;
    tr.innerHTML = `
      <td><input type="checkbox" data-index="${row.index}" ${checked ? "checked" : ""}></td>
      <td>${row.name || ""}</td>
      <td>${row.email || ""}</td>
      <td>${row.new || ""}</td>
      <td>${row.send || ""}</td>
    `;
    csvRows.appendChild(tr);
  });
};

const loadCsvPreview = async () => {
  if (!csvInput.files.length) {
    csvStats.textContent = "Choose a CSV file first";
    return;
  }
  lastCsvFilename = csvInput.files[0].name || "updated_list.csv";
  const data = new FormData();
  data.set("csv_file", csvInput.files[0]);
  csvStats.textContent = "Loading...";
  try {
    const response = await fetch("/api/preview", { method: "POST", body: data });
    const payload = await response.json();
    if (!response.ok || !payload.ok) {
      csvStats.textContent = payload.error || "Failed to parse CSV";
      return;
    }
    lastCsvRows = payload.rows || [];
    lastCsvHeaders = payload.headers || [];
    lastCsvRows.forEach((row) => {
      row.selected = row.eligible;
    });
    renderCsvRows(lastCsvRows);
  } catch (err) {
    csvStats.textContent = err.message;
  }
};

const getSelectedIndices = () => {
  return Array.from(csvRows.querySelectorAll("input[type='checkbox']:checked"))
    .map((input) => Number.parseInt(input.dataset.index, 10))
    .filter((value) => Number.isFinite(value));
};

const markRowsByIndex = (indices, field, value) => {
  indices.forEach((index) => {
    const row = lastCsvRows.find((item) => item.index === index);
    if (!row) return;
    row[field] = value;
    if (row.raw) {
      row.raw[field] = value;
    }
    row.eligible = row.new === "1" && row.send !== "1";
  });
};

const syncSelectionToNew = () => {
  const selected = getSelectedIndices();
  markRowsByIndex(selected, "new", "1");
};

previewCsvButton.addEventListener("click", loadCsvPreview);

csvInput.addEventListener("change", () => {
  lastCsvRows = [];
  csvRows.innerHTML = "";
  loadCsvPreview();
});

selectEligibleButton.addEventListener("click", () => {
  csvRows.querySelectorAll("input[type='checkbox']").forEach((input) => {
    const index = Number.parseInt(input.dataset.index, 10);
    const row = lastCsvRows.find((item) => item.index === index);
    input.checked = row ? row.eligible : false;
    if (row) {
      row.selected = input.checked;
    }
  });
});

selectAllButton.addEventListener("click", () => {
  csvRows.querySelectorAll("input[type='checkbox']").forEach((input) => {
    input.checked = true;
    const index = Number.parseInt(input.dataset.index, 10);
    const row = lastCsvRows.find((item) => item.index === index);
    if (row) {
      row.selected = true;
    }
  });
});

clearSelectionButton.addEventListener("click", () => {
  csvRows.querySelectorAll("input[type='checkbox']").forEach((input) => {
    input.checked = false;
    const index = Number.parseInt(input.dataset.index, 10);
    const row = lastCsvRows.find((item) => item.index === index);
    if (row) {
      row.selected = false;
    }
  });
});

csvRows.addEventListener("change", (event) => {
  const checkbox = event.target;
  if (checkbox.type !== "checkbox") return;
  const index = Number.parseInt(checkbox.dataset.index, 10);
  const row = lastCsvRows.find((item) => item.index === index);
  if (!row) return;
  row.selected = checkbox.checked;
  if (checkbox.checked) {
    markRowsByIndex([index], "new", "1");
  }
  renderCsvRows(lastCsvRows);
});

markNewButton.addEventListener("click", () => {
  markRowsByIndex(getSelectedIndices(), "new", "1");
  renderCsvRows(lastCsvRows);
});

clearNewButton.addEventListener("click", () => {
  markRowsByIndex(getSelectedIndices(), "new", "");
  renderCsvRows(lastCsvRows);
});

markSendButton.addEventListener("click", () => {
  markRowsByIndex(getSelectedIndices(), "send", "1");
  renderCsvRows(lastCsvRows);
});

clearSendButton.addEventListener("click", () => {
  markRowsByIndex(getSelectedIndices(), "send", "");
  renderCsvRows(lastCsvRows);
});

downloadCsvButton.addEventListener("click", () => {
  if (!lastCsvRows.length) {
    csvStats.textContent = "No CSV rows to download";
    return;
  }
  const csvContent = buildCsvContent();
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = lastCsvFilename || "updated_list.csv";
  link.click();
  URL.revokeObjectURL(url);
});

const buildCsvContent = () => {
  const headers = lastCsvHeaders.length ? lastCsvHeaders : ["first", "last", "email", "new", "send"];
  const lines = [headers.join(",")];
  lastCsvRows.forEach((row) => {
    const raw = row.raw || {};
    const line = headers
      .map((key) => {
        const value = (raw[key] ?? "").toString();
        if (value.includes("\"") || value.includes(",") || value.includes("\\n")) {
          return `"${value.replace(/\"/g, "\"\"")}"`;
        }
        return value;
      })
      .join(",");
    lines.push(line);
  });
  return lines.join("\n");
};

const saveCsvToServer = async (targetPath) => {
  if (!lastCsvRows.length) {
    csvStats.textContent = "No CSV rows to save";
    return;
  }
  const data = new FormData();
  data.set("csv_content", buildCsvContent());
  data.set("filename", lastCsvFilename || "updated_list.csv");
  data.set("target_path", targetPath || "");
  csvStats.textContent = "Saving...";
  try {
    const response = await fetch("/api/save-csv", { method: "POST", body: data });
    const payload = await response.json();
    if (!response.ok || !payload.ok) {
      csvStats.textContent = payload.error || "Failed to save CSV";
      return;
    }
    csvStats.textContent = `Saved to ${payload.saved_to}`;
  } catch (err) {
    csvStats.textContent = err.message;
  }
};

saveToOutputButton.addEventListener("click", () => {
  saveCsvToServer("");
});

saveToPathButton.addEventListener("click", () => {
  const targetPath = savePathInput.value.trim();
  if (!targetPath) {
    csvStats.textContent = "Enter a save path first";
    return;
  }
  saveCsvToServer(targetPath);
});
