const API_BASE = "";

let apiKey = "";
let apiVerified = false;
let selectedFile = null;
let excelResults = [];
let currentProductIndex = 0;

const FIELD_KEYS = [
  "product_name", "product_type", "category", "sub_category",
  "voltage", "power", "current", "rpm",
  "material", "size", "manufacturer", "ip_rating",
  "efficiency", "weight", "notes",
  "flow_rate", "head", "pressure_rating", "end_connection",
  "operation_type", "temperature_range", "actuation", "pressure_range",
  "process_connection", "accuracy_class"
];

window.addEventListener("load", () => {
  checkServerStatus();
  setupDragDrop();
});

async function checkServerStatus() {
  const dot = document.getElementById("serverDot");
  const text = document.getElementById("serverStatusText");
  try {
    const res = await fetch(`${API_BASE}/api/health`, { signal: AbortSignal.timeout(4000) });
    const data = await res.json();
    if (data.status === "ok") {
      dot.classList.add("online");
      text.textContent = "Server Online";
    } else {
      text.textContent = "Server Error";
    }
  } catch {
    text.textContent = "Server Offline";
    dot.classList.remove("online");
  }
}

function switchTab(tabName, btnEl) {
  document.querySelectorAll(".tab-content").forEach(t => {
    t.classList.add("hidden");
    t.classList.remove("fade-in");
  });
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
  const target = document.getElementById(`tab-${tabName}`);
  target.classList.remove("hidden");
  setTimeout(() => target.classList.add("fade-in"), 10);
  btnEl.classList.add("active");
}

function toggleKeyVisibility() {
  const input = document.getElementById("apiKeyInput");
  const btn = event.target;
  if (input.type === "password") {
    input.type = "text";
    btn.textContent = "Hide";
  } else {
    input.type = "password";
    btn.textContent = "Show";
  }
}

async function verifyApiKey() {
  const input = document.getElementById("apiKeyInput");
  const btn = document.getElementById("verifyKeyBtn");
  const badge = document.getElementById("apiStatusBadge");
  const alertBox = document.getElementById("apiAlertBox");
  const key = input.value.trim();
  if (!key) {
    showAlert(alertBox, "error", "Please paste your Gemini API key first.");
    return;
  }
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner"></span> Verifying...`;
  alertBox.innerHTML = "";
  try {
    const res = await fetch(`${API_BASE}/api/verify-key`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ api_key: key }),
    });
    const data = await res.json();
    if (data.success) {
      apiKey = key;
      apiVerified = true;
      badge.textContent = "Verified";
      badge.className = "api-status-badge verified";
      showAlert(alertBox, "success", "API key verified successfully. You can now run analysis.");
      document.getElementById("processExcelBtn").disabled = !selectedFile;
    } else {
      apiVerified = false;
      badge.textContent = "Invalid";
      badge.className = "api-status-badge unverified";
      showAlert(alertBox, "error", data.message || "API key verification failed.");
    }
  } catch {
    showAlert(alertBox, "error", "Could not connect to server. Make sure server.py is running.");
  }
  btn.disabled = false;
  btn.textContent = "Verify Key";
}

function showAlert(container, type, message) {
  container.innerHTML = `
    <div class="alert alert-${type}">
      <div class="alert-dot"></div>
      <span>${message}</span>
    </div>`;
}

function setupDragDrop() {
  const zone = document.getElementById("uploadZone");
  if (!zone) return;
  zone.addEventListener("dragover", (e) => { e.preventDefault(); zone.classList.add("drag-over"); });
  zone.addEventListener("dragleave", () => { zone.classList.remove("drag-over"); });
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("drag-over");
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith(".xlsx") || file.name.endsWith(".xls"))) {
      setSelectedFile(file);
    } else {
      alert("Please upload an Excel file (.xlsx or .xls)");
    }
  });
}

function handleFileSelect(event) {
  const file = event.target.files[0];
  if (file) setSelectedFile(file);
}

function setSelectedFile(file) {
  selectedFile = file;
  const infoEl = document.getElementById("fileSelectedInfo");
  const sizeKB = (file.size / 1024).toFixed(1);
  infoEl.classList.remove("hidden");
  infoEl.innerHTML = `
    <div class="file-selected">
      <div>
        <div class="file-name">${file.name}</div>
        <div class="file-size">${sizeKB} KB</div>
      </div>
      <span class="badge badge-pass">Ready</span>
    </div>`;
  if (apiVerified) document.getElementById("processExcelBtn").disabled = false;
}

async function downloadTemplate() {
  try {
    const res = await fetch(`${API_BASE}/api/download-template`);
    if (!res.ok) throw new Error("Failed");
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "CatalogAI_Template.xlsx";
    a.click();
    URL.revokeObjectURL(url);
  } catch {
    alert("Could not download template. Make sure the server is running.");
  }
}

async function processExcel() {
  if (!apiVerified) { alert("Please verify your API key first."); return; }
  if (!selectedFile) { alert("Please select an Excel file first."); return; }

  document.getElementById("progressCard").classList.remove("hidden");
  document.getElementById("excelResultsSection").classList.add("hidden");
  document.getElementById("processExcelBtn").disabled = true;
  animateProgress();

  const formData = new FormData();
  formData.append("file", selectedFile);
  formData.append("api_key", apiKey);

  try {
    const res = await fetch(`${API_BASE}/api/process-excel`, { method: "POST", body: formData });
    const data = await res.json();
    setProgressComplete();
    if (data.success) {
      excelResults = data.results;
      setTimeout(() => {
        document.getElementById("progressCard").classList.add("hidden");
        renderExcelResults(data.results);
      }, 600);
    } else {
      document.getElementById("progressCard").classList.add("hidden");
      alert("Error: " + data.message);
    }
  } catch (err) {
    document.getElementById("progressCard").classList.add("hidden");
    alert("Server error. Make sure server.py is running.");
  }
  document.getElementById("processExcelBtn").disabled = false;
}

let progressInterval = null;

function animateProgress() {
  const steps = ["pStep1", "pStep2", "pStep3", "pStep4", "pStep5"];
  const labels = ["Reading Excel file...", "Sending to AI for analysis...", "Detecting conflicts and enriching data...", "Generating readiness scores...", "Preparing results..."];
  const percents = [10, 30, 60, 85, 95];
  let cs = 0;
  updateProgress(percents[0], labels[0]);
  markStep(steps[0], "active");
  progressInterval = setInterval(() => {
    if (cs < steps.length - 1) {
      markStep(steps[cs], "done");
      cs++;
      updateProgress(percents[cs], labels[cs]);
      markStep(steps[cs], "active");
    }
  }, 2000);
}

function updateProgress(p, l) {
  document.getElementById("progressBarFill").style.width = p + "%";
  document.getElementById("progressPercent").textContent = p + "%";
  document.getElementById("progressLabelText").textContent = l;
}

function markStep(id, state) {
  const el = document.getElementById(id);
  el.classList.remove("active", "done");
  el.classList.add(state);
}

function setProgressComplete() {
  clearInterval(progressInterval);
  updateProgress(100, "Analysis complete.");
  ["pStep1", "pStep2", "pStep3", "pStep4", "pStep5"].forEach(id => markStep(id, "done"));
}

function renderExcelResults(results) {
  const section = document.getElementById("excelResultsSection");
  section.classList.remove("hidden");

  const avgScore = results.length
    ? Math.round(results.reduce((s, r) => s + (r.readiness_score?.total || 0), 0) / results.length)
    : 0;
  const totalConflicts = results.reduce((s, r) => s + (r.conflicts?.length || 0), 0);
  const humanCheckCount = results.reduce((s, r) => {
    const fields = r.fields || {};
    return s + Object.values(fields).filter(f => f && f.human_check).length;
  }, 0);

  document.getElementById("summaryStats").innerHTML = `
    <div class="stat-box"><div class="stat-number">${results.length}</div><div class="stat-desc">Products Processed</div></div>
    <div class="stat-box"><div class="stat-number">${avgScore}%</div><div class="stat-desc">Avg Readiness</div></div>
    <div class="stat-box"><div class="stat-number">${totalConflicts}</div><div class="stat-desc">Conflicts Found</div></div>
    <div class="stat-box"><div class="stat-number">${humanCheckCount}</div><div class="stat-desc">Human Checks</div></div>`;

  renderMainResultTable(results);

  const nav = document.getElementById("productNav");
  nav.innerHTML = "";
  results.forEach((r, i) => {
    const name = r.product_name || r.original_data?.['Product Name'] || `Product ${i + 1}`;
    const btn = document.createElement("button");
    btn.className = "product-nav-btn" + (i === 0 ? " active" : "");
    btn.textContent = name.length > 22 ? name.substring(0, 22) + "..." : name;
    btn.onclick = () => {
      document.querySelectorAll(".product-nav-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      currentProductIndex = i;
      renderProductDetail(results[i]);
    };
    nav.appendChild(btn);
  });

  if (results.length > 0) renderProductDetail(results[0]);
}

function renderMainResultTable(results) {
  const container = document.getElementById("mainResultTable");
  if (!container) return;

  // Helper to get value from puraana format
  function getVal(r, key) {
    // Try enriched_fields first
    const enriched = r.enriched_fields || {};
    if (enriched[key]) {
      if (typeof enriched[key] === 'object') return enriched[key].value || "—";
      return enriched[key];
    }
    // Try original data (case insensitive)
    const orig = r.original_data || {};
    for (const k in orig) {
      if (k.toLowerCase().includes(key.toLowerCase()) || key.toLowerCase().includes(k.toLowerCase())) {
        return orig[k];
      }
    }
    return "—";
  }

  function getConf(r, key) {
    const enriched = r.enriched_fields || {};
    if (enriched[key] && typeof enriched[key] === 'object') {
      return enriched[key].confidence || 0;
    }
    // If from original data, it's 100% confident
    const orig = r.original_data || {};
    for (const k in orig) {
      if (k.toLowerCase().includes(key.toLowerCase()) || key.toLowerCase().includes(k.toLowerCase())) {
        return 100;
      }
    }
    return 0;
  }

  function isHumanCheck(r, key) {
    const enriched = r.enriched_fields || {};
    if (enriched[key] && typeof enriched[key] === 'object') {
      return enriched[key].human_check || false;
    }
    return false;
  }

  function getSource(r, key) {
    const enriched = r.enriched_fields || {};
    if (enriched[key] && typeof enriched[key] === 'object') {
      return enriched[key].source || "";
    }
    const orig = r.original_data || {};
    for (const k in orig) {
      if (k.toLowerCase().includes(key.toLowerCase()) || key.toLowerCase().includes(k.toLowerCase())) {
        return "Provided by User";
      }
    }
    return "";
  }

  const headers = [
    "S.No", "Product Name", "Product Type", "Category",
    "Score", "Grade", "Conflicts", "Human Check", "Status",
    ...FIELD_KEYS
  ];

  const headerRow = headers.map(h => `<th style="white-space:nowrap;">${h}</th>`).join("");

  const bodyRows = results.map((r, idx) => {
    if (r.error) {
      return `<tr><td>${idx + 1}</td><td colspan="${headers.length - 1}" style="color:var(--accent-danger);">Error: ${r.error}</td></tr>`;
    }

    const score = r.readiness_score || {};
    const conflicts = r.conflicts || [];

    // Count human checks
    const enriched = r.enriched_fields || {};
    let humanChecks = 0;
    for (const k in enriched) {
      if (enriched[k] && typeof enriched[k] === 'object' && enriched[k].human_check) {
        humanChecks++;
      }
    }

    let status = "Ready";
    let statusClass = "badge-pass";
    if (conflicts.length > 0) { status = "Has Conflicts"; statusClass = "badge-fail"; }
    else if (humanChecks > 0) { status = "Needs Review"; statusClass = "badge-warn"; }
    else if ((score.total || 0) < 50) { status = "Poor Quality"; statusClass = "badge-fail"; }

    const grade = score.grade || "N/A";
    const gradeClass = gradeToClass(grade);

    // Get product name from multiple possible sources
    const productName = r.product_name || getVal(r, 'product_name') || r.original_data?.['Product Name'] || "—";
    const productType = r.product_type || getVal(r, 'product_type') || r.original_data?.['Product Type'] || "—";
    const category = getVal(r, 'category') || "—";

    const basicCells = [
      `<td>${idx + 1}</td>`,
      `<td style="font-weight:600;color:var(--text-primary);">${productName}</td>`,
      `<td>${productType}</td>`,
      `<td>${category}</td>`,
      `<td style="font-family:var(--font-mono);color:var(--accent-primary);font-weight:600;">${score.total || 0}%</td>`,
      `<td><span class="grade-badge ${gradeClass}" style="font-size:11px;padding:2px 8px;">${grade}</span></td>`,
      `<td>${conflicts.length > 0 ? `<span class="badge badge-fail">${conflicts.length}</span>` : `<span class="badge badge-pass">0</span>`}</td>`,
      `<td>${humanChecks > 0 ? `<span class="badge badge-warn">${humanChecks}</span>` : `<span class="badge badge-pass">0</span>`}</td>`,
      `<td><span class="badge ${statusClass}">${status}</span></td>`,
    ].join("");

    const fieldCells = FIELD_KEYS.map(key => {
      const val = getVal(r, key);
      const conf = getConf(r, key);
      const hc = isHumanCheck(r, key);
      const src = getSource(r, key);

      let cellStyle = "";
      let icon = "";

      if (val === "—" || val === "") {
        cellStyle = "color:var(--text-muted);";
      } else if (hc) {
        cellStyle = "background:rgba(255,217,61,0.08);color:#ffe080;";
        icon = ' <span style="font-size:10px;">!</span>';
      } else if (conf >= 90) {
        cellStyle = "color:var(--text-primary);";
      } else if (conf >= 60) {
        cellStyle = "color:#a09aff;";
      } else if (conf > 0) {
        cellStyle = "color:#ff9a9a;";
      } else {
        cellStyle = "color:var(--text-muted);";
      }

      const tooltip = src ? `title="Confidence: ${conf}% | Source: ${src}"` : "";

      return `<td style="${cellStyle}font-family:var(--font-mono);font-size:12px;" ${tooltip}>${val}${icon}</td>`;
    }).join("");

    return `<tr>${basicCells}${fieldCells}</tr>`;
  }).join("");

  container.innerHTML = `
    <div style="overflow-x:auto; max-height:500px; overflow-y:auto; border:1px solid var(--border-color); border-radius:var(--radius-md); margin-bottom:16px;">
      <table class="data-table" style="min-width:2400px;">
        <thead style="position:sticky;top:0;z-index:10;background:var(--bg-primary);">
          <tr>${headerRow}</tr>
        </thead>
        <tbody>${bodyRows}</tbody>
      </table>
    </div>
    <div style="display:flex;gap:14px;margin-bottom:8px;font-size:11px;color:var(--text-muted);flex-wrap:wrap;">
      <div><span style="display:inline-block;width:10px;height:10px;background:#f0f0f5;border-radius:2px;margin-right:4px;vertical-align:middle;"></span> High Confidence (90%+)</div>
      <div><span style="display:inline-block;width:10px;height:10px;background:#a09aff;border-radius:2px;margin-right:4px;vertical-align:middle;"></span> Medium (60-89%)</div>
      <div><span style="display:inline-block;width:10px;height:10px;background:#ff9a9a;border-radius:2px;margin-right:4px;vertical-align:middle;"></span> Low (below 60%)</div>
      <div><span style="display:inline-block;width:10px;height:10px;background:#ffe080;border-radius:2px;margin-right:4px;vertical-align:middle;"></span> Needs Human Check</div>
    </div>
    <div style="font-size:11px;color:var(--text-muted);margin-bottom:16px;">Tip: Hover over any cell to see confidence percentage and source.</div>`;
}

function renderProductDetail(result) {
  const container = document.getElementById("productDetailView");
  if (!container) return;

  if (result.error) {
    container.innerHTML = `<div class="alert alert-error"><div class="alert-dot"></div><span>Error: ${result.error}</span></div>`;
    return;
  }

  const score = result.readiness_score || {};
  const grade = score.grade || "N/A";
  const enriched = result.enriched_fields || {};
  const orig = result.original_data || {};
  const conflicts = result.conflicts || [];
  const crossVerif = result.cross_verification || [];
  const missing = result.missing_fields || [];

  const productName = result.product_name || orig['Product Name'] || "Unknown Product";
  const productType = result.product_type || orig['Product Type'] || "";

  // Build all fields table combining original + enriched
  const allFieldsMap = {};

  // Add original data
  for (const key in orig) {
    allFieldsMap[key] = {
      value: orig[key],
      confidence: 100,
      source: "Provided by User",
      human_check: false,
      human_check_reason: ""
    };
  }

  // Add/override with enriched data
  for (const key in enriched) {
    const details = enriched[key];
    if (typeof details === 'object') {
      allFieldsMap[key] = {
        value: details.value || "—",
        confidence: details.confidence || 0,
        source: details.source || "AI Enriched",
        human_check: details.human_check || false,
        human_check_reason: details.human_check_reason || ""
      };
    }
  }

  const fieldRows = Object.entries(allFieldsMap).map(([key, fd]) => {
    const conf = fd.confidence || 0;
    const confClass = conf >= 75 ? "conf-high" : conf >= 50 ? "conf-mid" : "conf-low";
    const statusBadge = fd.human_check
      ? `<span class="badge badge-warn">Human Check</span>`
      : `<span class="badge badge-pass">OK</span>`;

    return `<tr>
      <td style="font-weight:600;">${key}</td>
      <td style="font-family:var(--font-mono);font-size:12px;">${fd.value || "—"}</td>
      <td><div class="confidence-bar-wrap"><div class="confidence-bar"><div class="confidence-bar-inner ${confClass}" style="width:${conf}%"></div></div><span class="confidence-num">${conf}%</span></div></td>
      <td><span class="badge badge-neutral">${fd.source || "—"}</span></td>
      <td>${statusBadge}</td>
    </tr>`;
  }).join("");

  let conflictHTML = "";
  if (conflicts.length === 0) {
    conflictHTML = `<div class="alert alert-success"><div class="alert-dot"></div><span>No conflicts detected.</span></div>`;
  } else {
    const cRows = conflicts.map(c => {
      const sc = c.severity === "HIGH" ? "badge-fail" : c.severity === "MEDIUM" ? "badge-warn" : "badge-neutral";
      return `<tr>
        <td>${c.field || "—"}</td>
        <td>${c.issue || "—"}</td>
        <td><span class="badge ${sc}">${c.severity || "—"}</span></td>
        <td style="font-size:12px;color:var(--text-muted);">${c.recommendation || "—"}</td>
      </tr>`;
    }).join("");
    conflictHTML = `<div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>Field</th><th>Issue</th><th>Severity</th><th>Recommendation</th></tr></thead><tbody>${cRows}</tbody></table></div>`;
  }

  let crossHTML = "";
  if (crossVerif.length > 0) {
    const vRows = crossVerif.map(v => {
      const sc = v.status === "PASS" ? "badge-pass" : v.status === "FAIL" ? "badge-fail" : "badge-neutral";
      return `<tr>
        <td>${v.check || "—"}</td>
        <td style="font-family:var(--font-mono);font-size:12px;">${v.formula || "—"}</td>
        <td><span class="badge ${sc}">${v.status || "—"}</span></td>
        <td style="font-size:12px;color:var(--text-muted);">${v.details || "—"}</td>
      </tr>`;
    }).join("");
    crossHTML = `<div class="section-label">Cross Verification</div><div style="overflow-x:auto;margin-bottom:24px;"><table class="data-table"><thead><tr><th>Check</th><th>Formula</th><th>Status</th><th>Details</th></tr></thead><tbody>${vRows}</tbody></table></div>`;
  }

  let missingHTML = "";
  if (missing.length > 0) {
    const mRows = missing.map(m => `<tr>
      <td style="font-weight:600;">${m.field || "—"}</td>
      <td style="font-size:12px;color:var(--text-muted);">${m.reason || "—"}</td>
      <td style="font-family:var(--font-mono);font-size:12px;">${m.suggested_value || "Unknown"}</td>
    </tr>`).join("");
    missingHTML = `<div class="section-label">Missing Fields</div><div style="overflow-x:auto;margin-bottom:24px;"><table class="data-table"><thead><tr><th>Field</th><th>Why Important</th><th>Suggested Value</th></tr></thead><tbody>${mRows}</tbody></table></div>`;
  }

  container.innerHTML = `
    <div class="fade-in">
      <div class="result-header">
        <div class="result-product-info">
          <h2>${productName}</h2>
          <p>${productType} — ${result.summary || ""}</p>
        </div>
        <span class="grade-badge ${gradeToClass(grade)}">${grade}</span>
      </div>

      <div class="section-label">Readiness Score</div>
      <div class="score-grid" style="margin-bottom:24px;">
        <div class="score-card"><div class="score-value">${score.total || 0}%</div><div class="score-label">Overall</div></div>
        <div class="score-card"><div class="score-value">${score.completeness || 0}%</div><div class="score-label">Completeness</div></div>
        <div class="score-card"><div class="score-value">${score.accuracy || 0}%</div><div class="score-label">Accuracy</div></div>
        <div class="score-card"><div class="score-value">${score.verification || 0}%</div><div class="score-label">Verification</div></div>
      </div>

      <div class="section-label">All Fields — Detailed View</div>
      <div style="overflow-x:auto;margin-bottom:24px;">
        <table class="data-table">
          <thead><tr><th>Field</th><th>Value</th><th>Confidence</th><th>Source</th><th>Status</th></tr></thead>
          <tbody>${fieldRows}</tbody>
        </table>
      </div>

      <div class="section-label">Conflicts</div>
      <div style="margin-bottom:24px;">${conflictHTML}</div>

      ${crossHTML}

      ${missingHTML}
    </div>`;
}

function gradeToClass(grade) {
  const map = { "A+": "grade-A-plus", "A": "grade-A", "B": "grade-B", "C": "grade-C", "D": "grade-D", "F": "grade-F" };
  return map[grade] || "grade-F";
}

async function analyzeSingle() {
  if (!apiVerified) {
    showAlert(document.getElementById("singleAlertBox"), "error", "Please verify your API key first.");
    return;
  }
  const productData = {};
  const fieldMap = {
    "Product Name": "s_name", "Brand": "s_brand", "Model Number": "s_model",
    "Product Type": "s_type", "Power (W)": "s_power", "Voltage (V)": "s_voltage",
    "Current (A)": "s_current", "Frequency (Hz)": "s_freq", "RPM": "s_rpm",
    "Efficiency (%)": "s_eff", "IP Rating": "s_ip", "Weight (kg)": "s_weight", "Notes": "s_notes",
  };
  Object.entries(fieldMap).forEach(([label, id]) => {
    const val = document.getElementById(id)?.value?.trim();
    if (val) productData[label] = val;
  });
  if (Object.keys(productData).length === 0) {
    showAlert(document.getElementById("singleAlertBox"), "error", "Please fill in at least one field.");
    return;
  }

  const btn = document.getElementById("analyzeSingleBtn");
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner"></span> Analyzing...`;
  document.getElementById("singleAlertBox").innerHTML = "";
  document.getElementById("singleResultSection").classList.add("hidden");

  try {
    const res = await fetch(`${API_BASE}/api/analyze-single`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ api_key: apiKey, product_data: productData }),
    });
    const data = await res.json();
    if (data.success) {
      const section = document.getElementById("singleResultSection");
      const card = document.getElementById("singleResultCard");
      section.classList.remove("hidden");

      const r = data.data;
      excelResults = [r];
      card.innerHTML = `<div class="card-title">Analysis Result</div><div id="singleDetailView"></div>`;

      const originalDetailView = document.getElementById("productDetailView");
      const backupId = originalDetailView ? originalDetailView.id : null;

      // Render into singleDetailView using same function
      const tempDiv = document.getElementById("singleDetailView");
      tempDiv.id = "productDetailView";
      renderProductDetail(r);
      tempDiv.id = "singleDetailView";

      card.scrollIntoView({ behavior: "smooth", block: "start" });
    } else {
      showAlert(document.getElementById("singleAlertBox"), "error", data.message || "Analysis failed.");
    }
  } catch {
    showAlert(document.getElementById("singleAlertBox"), "error", "Server error. Make sure server.py is running.");
  }
  btn.disabled = false;
  btn.innerHTML = "Analyze Product";
}

async function downloadResults() {
  if (!excelResults.length) { alert("No results to download."); return; }
  try {
    const res = await fetch(`${API_BASE}/api/download-results`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ results: excelResults }),
    });
    if (!res.ok) throw new Error("Failed");
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "CatalogAI_Results.xlsx";
    a.click();
    URL.revokeObjectURL(url);
  } catch {
    alert("Could not download results.");
  }
}

function resetExcel() {
  selectedFile = null;
  excelResults = [];
  currentProductIndex = 0;
  document.getElementById("excelFileInput").value = "";
  document.getElementById("fileSelectedInfo").classList.add("hidden");
  document.getElementById("fileSelectedInfo").innerHTML = "";
  document.getElementById("excelResultsSection").classList.add("hidden");
  document.getElementById("progressCard").classList.add("hidden");
  document.getElementById("processExcelBtn").disabled = true;
  updateProgress(0, "Initializing...");
  ["pStep1", "pStep2", "pStep3", "pStep4", "pStep5"].forEach(id => {
    document.getElementById(id).classList.remove("active", "done");
  });
}
// ─── CHATBOT ─────────────────────────────────────────────
let chatOpen = false;
let chatHistory = [];

function toggleChat() {
  const win = document.getElementById("chatWindow");
  const btn = document.getElementById("chatToggleBtn");
  const iconOpen = document.getElementById("chatIconOpen");
  const iconClose = document.getElementById("chatIconClose");

  chatOpen = !chatOpen;

  if (chatOpen) {
    win.classList.add("open");
    btn.classList.add("active");
    iconOpen.style.display = "none";
    iconClose.style.display = "block";
    setTimeout(() => document.getElementById("chatInput").focus(), 300);
  } else {
    win.classList.remove("open");
    btn.classList.remove("active");
    iconOpen.style.display = "block";
    iconClose.style.display = "none";
  }
}

function sendSuggestion(text) {
  document.getElementById("chatInput").value = text;
  sendChatMessage();
}

async function sendChatMessage() {
  const input = document.getElementById("chatInput");
  const sendBtn = document.getElementById("chatSendBtn");
  const body = document.getElementById("chatBody");
  const suggestions = document.getElementById("chatSuggestions");
  const message = input.value.trim();

  if (!message) return;

  if (!apiVerified) {
    addChatMessage("bot", "Please verify your Gemini API key first at the top of the page to use the assistant.");
    return;
  }

  // Hide suggestions after first message
  suggestions.style.display = "none";

  // Add user message
  addChatMessage("user", message);
  chatHistory.push({ role: "user", content: message });
  input.value = "";
  sendBtn.disabled = true;

  // Show typing indicator
  const typingEl = document.createElement("div");
  typingEl.className = "chat-typing";
  typingEl.id = "chatTyping";
  typingEl.innerHTML = "<span></span><span></span><span></span>";
  body.appendChild(typingEl);
  body.scrollTop = body.scrollHeight;

  try {
    const res = await fetch(`${API_BASE}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        api_key: apiKey,
        message: message,
        history: chatHistory,
      }),
    });
    const data = await res.json();

    // Remove typing
    const typing = document.getElementById("chatTyping");
    if (typing) typing.remove();

    if (data.success) {
      addChatMessage("bot", data.reply);
      chatHistory.push({ role: "bot", content: data.reply });
    } else {
      addChatMessage("bot", "Sorry, I couldn't process that. " + (data.message || ""));
    }
  } catch {
    const typing = document.getElementById("chatTyping");
    if (typing) typing.remove();
    addChatMessage("bot", "Connection error. Make sure the server is running.");
  }

  sendBtn.disabled = false;
  input.focus();
}

function addChatMessage(role, text) {
  const body = document.getElementById("chatBody");
  const msg = document.createElement("div");
  msg.className = "chat-msg " + role;
  // Convert line breaks to <br>
  msg.innerHTML = text.replace(/\n/g, "<br>");
  body.appendChild(msg);
  body.scrollTop = body.scrollHeight;
}