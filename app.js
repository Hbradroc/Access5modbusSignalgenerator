/**
 * Modbus signal matcher — browser side.
 * Matches Access JSON/.cas Variables export to Access5ModbusSignals.xls sheet "List".
 */

/* global XLSX */

const DEFAULT_XLS = "data/Access5ModbusSignals.xls";

function normalizeName(name) {
  return String(name).replace(/[^A-Za-z0-9]/g, "").toLowerCase();
}

function parseAccessJsonObject(data) {
  const variables = Array.isArray(data.Variables) ? data.Variables : [];
  const baseToRaw = {};
  const baseToDesc = {};

  for (const item of variables) {
    if (!item || typeof item !== "object" || Object.keys(item).length !== 1) continue;
    const key = Object.keys(item)[0];
    const value = item[key];

    if (key.endsWith("_PublicDescription")) {
      const base = key.slice(0, -"_PublicDescription".length);
      if (typeof value === "string") baseToDesc[base] = value;
      continue;
    }

    const base = key.split(".", 1)[0];
    if (!(base in baseToRaw) && !key.includes(".")) {
      baseToRaw[base] = value;
    }
  }

  const exact = {};
  const normalized = {};
  for (const k of Object.keys(baseToRaw)) {
    exact[k] = k;
    normalized[normalizeName(k)] = k;
  }

  return { exact, normalized, baseToRaw, baseToDesc };
}

async function parseUserFile(file) {
  const lower = file.name.toLowerCase();
  const buf = await file.arrayBuffer();

  if (lower.endsWith(".json") || lower.endsWith(".cas")) {
    const text = new TextDecoder("utf-8", { fatal: false }).decode(buf);
    let data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      throw new Error(
        `${file.name}: not valid UTF-8 JSON. Export must be Access JSON with "Variables" (same as .json export).`
      );
    }
    if (!data || typeof data !== "object" || !Array.isArray(data.Variables)) {
      throw new Error('Expected top-level JSON with a "Variables" array (Access PCB export).');
    }
    return parseAccessJsonObject(data);
  }

  throw new Error("Unsupported file type. Use .json or .cas");
}

function parseXlsArrayBuffer(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  const sheetName =
    wb.SheetNames.find((n) => /^list$/i.test(String(n).trim())) || wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  if (!sheet) throw new Error("No sheets found in Excel file.");

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
  });

  const headerRowIdx = 1;
  if (rows.length <= headerRowIdx + 1) {
    throw new Error("Spreadsheet has too few rows (expected headers on row 2).");
  }

  const headerCells = rows[headerRowIdx];
  const headers = headerCells.map((h, i) => (String(h).trim() || `col_${i}`));

  const out = [];
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const rowArr = rows[r];
    if (!rowArr || !String(rowArr[0] || "").trim()) continue;

    const row = {};
    for (let c = 0; c < headers.length; c++) {
      let v = rowArr[c];
      if (typeof v === "number" && Number.isInteger(v)) {
        /* keep */
      } else if (typeof v === "string" && v.trim() !== "" && !Number.isNaN(Number(v))) {
        const n = Number(v);
        if (String(n) === v.trim()) v = n;
      }
      row[headers[c]] = v ?? "";
    }
    out.push(row);
  }

  return out;
}

function buildMatches(xlsRows, idx) {
  const { exact, normalized, baseToRaw, baseToDesc } = idx;
  const matches = [];

  for (const row of xlsRows) {
    const signal = String(row["Signal name"] ?? "").trim();
    if (!signal) continue;

    let matchType = "none";
    let matchedName = "";

    if (exact[signal] !== undefined) {
      matchType = "exact";
      matchedName = exact[signal];
    } else {
      const ns = normalizeName(signal);
      if (normalized[ns] !== undefined) {
        matchType = "normalized";
        matchedName = normalized[ns];
      }
    }

    if (matchType === "none") continue;

    matches.push({
      match_type: matchType,
      signal_name_xls: signal,
      matched_json_signal: matchedName,
      json_value: baseToRaw[matchedName] ?? "",
      json_public_description: baseToDesc[matchedName] ?? "",
      // Excel List sheet columns B, F, J, K (1-based)
      column_B_EXOL_Type: row["EXOL Type"] ?? "",
      column_F_Bacnet: row.Bacnet ?? "",
      column_J_Description: row.Description ?? "",
      column_K_BACnet_Address: row["BACnet Address"] ?? "",
      exol_type: row["EXOL Type"] ?? "",
      modbus_type: row["Modbus Type"] ?? "",
      modbus_address: row["Modbus address"] ?? "",
      scale: row.Scale ?? "",
      bacnet: row.Bacnet ?? "",
      default_value: row["Default value"] ?? "",
    });
  }

  return matches;
}

function escCsv(v) {
  const s = v === null || v === undefined ? "" : String(v);
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function toCsv(rows) {
  if (rows.length === 0) return "";
  const keys = Object.keys(rows[0]);
  const lines = [
    keys.join(","),
    ...rows.map((r) => keys.map((k) => escCsv(r[k])).join(",")),
  ];
  return lines.join("\r\n");
}

// --- UI ---

const el = {
  xlsStatus: document.getElementById("xlsStatus"),
  userStatus: document.getElementById("userStatus"),
  xlsFile: document.getElementById("xlsFile"),
  userFile: document.getElementById("userFile"),
  runBtn: document.getElementById("runBtn"),
  downloadCsv: document.getElementById("downloadCsv"),
  stats: document.getElementById("stats"),
  tbody: document.querySelector("#resultTable tbody"),
  emptyHint: document.getElementById("emptyHint"),
};

let xlsRows = null;
let userIdx = null;
let lastMatches = [];

async function loadDefaultXls() {
  try {
    const res = await fetch(DEFAULT_XLS);
    if (!res.ok) throw new Error(res.statusText);
    const buf = await res.arrayBuffer();
    xlsRows = parseXlsArrayBuffer(buf);
    el.xlsStatus.textContent = `Ready: loaded ${xlsRows.length} Excel signals from ${DEFAULT_XLS}`;
    el.xlsStatus.classList.remove("muted");
  } catch (e) {
    el.xlsStatus.textContent = `Could not load default file (${e.message}). Upload Access5ModbusSignals.xls below.`;
    el.xlsStatus.classList.add("muted");
  }
  updateRunEnabled();
}

el.xlsFile.addEventListener("change", async (ev) => {
  const f = ev.target.files[0];
  if (!f) return;
  try {
    const buf = await f.arrayBuffer();
    xlsRows = parseXlsArrayBuffer(buf);
    el.xlsStatus.textContent = `Ready: ${xlsRows.length} signals from "${f.name}"`;
    el.xlsStatus.classList.remove("muted");
  } catch (err) {
    el.xlsStatus.textContent = `Error reading Excel: ${err.message}`;
    xlsRows = null;
  }
  updateRunEnabled();
});

el.userFile.addEventListener("change", async (ev) => {
  const f = ev.target.files[0];
  if (!f) {
    userIdx = null;
    el.userStatus.textContent = "No file selected.";
    updateRunEnabled();
    return;
  }
  try {
    userIdx = await parseUserFile(f);
    const n = Object.keys(userIdx.baseToRaw).length;
    el.userStatus.textContent = `Parsed "${f.name}": ${n} top-level variables.`;
    el.userStatus.classList.remove("muted");
  } catch (err) {
    userIdx = null;
    el.userStatus.textContent = err.message;
    el.userStatus.classList.add("muted");
  }
  updateRunEnabled();
});

function updateRunEnabled() {
  el.runBtn.disabled = !(xlsRows && userIdx);
  el.downloadCsv.disabled = lastMatches.length === 0;
}

el.runBtn.addEventListener("click", () => {
  if (!xlsRows || !userIdx) return;
  lastMatches = buildMatches(xlsRows, userIdx);
  renderTable(lastMatches);

  const exact = lastMatches.filter((r) => r.match_type === "exact").length;
  const norm = lastMatches.filter((r) => r.match_type === "normalized").length;
  el.stats.hidden = false;
  el.stats.innerHTML =
    `Matched <strong>${lastMatches.length}</strong> Excel rows ` +
    `(exact <strong>${exact}</strong>, normalized <strong>${norm}</strong>). ` +
    `Excel total rows: ${xlsRows.length}. Unmatched Excel rows omitted from table.`;

  el.emptyHint.hidden = lastMatches.length > 0;
  el.downloadCsv.disabled = lastMatches.length === 0;
  updateRunEnabled();
});

function renderTable(rows) {
  el.tbody.innerHTML = "";
  for (const r of rows) {
    const tr = document.createElement("tr");
    const badge =
      r.match_type === "exact"
        ? '<span class="badge exact">exact</span>'
        : '<span class="badge norm">normalized</span>';
    tr.innerHTML = `
      <td>${badge}</td>
      <td>${escapeHtml(r.signal_name_xls)}</td>
      <td>${escapeHtml(r.matched_json_signal)}</td>
      <td>${escapeHtml(formatVal(r.json_value))}</td>
      <td>${escapeHtml(String(r.json_public_description))}</td>
      <td>${escapeHtml(String(r.column_B_EXOL_Type))}</td>
      <td>${escapeHtml(String(r.column_F_Bacnet))}</td>
      <td>${escapeHtml(String(r.column_J_Description))}</td>
      <td>${escapeHtml(String(r.column_K_BACnet_Address))}</td>
      <td>${escapeHtml(String(r.modbus_type))}</td>
      <td>${escapeHtml(String(r.modbus_address))}</td>
      <td>${escapeHtml(String(r.scale))}</td>
      <td>${escapeHtml(String(r.bacnet))}</td>
    `;
    el.tbody.appendChild(tr);
  }
}

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s;
  return d.innerHTML;
}

function formatVal(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return JSON.stringify(v);
  return String(v);
}

el.downloadCsv.addEventListener("click", () => {
  if (lastMatches.length === 0) return;
  const csv = toCsv(lastMatches);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "modbus_equivalent_mapping.csv";
  a.click();
  URL.revokeObjectURL(a.href);
});

loadDefaultXls();
