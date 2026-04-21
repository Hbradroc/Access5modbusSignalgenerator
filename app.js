/**
 * Modbus signal matcher — static page, no server.
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
        `${file.name}: not valid UTF-8 JSON. Use Access export with "Variables" array.`
      );
    }
    if (!data || typeof data !== "object" || !Array.isArray(data.Variables)) {
      throw new Error('JSON must include top-level "Variables" array.');
    }
    return parseAccessJsonObject(data);
  }

  throw new Error("Use .json or .cas");
}

function parseXlsArrayBuffer(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  const sheetName =
    wb.SheetNames.find((n) => /^list$/i.test(String(n).trim())) || wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  if (!sheet) throw new Error("No sheets in workbook.");

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
  });

  const headerRowIdx = 1;
  if (rows.length <= headerRowIdx + 1) {
    throw new Error("Spreadsheet too small (expected headers on row 2).");
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

function downloadCsv(rows) {
  const csv = toCsv(rows);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "modbus_equivalent_mapping.csv";
  a.click();
  URL.revokeObjectURL(a.href);
}

// --- UI ---

const el = {
  log: document.getElementById("log"),
  xlsFile: document.getElementById("xlsFile"),
  userFile: document.getElementById("userFile"),
  actionBtn: document.getElementById("actionBtn"),
  tbody: document.querySelector("#resultTable tbody"),
  preview: document.getElementById("preview"),
};

let xlsRows = null;
let userIdx = null;
let lastMatches = [];

function setLog(text) {
  el.log.textContent = text;
}

function appendLog(line) {
  el.log.textContent += (el.log.textContent ? "\n" : "") + line;
  el.log.scrollTop = el.log.scrollHeight;
}

function updateActionEnabled() {
  el.actionBtn.disabled = !(xlsRows && userIdx);
}

async function loadDefaultXls() {
  setLog("Loading default signal list…");
  try {
    const res = await fetch(DEFAULT_XLS);
    if (!res.ok) throw new Error(res.statusText);
    const buf = await res.arrayBuffer();
    xlsRows = parseXlsArrayBuffer(buf);
    setLog(
      `Default signal list ready: ${xlsRows.length} rows from Access5ModbusSignals.xls\n` +
        `Upload your .json or .cas, then click “Match and download CSV”.`
    );
  } catch (e) {
    xlsRows = null;
    setLog(
      `Could not load built-in spreadsheet (${e.message}).\n` +
        `Upload Access5ModbusSignals.xls in the optional field above (needed for file:// or missing data/).`
    );
  }
  updateActionEnabled();
}

el.xlsFile.addEventListener("change", async (ev) => {
  const f = ev.target.files[0];
  if (!f) {
    await loadDefaultXls();
    return;
  }
  try {
    const buf = await f.arrayBuffer();
    xlsRows = parseXlsArrayBuffer(buf);
    appendLog(`Using uploaded signal list: ${f.name} (${xlsRows.length} rows)`);
  } catch (err) {
    xlsRows = null;
    appendLog(`Excel error: ${err.message}`);
  }
  updateActionEnabled();
});

el.userFile.addEventListener("change", async (ev) => {
  const f = ev.target.files[0];
  if (!f) {
    userIdx = null;
    appendLog("Access export cleared.");
    updateActionEnabled();
    return;
  }
  try {
    userIdx = await parseUserFile(f);
    const n = Object.keys(userIdx.baseToRaw).length;
    appendLog(`Loaded ${f.name}: ${n} top-level variables.`);
  } catch (err) {
    userIdx = null;
    appendLog(`Error: ${err.message}`);
  }
  updateActionEnabled();
});

function renderTable(rows) {
  el.tbody.innerHTML = "";
  for (const r of rows) {
    const tr = document.createElement("tr");
    const badge =
      r.match_type === "exact"
        ? '<span class="badge exact">exact</span>'
        : '<span class="badge norm">norm</span>';
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

el.actionBtn.addEventListener("click", () => {
  if (!xlsRows || !userIdx) return;

  lastMatches = buildMatches(xlsRows, userIdx);
  const exact = lastMatches.filter((r) => r.match_type === "exact").length;
  const norm = lastMatches.filter((r) => r.match_type === "normalized").length;

  appendLog(
    `—\nMatched ${lastMatches.length} Excel rows (exact ${exact}, normalized ${norm}). ` +
      `Excel rows scanned: ${xlsRows.length}. Unmatched rows omitted from CSV.`
  );

  if (lastMatches.length === 0) {
    appendLog("No CSV written (zero matches). Check file / signal names.");
    renderTable([]);
    el.preview.open = true;
    return;
  }

  downloadCsv(lastMatches);
  appendLog("Downloaded modbus_equivalent_mapping.csv");
  renderTable(lastMatches);
  el.preview.open = true;
});

loadDefaultXls();
