/**
 * Modbus signal matcher — static page, no server.
 * Bundled signal list: data/modbus_signals.json (from Access5ModbusSignals.xls).
 * User upload: Access .json / .cas with Variables array only.
 */

const SIGNALS_JSON = "data/modbus_signals.json";

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

const el = {
  log: document.getElementById("log"),
  userFile: document.getElementById("userFile"),
  actionBtn: document.getElementById("actionBtn"),
  tbody: document.querySelector("#resultTable tbody"),
  preview: document.getElementById("preview"),
};

let signalRows = null;
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
  el.actionBtn.disabled = !(signalRows && userIdx);
}

async function loadBundledSignals() {
  setLog("Loading signal list…");
  try {
    const res = await fetch(SIGNALS_JSON);
    if (!res.ok) throw new Error(res.statusText);
    const data = await res.json();
    if (!data || !Array.isArray(data.rows)) {
      throw new Error("Invalid modbus_signals.json (expected .rows array).");
    }
    signalRows = data.rows;
    const src = data.source || "modbus_signals.json";
    setLog(
      `Signal list ready: ${signalRows.length} rows (from ${src}).\n` +
        `Upload your Access .json or .cas, then click “Match and download CSV”.`
    );
  } catch (e) {
    signalRows = null;
    setLog(
      `Could not load ${SIGNALS_JSON} (${e.message}).\n` +
        `Rebuild with: python tools/build_signals_json.py`
    );
  }
  updateActionEnabled();
}

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
  if (!signalRows || !userIdx) return;

  lastMatches = buildMatches(signalRows, userIdx);
  const exact = lastMatches.filter((r) => r.match_type === "exact").length;
  const norm = lastMatches.filter((r) => r.match_type === "normalized").length;

  appendLog(
    `—\nMatched ${lastMatches.length} rows (exact ${exact}, normalized ${norm}). ` +
      `Signal list size: ${signalRows.length}. Unmatched rows omitted from CSV.`
  );

  if (lastMatches.length === 0) {
    appendLog("No CSV written (zero matches).");
    renderTable([]);
    el.preview.open = true;
    return;
  }

  downloadCsv(lastMatches);
  appendLog("Downloaded modbus_equivalent_mapping.csv");
  renderTable(lastMatches);
  el.preview.open = true;
});

loadBundledSignals();
