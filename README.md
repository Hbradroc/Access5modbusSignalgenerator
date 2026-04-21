# Modbus signal matcher (web)

**Repository:** [github.com/Hbradroc/Access5modbusSignalgenerator](https://github.com/Hbradroc/Access5modbusSignalgenerator)

**Live site (after Pages is enabled):** `https://hbradroc.github.io/Access5modbusSignalgenerator/`

---

UI matches the [Magic JSON Converter](https://hbradroc.github.io/dvf2Json/) (centered card, Space Grotesk, one action, log). **Fully static** — no server, no SheetJS.

## What you upload

Only your Access **`.json` or `.cas`** export (UTF-8 JSON with a top-level **`Variables`** array).

## Where the signal list lives

The Modbus signal table is **not** read from `.xls` in the browser. It is shipped as:

| File | Role |
|------|------|
| `data/modbus_signals.json` | All rows from `Access5ModbusSignals.xls` sheet **List**, generated offline (~500 KB) |

To refresh after you change the Excel file in `Modbus Research/`:

```bash
cd modbus-signal-matcher-web
python tools/build_signals_json.py
```

(Uses `../Access5ModbusSignals.xls` by default.) Commit the updated `data/modbus_signals.json`.

## Other files

| File | Role |
|------|------|
| `index.html` | Page |
| `app.js` | Fetch bundled JSON + match + CSV download |
| `styles.css` | Layout |
| `tools/build_signals_json.py` | XLS → JSON exporter (needs `pip install xlrd`) |

## `.cas` files

Treated as **UTF-8 JSON** with the same shape as `.json`. If yours is not JSON, use the `.json` export.

## GitHub Pages

**Settings → Pages →** Deploy from branch **`main`**, folder **`/ (root)`**.  
`.nojekyll` is included so Jekyll does not process the site.

## Local preview

```bash
cd modbus-signal-matcher-web
python -m http.server 8080
```

Open `http://localhost:8080/` (so `fetch` can load `data/modbus_signals.json`).
