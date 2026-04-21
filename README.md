# Modbus signal matcher (web)

**Repository:** [github.com/Hbradroc/Access5modbusSignalgenerator](https://github.com/Hbradroc/Access5modbusSignalgenerator)

After you turn on GitHub Pages (see below), the site will be at:

**`https://hbradroc.github.io/Access5modbusSignalgenerator/`**

---

UI matches the style and flow of the [Magic JSON Converter](https://hbradroc.github.io/dvf2Json/) (centered card, Space Grotesk, one main action, log area). Still fully static—no backend.

Static single-page app: loads **`data/Access5ModbusSignals.xls`** by default, you upload your Access export **`.json` or `.cas`** (same structure as the PCB JSON: top-level `"Variables"` array). It outputs **exact + normalized** matches in a table and as **CSV download** (same logic as `match_modbus_equivalents.py`).

## Files

| File | Role |
|------|------|
| `index.html` | Page |
| `app.js` | Matching + XLS/JSON parsing (SheetJS from CDN) |
| `styles.css` | Layout |
| `data/Access5ModbusSignals.xls` | Default Modbus signal list (copy from `Modbus Research/Access5ModbusSignals.xls` if missing) |

## `.cas` files

The app treats **`.cas` as UTF-8 JSON** with the same shape as your `.json` export (`FileType`, `Variables`, …). If your tool exports binary or non-JSON `.cas`, upload the **`.json`** export instead or rename a valid JSON export to `.cas` for testing.

## Host on GitHub Pages (one-time setup)

1. Open the repo on GitHub: [Access5modbusSignalgenerator](https://github.com/Hbradroc/Access5modbusSignalgenerator).
2. Go to **Settings → Pages** (left sidebar).
3. Under **Build and deployment → Source**, choose **Deploy from a branch**.
4. **Branch:** `main`, folder **`/ (root)`**, then **Save**.
5. Wait 1–2 minutes; refresh **Pages** until it shows “Your site is live at …”.

This repo already includes **`.nojekyll`** at the root so GitHub Pages does not run Jekyll on your static files.

## Local use

Serving over `http://` is required so the default **`fetch(data/Access5ModbusSignals.xls)`** works. Opening `index.html` as `file://` often blocks `fetch`; in that case use **“Upload Access5ModbusSignals.xls manually”** or run a local server, for example:

```bash
cd modbus-signal-matcher-web
python -m http.server 8080
```

Then open `http://localhost:8080/`.

## Large Excel file

`Access5ModbusSignals.xls` is ~3.5 MB. GitHub accepts it; if you prefer not to commit it, omit `data/` from the repo and always upload the XLS in the web UI.
