# Modbus signal matcher (web)

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

## Host on GitHub Pages

1. Push this folder to your repository (or copy its contents into the repo’s **`docs/`** folder).
2. In the repo: **Settings → Pages → Build and deployment → Branch**, select the branch and **`/docs`** (if you used `docs/`) or **`/(root)`**.
3. Optional: add an empty **`.nojekyll`** file next to `index.html` so Jekyll does not touch the site (recommended if you deploy from `/docs`).

After deploy, open:

`https://<user>.github.io/<repo>/`  
(or `.../modbus-signal-matcher-web/` if you only uploaded this subfolder without making it the site root—then ensure `data/Access5ModbusSignals.xls` is under that path.)

## Local use

Serving over `http://` is required so the default **`fetch(data/Access5ModbusSignals.xls)`** works. Opening `index.html` as `file://` often blocks `fetch`; in that case use **“Upload Access5ModbusSignals.xls manually”** or run a local server, for example:

```bash
cd modbus-signal-matcher-web
python -m http.server 8080
```

Then open `http://localhost:8080/`.

## Large Excel file

`Access5ModbusSignals.xls` is ~3.5 MB. GitHub accepts it; if you prefer not to commit it, omit `data/` from the repo and always upload the XLS in the web UI.
