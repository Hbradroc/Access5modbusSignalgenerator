#!/usr/bin/env python3
"""
Export Access5ModbusSignals.xls sheet "List" to JSON for the static web app.
Run after updating the XLS in Modbus Research/:

  python tools/build_signals_json.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

try:
    import xlrd
except ImportError:
    print("pip install xlrd", file=sys.stderr)
    raise SystemExit(1)


def load_modbus_xls(path: Path, sheet_name: str = "List") -> list[dict]:
    wb = xlrd.open_workbook(str(path))
    sh = wb.sheet_by_name(sheet_name)
    header_row_idx = 1
    headers = [str(sh.cell_value(header_row_idx, c)).strip() for c in range(sh.ncols)]

    rows: list[dict] = []
    for r in range(header_row_idx + 1, sh.nrows):
        signal_name = str(sh.cell_value(r, 0)).strip()
        if not signal_name:
            continue
        row: dict = {}
        for c, header in enumerate(headers):
            key = header if header else f"col_{c}"
            value = sh.cell_value(r, c)
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            row[key] = value
        rows.append(row)
    return rows


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    default_xls = root.parent / "Access5ModbusSignals.xls"
    xls_path = Path(sys.argv[1]) if len(sys.argv) > 1 else default_xls
    out_path = root / "data" / "modbus_signals.json"

    if not xls_path.is_file():
        print(f"XLS not found: {xls_path}", file=sys.stderr)
        return 1

    rows = load_modbus_xls(xls_path)
    payload = {
        "version": 1,
        "source": xls_path.name,
        "sheet": "List",
        "row_count": len(rows),
        "rows": rows,
    }
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    print(f"Wrote {out_path} ({len(rows)} signals)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
