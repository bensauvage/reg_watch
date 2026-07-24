#!/usr/bin/env python3
"""
build_site.py — reads lu_weekly_regwatch.xlsx and regenerates index.html
from template.html. Run automatically via LaunchAgent on Excel change, or manually:

    python3 build_site.py
"""

import hashlib
import json
import re
import sys
from datetime import datetime
from pathlib import Path

XLSX     = Path(__file__).parent / "lu_weekly_regwatch.xlsx"
TEMPLATE = Path(__file__).parent / "template.html"
OUTPUT   = Path(__file__).parent / "index.html"
SNAPSHOT = Path(__file__).parent / ".claude" / "last_snapshot.json"

# Column indices in "Reg Watch" sheet (0-based, row 0 = headers)
COL_ID      = 0
COL_TOPIC   = 1
COL_SUMMARY = 2
COL_IMPACT  = 3
COL_PUBDATE = 4
COL_ISSUER  = 5
COL_STATUS  = 6
COL_APPDATE = 7
COL_CPX_O   = 8
COL_CPX_IT  = 9
COL_LINK    = 10

DASHBOARD_SHEET = "Dashboard"
REGWATCH_SHEET  = "Reg Watch"


def install_openpyxl():
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])


def read_excel(path: Path) -> list[dict]:
    try:
        import openpyxl
    except ImportError:
        print("openpyxl not found — installing…")
        install_openpyxl()
        import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[REGWATCH_SHEET]
    rows = list(ws.iter_rows(values_only=True))

    items = []
    for row in rows[1:]:
        item_id = row[COL_ID]
        if not isinstance(item_id, (int, float)):
            continue
        item_id = int(item_id)

        def cell(idx):
            v = row[idx]
            if v is None:
                return ""
            if isinstance(v, datetime):
                return v.strftime("%-d %b %Y")
            return str(v).strip()

        def cpx(idx):
            v = row[idx]
            try:
                return int(float(v))
            except (TypeError, ValueError):
                return 1

        items.append({
            "id":      item_id,
            "topic":   cell(COL_TOPIC),
            "summary": cell(COL_SUMMARY),
            "impact":  cell(COL_IMPACT),
            "pubDate": cell(COL_PUBDATE),
            "issuer":  cell(COL_ISSUER),
            "status":  cell(COL_STATUS),
            "appDate": cell(COL_APPDATE),
            "cpxOrga": cpx(COL_CPX_O),
            "cpxIT":   cpx(COL_CPX_IT),
            "link":    cell(COL_LINK),
        })

    return items


def item_fingerprint(item: dict) -> str:
    """MD5 of all content fields — changes when anything in the row is edited."""
    fields = "|".join(str(item.get(k, "")) for k in
                      ["summary", "status", "appDate", "pubDate",
                       "issuer", "topic", "impact", "cpxOrga", "cpxIT", "link"])
    return hashlib.md5(fields.encode()).hexdigest()


def mark_new_items(items: list[dict]) -> list[dict]:
    """
    Compare items against the previous snapshot.
    - New item (id not in snapshot) → isNew: true
    - Changed item (fingerprint differs) → isNew: true
    - First ever run (no snapshot file) → nothing highlighted; snapshot is created
    After comparison the snapshot is updated to the current state, so highlights
    clear automatically on the next run.
    """
    first_run = not SNAPSHOT.exists()
    snapshot = {}
    if not first_run:
        try:
            snapshot = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
        except Exception:
            first_run = True

    new_snapshot = {}
    new_count = 0
    for item in items:
        fp = item_fingerprint(item)
        key = str(item["id"])
        if first_run:
            item["isNew"] = False
        else:
            item["isNew"] = (key not in snapshot or snapshot[key] != fp)
            if item["isNew"]:
                new_count += 1
        new_snapshot[key] = fp

    SNAPSHOT.parent.mkdir(parents=True, exist_ok=True)
    SNAPSHOT.write_text(json.dumps(new_snapshot, indent=2), encoding="utf-8")

    if first_run:
        print("  → Snapshot created (first run — no highlights this time)")
    else:
        print(f"  → {new_count} new/updated item(s) highlighted")

    return items


def read_last_update(path: Path) -> str:
    """Return today's date — always reflects when the script actually ran."""
    return datetime.now().strftime("%-d %B %Y")


def js_string(s: str) -> str:
    """Escape a Python string for embedding in a JS string literal."""
    return (
        s.replace("\\", "\\\\")
         .replace('"', '\\"')
         .replace("\n", "\\n")
         .replace("\r", "")
    )


def generate_items_js(items: list[dict]) -> str:
    lines = ["const ITEMS = ["]
    for item in items:
        parts = [
            f'id:{item["id"]}',
            f'topic:"{js_string(item["topic"])}"',
            f'summary:"{js_string(item["summary"])}"',
            f'impact:"{js_string(item["impact"])}"',
            f'pubDate:"{js_string(item["pubDate"])}"',
            f'issuer:"{js_string(item["issuer"])}"',
            f'status:"{js_string(item["status"])}"',
            f'appDate:"{js_string(item["appDate"])}"',
            f'cpxOrga:{item["cpxOrga"]}',
            f'cpxIT:{item["cpxIT"]}',
            f'link:"{js_string(item["link"])}"',
            f'isNew:{"true" if item.get("isNew") else "false"}',
        ]
        lines.append("  {" + ",".join(parts) + "},")
    lines.append("];")
    return "\n".join(lines)


def build(xlsx_path=XLSX, template_path=TEMPLATE, output_path=OUTPUT):
    print(f"Reading: {xlsx_path.name}")
    items = read_excel(xlsx_path)
    items = mark_new_items(items)
    last_update = read_last_update(xlsx_path)
    new_count = sum(1 for i in items if i.get("isNew"))
    print(f"  → {len(items)} items · Last update: {last_update}")

    template = template_path.read_text(encoding="utf-8")

    build_ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    meta_js = (
        f'const BUILD_META = {{'
        f'lastUpdate:"{js_string(last_update)}",'
        f'builtAt:"{build_ts}",'
        f'itemCount:{len(items)},'
        f'newCount:{new_count}'
        f'}};'
    )
    template = template.replace("// %%BUILD_META%%", meta_js)
    template = template.replace("%%ITEMS_JS%%", generate_items_js(items))

    output_path.write_text(template, encoding="utf-8")
    print(f"Written: {output_path.name}  ({output_path.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    if not XLSX.exists():
        print(f"ERROR: Excel file not found: {XLSX}", file=sys.stderr)
        sys.exit(1)
    if not TEMPLATE.exists():
        print(f"ERROR: Template not found: {TEMPLATE}", file=sys.stderr)
        sys.exit(1)
    build()
    print("Done.")
