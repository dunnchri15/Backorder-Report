import os
import json
import pickle
import csv
import io
from datetime import datetime, date
from flask import Flask, request, redirect, url_for, render_template, jsonify

import openpyxl

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "rg-backorder-2026")

DATA_DIR = os.environ.get("DATA_DIR", "/tmp/backorder_data")
os.makedirs(DATA_DIR, exist_ok=True)

PARTS_FILE  = os.path.join(DATA_DIR, "parts.pkl")
NOTES_FILE  = os.path.join(DATA_DIR, "notes.json")
FEEDER_FILE = os.path.join(DATA_DIR, "feeder.json")
META_FILE   = os.path.join(DATA_DIR, "meta.json")

PURCHASED_PLANNERS = {"boes","craig","devowe","glynn","salcedo","slifer","zhang","tracy"}


# ── Storage helpers ───────────────────────────────────────────

def load_notes():
    try:
        with open(NOTES_FILE) as f: return json.load(f)
    except: return {}

def save_notes(notes):
    with open(NOTES_FILE,"w") as f: json.dump(notes, f)

def load_feeder():
    try:
        with open(FEEDER_FILE) as f: return set(json.load(f))
    except: return None

def save_feeder(s):
    with open(FEEDER_FILE,"w") as f: json.dump(list(s), f)

def load_parts():
    try:
        with open(PARTS_FILE,"rb") as f: return pickle.load(f)
    except: return None

def save_parts(data):
    with open(PARTS_FILE,"wb") as f: pickle.dump(data, f)

def load_meta():
    try:
        with open(META_FILE) as f: return json.load(f)
    except: return {}

def save_meta(m):
    with open(META_FILE,"w") as f: json.dump(m, f)


# ── Parsing helpers ───────────────────────────────────────────

def fmt_date(val):
    if not val or str(val).strip() in ("", "None", "nan"): return ""
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d %H:%M:%S"):
        try: return datetime.strptime(s[:len(fmt)], fmt).strftime("%Y-%m-%d")
        except: pass
    # Excel serial number
    try:
        serial = int(float(s))
        if 40000 < serial < 60000:
            d = datetime(1899, 12, 30) + __import__('datetime').timedelta(days=serial)
            return d.strftime("%Y-%m-%d")
    except: pass
    return s[:10] if len(s) >= 10 else ""

def col_find(headers, *names):
    """Case-insensitive fuzzy column finder, returns index or None."""
    lu = {h.lower().replace(" ","").replace("_","").replace("#",""): i for i, h in enumerate(headers)}
    for n in names:
        k = n.lower().replace(" ","").replace("_","").replace("#","")
        if k in lu: return lu[k]
    return None

def safe_float(val):
    try: return float(str(val).replace(",","").strip())
    except: return 0.0

def read_xlsx_rows(file_bytes):
    """Read XLSX, pick the sheet with most columns, return (headers, rows)."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    best_sheet, best_cols = wb.worksheets[0], 0
    for ws in wb.worksheets:
        ws_iter = ws.iter_rows(min_row=1, max_row=1, values_only=True)
        first = next(ws_iter, ())
        ncols = sum(1 for c in first if c is not None)
        if ncols > best_cols:
            best_cols = ncols
            best_sheet = ws
    rows_iter = best_sheet.iter_rows(values_only=True)
    headers = [str(c).strip() if c is not None else "" for c in next(rows_iter, [])]
    rows = list(rows_iter)
    wb.close()
    return headers, rows

def read_csv_rows(file_bytes):
    """Read CSV bytes, return (headers, rows)."""
    for enc in ("utf-8-sig", "latin-1", "utf-8"):
        try:
            text = file_bytes.decode(enc)
            reader = csv.reader(io.StringIO(text))
            rows = list(reader)
            if rows:
                return rows[0], rows[1:]
        except: pass
    return [], []

def process_rows(headers, rows, feeder_set):
    """Convert raw rows into parts dict + summary dict."""
    C = {
        "coord":   col_find(headers,"project coordinator","coordinator"),
        "proj":    col_find(headers,"customer address","customeraddress","project"),
        "order":   col_find(headers,"order no","orderno","order"),
        "po_st":   col_find(headers,"po status","postatus"),
        "part":    col_find(headers,"part no","partno","part","partnumber"),
        "line_st": col_find(headers,"line status","linestatus"),
        "oqty":    col_find(headers,"order qty","orderqty"),
        "sqty":    col_find(headers,"shipped qty","shippedqty"),
        "opqty":   col_find(headers,"open qty","openqty"),
        "inv":     col_find(headers,"inventory"),
        "val":     col_find(headers,"unshipped value","unshippedvalue"),
        "ship":    col_find(headers,"ship date","shipdate"),
        "due":     col_find(headers,"due date","duedate"),
        "bldg":    col_find(headers,"building"),
        "planner": col_find(headers,"planner","buyer"),
        "purch":   col_find(headers,"purchased"),
    }

    def g(row, key):
        idx = C.get(key)
        if idx is None or idx >= len(row): return ""
        v = row[idx]
        return "" if v is None else str(v).strip()

    today = date.today().strftime("%Y-%m-%d")
    parts, summary, max_date = {}, {}, ""
    row_count = 0

    for row in rows:
        if not any(c for c in row if c is not None and str(c).strip()): continue
        row_count += 1

        coord   = g(row,"coord") or "Unassigned"
        proj    = g(row,"proj")  or "Unknown"
        key     = f"{coord}|||{proj}"
        part_no = g(row,"part").rstrip(" -").strip()
        po_st   = g(row,"po_st")
        line_st = g(row,"line_st")
        bldg    = g(row,"bldg")
        planner = g(row,"planner")
        order   = g(row,"order")
        oqty    = safe_float(g(row,"oqty"))
        sqty    = safe_float(g(row,"sqty"))
        opqty   = safe_float(g(row,"opqty"))
        inv     = safe_float(g(row,"inv"))
        val     = safe_float(g(row,"val"))
        ship    = fmt_date(g(row,"ship"))
        due     = fmt_date(g(row,"due"))
        overdue = 1 if ship and "2020" < ship < "2100" and ship < today else 0

        if ship and "2020" < ship < "2100" and ship > max_date:
            max_date = ship

        purchased = 0
        if feeder_set is not None:
            purchased = 1 if part_no in feeder_set else 0
        else:
            purch_val = g(row,"purch").lower()
            if purch_val == "yes": purchased = 1

        # [order, po_st, part, line_st, oqty, sqty, opqty, inv, val, ship, due, overdue, bldg, purchased, planner]
        p = [order, po_st, part_no, line_st, oqty, sqty, opqty, inv, val, ship, due, overdue, bldg, purchased, planner]

        parts.setdefault(key, []).append(p)
        summary.setdefault(coord, {})
        summary[coord].setdefault(proj, {"open_qty":0,"inventory":0,"value":0,"count":0,"overdue":0})
        m = summary[coord][proj]
        m["open_qty"] += opqty; m["inventory"] += inv; m["value"] += val; m["count"] += 1
        if overdue: m["overdue"] += 1

    return parts, summary, max_date or today, row_count


def parse_feeder_csvs(file_list):
    purchased = set()
    for f in file_list:
        try:
            raw = f.read()
            for enc in ("utf-8-sig","latin-1","utf-8"):
                try:
                    text = raw.decode(enc)
                    reader = csv.DictReader(io.StringIO(text))
                    for row in reader:
                        src     = row.get("Part Source","").lower()
                        planner = row.get("Planner","").lower().strip()
                        part_no = row.get("Part No","").strip()
                        if not part_no: continue
                        if "purchased" in src:
                            purchased.add(part_no)
                        elif "manufactured" in src and planner in PURCHASED_PLANNERS:
                            purchased.add(part_no)
                    break
                except: pass
        except: pass
    return purchased


# ── Routes ────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html",
        meta=load_meta(),
        has_data=load_parts() is not None,
        has_feeder=load_feeder() is not None,
        note_count=len(load_notes()),
    )

@app.route("/upload", methods=["POST"])
def upload():
    feeder_files = request.files.getlist("feeder_csvs")
    feeder_set   = load_feeder()
    if feeder_files and any(f.filename for f in feeder_files):
        feeder_set = parse_feeder_csvs(feeder_files)
        save_feeder(feeder_set)

    report_file = request.files.get("report")
    if not report_file or not report_file.filename:
        return redirect(url_for("index"))

    try:
        raw = report_file.read()
        fname = report_file.filename.lower()
        if fname.endswith(".csv"):
            headers, rows = read_csv_rows(raw)
        else:
            headers, rows = read_xlsx_rows(raw)

        parts, summary, report_date, row_count = process_rows(headers, rows, feeder_set)
        save_parts({"parts": parts, "summary": summary})
        save_meta({
            "report_date":  report_date,
            "report_name":  report_file.filename,
            "row_count":    row_count,
            "uploaded_at":  datetime.now().strftime("%Y-%m-%d %H:%M"),
        })
    except Exception as e:
        return render_template("index.html", meta=load_meta(), has_data=False,
            has_feeder=load_feeder() is not None, note_count=len(load_notes()), error=str(e))

    return redirect(url_for("dashboard"))

@app.route("/dashboard")
def dashboard():
    data = load_parts()
    if not data: return redirect(url_for("index"))
    meta  = load_meta()
    notes = load_notes()
    return render_template("dashboard.html",
        meta=meta,
        parts_json=json.dumps(data["parts"],   separators=(",",":")),
        summary_json=json.dumps(data["summary"], separators=(",",":")),
        notes_json=json.dumps(notes,           separators=(",",":")),
        has_feeder=load_feeder() is not None,
    )

@app.route("/api/notes", methods=["GET"])
def get_notes():
    return jsonify(load_notes())

@app.route("/api/notes", methods=["POST"])
def set_notes():
    save_notes(request.get_json(force=True) or {})
    return jsonify({"ok": True})

if __name__ == "__main__":
    app.run(debug=True)
