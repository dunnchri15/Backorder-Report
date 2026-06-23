import os
import json
import pickle
import csv
import io
import traceback
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


# ── Storage ───────────────────────────────────────────────────

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


# ── Parsing ───────────────────────────────────────────────────

def fmt_date(val):
    if not val or str(val).strip() in ("", "None", "nan"): return ""
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try: return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except: pass
    try:
        serial = int(float(s))
        if 40000 < serial < 60000:
            from datetime import timedelta
            d = datetime(1899, 12, 30) + timedelta(days=serial)
            return d.strftime("%Y-%m-%d")
    except: pass
    return s[:10] if len(s) >= 10 else ""

def col_find(headers, *names):
    lu = {str(h).lower().replace(" ","").replace("_","").replace("#",""): i
          for i, h in enumerate(headers)}
    for n in names:
        k = n.lower().replace(" ","").replace("_","").replace("#","")
        if k in lu: return lu[k]
    return None

def safe_float(val):
    try: return float(str(val).replace(",","").strip())
    except: return 0.0

def read_xlsx_rows(file_bytes):
    """Read XLSX using openpyxl, return (headers, rows_as_lists)."""
    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes), read_only=True, data_only=True, keep_links=False
    )
    # Pick sheet with most non-empty header columns
    best_ws, best_cols = wb.worksheets[0], 0
    for ws in wb.worksheets:
        row_iter = ws.iter_rows(min_row=1, max_row=1, values_only=True)
        first = list(next(row_iter, []))
        ncols = sum(1 for c in first if c is not None)
        if ncols > best_cols:
            best_cols = ncols
            best_ws = ws

    all_rows = list(best_ws.iter_rows(values_only=True))
    wb.close()

    if not all_rows:
        return [], []

    headers = [str(c).strip() if c is not None else "" for c in all_rows[0]]
    data_rows = all_rows[1:]
    return headers, data_rows

def read_csv_rows(file_bytes):
    """Read CSV bytes, return (headers, rows_as_lists)."""
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
    parts, summary, max_date, row_count = {}, {}, "", 0

    for row in rows:
        # Skip completely empty rows
        if not any(c for c in row if c is not None and str(c).strip()):
            continue
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
            if g(row,"purch").lower() == "yes":
                purchased = 1

        p = [order, po_st, part_no, line_st,
             oqty, sqty, opqty, inv, val,
             ship, due, overdue, bldg, purchased, planner]

        parts.setdefault(key, []).append(p)
        summary.setdefault(coord, {})
        summary[coord].setdefault(proj, {"open_qty":0,"inventory":0,"value":0,"count":0,"overdue":0})
        m = summary[coord][proj]
        m["open_qty"] += opqty
        m["inventory"] += inv
        m["value"] += val
        m["count"] += 1
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
        error=None,
    )

@app.route("/upload", methods=["POST"])
def upload():
    error = None

    # Handle feeder CSVs
    feeder_files = request.files.getlist("feeder_csvs")
    feeder_set   = load_feeder()
    if feeder_files and any(f.filename for f in feeder_files):
        try:
            feeder_set = parse_feeder_csvs(feeder_files)
            save_feeder(feeder_set)
        except Exception as e:
            error = f"Feeder CSV error: {e}"

    # Handle main report
    report_file = request.files.get("report")
    if not report_file or not report_file.filename:
        return redirect(url_for("index"))

    try:
        raw   = report_file.read()
        fname = report_file.filename.lower()

        if fname.endswith(".csv"):
            headers, rows = read_csv_rows(raw)
        else:
            headers, rows = read_xlsx_rows(raw)

        if not headers:
            raise ValueError("No headers found — is this a valid XLSX or CSV file?")

        parts, summary, report_date, row_count = process_rows(headers, rows, feeder_set)

        if not parts:
            raise ValueError(f"No data rows parsed. Headers detected: {headers[:5]}")

        save_parts({"parts": parts, "summary": summary})
        save_meta({
            "report_date":  report_date,
            "report_name":  report_file.filename,
            "row_count":    row_count,
            "uploaded_at":  datetime.now().strftime("%Y-%m-%d %H:%M"),
        })
        return redirect(url_for("dashboard"))

    except Exception as e:
        tb = traceback.format_exc()
        print("UPLOAD ERROR:", tb)
        error = str(e)
        return render_template("index.html",
            meta=load_meta(), has_data=load_parts() is not None,
            has_feeder=load_feeder() is not None,
            note_count=len(load_notes()), error=error)


@app.route("/dashboard")
def dashboard():
    data = load_parts()
    if not data:
        return redirect(url_for("index"))
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

# Debug route — shows what happened on last upload attempt
@app.route("/debug")
def debug():
    meta = load_meta()
    data = load_parts()
    feeder = load_feeder()
    return jsonify({
        "meta": meta,
        "has_parts": data is not None,
        "parts_keys": list(data["parts"].keys())[:5] if data else [],
        "has_feeder": feeder is not None,
        "feeder_size": len(feeder) if feeder else 0,
        "notes_count": len(load_notes()),
        "data_dir": DATA_DIR,
        "data_dir_files": os.listdir(DATA_DIR),
    })

if __name__ == "__main__":
    app.run(debug=True)
