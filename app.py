import os
import json
import pickle
from datetime import datetime, date
from flask import Flask, request, redirect, url_for, render_template, jsonify

import pandas as pd

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "rg-backorder-2026")

DATA_DIR = os.environ.get("DATA_DIR", "/tmp/backorder_data")
os.makedirs(DATA_DIR, exist_ok=True)

PARTS_FILE  = os.path.join(DATA_DIR, "parts.pkl")
NOTES_FILE  = os.path.join(DATA_DIR, "notes.json")
FEEDER_FILE = os.path.join(DATA_DIR, "feeder.json")
META_FILE   = os.path.join(DATA_DIR, "meta.json")

PURCHASED_PLANNERS = {"boes","craig","devowe","glynn","salcedo","slifer","zhang","tracy"}


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


def fmt_date(val):
    if pd.isna(val) or val == "": return ""
    if isinstance(val, (datetime, date)): return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    for fmt in ("%Y-%m-%d","%m/%d/%Y","%m/%d/%y"):
        try: return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except: pass
    return s[:10] if len(s) >= 10 else s


def col_find(columns, *names):
    lu = {c.lower().replace(" ","").replace("_",""): c for c in columns}
    for n in names:
        k = n.lower().replace(" ","").replace("_","")
        if k in lu: return lu[k]
    return None


def parse_report(file_obj, filename, feeder_set):
    fname = filename.lower()
    if fname.endswith(".csv"):
        df = pd.read_csv(file_obj, encoding="latin1", low_memory=False)
    else:
        xl = pd.ExcelFile(file_obj)
        best, bcols = xl.sheet_names[0], 0
        for name in xl.sheet_names:
            tmp = xl.parse(name, nrows=1)
            if len(tmp.columns) > bcols: bcols=len(tmp.columns); best=name
        df = xl.parse(best, dtype=str)
    df = df.fillna("")
    cols = df.columns.tolist()

    C = {
        "coord":   col_find(cols,"project coordinator","coordinator"),
        "proj":    col_find(cols,"customer address","customeraddress","project"),
        "order":   col_find(cols,"order no","orderno","order#"),
        "po_st":   col_find(cols,"po status","postatus"),
        "part":    col_find(cols,"part no","partno","part#","partnumber"),
        "line_st": col_find(cols,"line status","linestatus"),
        "oqty":    col_find(cols,"order qty","orderqty"),
        "sqty":    col_find(cols,"shipped qty","shippedqty"),
        "opqty":   col_find(cols,"open qty","openqty"),
        "inv":     col_find(cols,"inventory"),
        "val":     col_find(cols,"unshipped value","unshippedvalue"),
        "ship":    col_find(cols,"ship date","shipdate"),
        "due":     col_find(cols,"due date","duedate"),
        "bldg":    col_find(cols,"building"),
        "planner": col_find(cols,"planner","buyer"),
        "purch":   col_find(cols,"purchased"),
    }

    today = date.today().strftime("%Y-%m-%d")
    parts, summary, max_date = {}, {}, ""

    for _, row in df.iterrows():
        coord   = (str(row[C["coord"]])  if C["coord"]   else "").strip() or "Unassigned"
        proj    = (str(row[C["proj"]])   if C["proj"]    else "").strip() or "Unknown"
        key     = f"{coord}|||{proj}"
        part_no = (str(row[C["part"]])   if C["part"]    else "").strip().rstrip(" -").strip()
        po_st   = (str(row[C["po_st"]])  if C["po_st"]   else "").strip()
        line_st = (str(row[C["line_st"]])if C["line_st"] else "").strip()
        bldg    = (str(row[C["bldg"]])   if C["bldg"]    else "").strip()
        planner = (str(row[C["planner"]])if C["planner"] else "").strip()
        order   = (str(row[C["order"]])  if C["order"]   else "").strip()

        def num(k):
            try: return float(str(row[C[k]]).replace(",","")) if C[k] else 0.0
            except: return 0.0

        oqty=num("oqty"); sqty=num("sqty"); opqty=num("opqty")
        inv=num("inv"); val=num("val")
        ship = fmt_date(row[C["ship"]] if C["ship"] else "")
        due  = fmt_date(row[C["due"]]  if C["due"]  else "")
        overdue = 1 if ship and "2020"<ship<"2100" and ship<today else 0
        if ship and "2020"<ship<"2100" and ship>max_date: max_date=ship

        purchased = 0
        if feeder_set is not None:
            purchased = 1 if part_no in feeder_set else 0
        elif C["purch"]:
            purchased = 1 if str(row[C["purch"]]).lower()=="yes" else 0

        p = [order, po_st, part_no, line_st,
             oqty, sqty, opqty, inv, val,
             ship, due, overdue, bldg, purchased, planner]

        parts.setdefault(key, []).append(p)
        summary.setdefault(coord, {})
        summary[coord].setdefault(proj, {"open_qty":0,"inventory":0,"value":0,"count":0,"overdue":0})
        m = summary[coord][proj]
        m["open_qty"]+=opqty; m["inventory"]+=inv; m["value"]+=val; m["count"]+=1
        if overdue: m["overdue"]+=1

    return parts, summary, max_date or today


def parse_feeder_csvs(file_list):
    purchased = set()
    for f in file_list:
        try:
            df = pd.read_csv(f, encoding="latin1", low_memory=False).fillna("")
            for _, row in df.iterrows():
                src     = str(row.get("Part Source","")).lower()
                planner = str(row.get("Planner","")).lower().strip()
                part_no = str(row.get("Part No","")).strip()
                if not part_no: continue
                if "purchased" in src:
                    purchased.add(part_no)
                elif "manufactured" in src and planner in PURCHASED_PLANNERS:
                    purchased.add(part_no)
        except: pass
    return purchased


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
        parts, summary, report_date = parse_report(
            report_file.stream, report_file.filename, feeder_set)
        save_parts({"parts": parts, "summary": summary})
        save_meta({
            "report_date":  report_date,
            "report_name":  report_file.filename,
            "row_count":    sum(len(v) for v in parts.values()),
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
