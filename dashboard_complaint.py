"""
dashboard_complaint.py
Prepare complaint data from Excel and serve the web dashboard locally.

Install:
    pip install pandas openpyxl flask

Run:
    python dashboard_complaint.py --excel complaint.xlsx

Open:
    http://127.0.0.1:9000
"""
from __future__ import annotations
import argparse, json
from pathlib import Path
import pandas as pd
from flask import Flask, jsonify, send_from_directory

BASE_DIR = Path(__file__).resolve().parent
WEB_DIR = BASE_DIR / "complaint_dashboard"
DATA_DIR = WEB_DIR / "data"
OUTPUT_JSON = DATA_DIR / "dashboard_data.json"

def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
    return df

def find_col(df, *names):
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for name in names:
        if name.lower() in lookup:
            return lookup[name.lower()]
    return None

def normalize(raw: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(raw)
    aliases = {
        "tag_age": ("TAG AGE", "Tag Age", "Age"),
        "status": ("Status",),
        "sub_area": ("Sub Area", "Sub_Area", "Case"),
        "related_unit": ("Related Unit", "Related_Unit", "Owner"),
        "feedback_number": ("Feedback Number", "Feedback_Number", "SR Number"),
        "month": ("Month",),
        "tat": ("Within TAT", "TAT"),
        "category": ("CATEG", "Category", "Product"),
    }
    out = pd.DataFrame(index=df.index)
    for target, candidates in aliases.items():
        col = find_col(df, *candidates)
        out[target] = df[col] if col else ""
    out = out.fillna("")
    out["month"] = out["month"].astype(str).str.replace(r"\.0$", "", regex=True)
    return out

def build_payload(df: pd.DataFrame) -> dict:
    closed = df["status"].str.lower().eq("closed")
    opened = ~closed
    months = sorted([x for x in df["month"].unique().tolist() if x and x.lower() != "nan"])
    selected = months[-1] if months else ""
    cur = df[df["month"].eq(selected)] if selected else df
    cur_closed = cur[cur["status"].str.lower().eq("closed")]

    # Executive KPI
    open_cases = int((cur["status"].str.lower() != "closed").sum())
    resolved_cases = int((cur["status"].str.lower() == "closed").sum())
    total = int(len(cur))

    # TAT / resolution
    age = cur["tag_age"].str.lower()
    age_closed = cur_closed["tag_age"].str.lower()
    def count_age(series, text): return int(series.str.contains(text, regex=False).sum())
    tat_open = {
        "0-5 BD": count_age(age[~closed.loc[cur.index]], "0-5"),
        "6-10 BD": count_age(age[~closed.loc[cur.index]], "6-10"),
        "11-20 BD": count_age(age[~closed.loc[cur.index]], "11-20"),
        "Overdue": int(age[~closed.loc[cur.index]].str.contains("overdue", case=False).sum()),
    }
    resolution = {
        "0-5 BD": count_age(age_closed, "0-5"),
        "6-10 BD": count_age(age_closed, "6-10"),
        "11-20 BD": count_age(age_closed, "11-20"),
        "Overdue": int(age_closed.str.contains("overdue", case=False).sum()),
    }

    # Monthly trend
    trend = df.groupby("month").size().reindex(months, fill_value=0)
    trend_rows = [{"month": str(k), "volume": int(v)} for k, v in trend.items()]

    # Category/Sub Area table + MoM
    group_col = "sub_area"
    pivot = df.pivot_table(index=group_col, columns="month", values="feedback_number",
                           aggfunc="count", fill_value=0)
    last_months = months[-4:]
    pivot = pivot.reindex(columns=last_months, fill_value=0)
    if selected in pivot.columns:
        pivot = pivot.sort_values(selected, ascending=False)
    category_rows = []
    for name, row in pivot.iterrows():
        values = {m: int(row[m]) for m in last_months}
        mom = 0
        if len(last_months) >= 2:
            prev, now = int(row[last_months[-2]]), int(row[last_months[-1]])
            mom = round(((now-prev)/prev)*100) if prev else (100 if now else 0)
        category_rows.append({"case": str(name), **values, "mom": mom})

    # Top 5 with owner
    top = cur.groupby(["sub_area","related_unit"]).size().reset_index(name="volume")
    top = top.sort_values("volume", ascending=False).head(5)
    top5 = [{"case": r.sub_area, "owner": r.related_unit, "volume": int(r.volume)}
            for r in top.itertuples()]

    # Product composition
    comp = cur.groupby("category").size().sort_values(ascending=False)
    composition = [{"category": str(k), "volume": int(v)} for k,v in comp.items()]

    details = df.tail(500).to_dict(orient="records")
    return {
        "meta": {"selected_month": selected, "months": months, "rows": int(len(df))},
        "kpi": {"open_cases": open_cases, "resolved_cases": resolved_cases,
                "mtd_volume": total, "ytd_volume": int(len(df))},
        "open_tat": tat_open,
        "resolution": resolution,
        "trend": trend_rows,
        "top5": top5,
        "categories": category_rows,
        "composition": composition,
        "detail": details,
    }

def process_excel(excel_path: str, sheet: str = "Data") -> dict:
    raw = pd.read_excel(excel_path, sheet_name=sheet)
    df = normalize(raw)
    payload = build_payload(df)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] JSON generated: {OUTPUT_JSON}")
    print(f"[OK] Rows: {len(df):,} | Selected month: {payload['meta']['selected_month']}")
    return payload

app = Flask(__name__, static_folder=str(WEB_DIR), static_url_path="")

@app.get("/")
def index():
    return send_from_directory(WEB_DIR, "index.html")

@app.get("/api/dashboard")
def dashboard_api():
    if not OUTPUT_JSON.exists():
        return jsonify({"error": "dashboard_data.json not generated yet"}), 404
    return jsonify(json.loads(OUTPUT_JSON.read_text(encoding="utf-8")))

@app.get("/<path:path>")
def static_files(path):
    return send_from_directory(WEB_DIR, path)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", help="Complaint Excel source file")
    parser.add_argument("--sheet", default="Data", help="Raw data sheet name (default: Data)")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", default=9000, type=int)
    args = parser.parse_args()
    if args.excel:
        process_excel(args.excel, args.sheet)
    elif not OUTPUT_JSON.exists():
        print("[INFO] No --excel supplied. Dashboard will use existing JSON if available.")
    print(f"[LIVE] http://{args.host}:{args.port}")
    app.run(host=args.host, port=args.port, debug=True)
