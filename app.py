from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import io
import os
import zipfile
import traceback

try:
    import python_calamine  # noqa
    EXCEL_ENGINE = "calamine"
except ImportError:
    EXCEL_ENGINE = "openpyxl"

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB

# Mapping from Master_Mapping_List.xlsx — columns B (TicketVault Company) and D (Inventory/Invoices Month-End File Prep)
# Rows where D is N/A are excluded (not in output files)
COMPANY_MAPPING = {
    "Damon and Crew":    "Non Y&S",
    "The Ticket Guy":    "Non Y&S",
    "YourTickets":       "Non Y&S",
    "GK LLC":            "Y&S",
    "Jacks YS":          "Y&S",
    "Levovitz":          "Y&S",
    "Needle Tickets LLC": "Y&S",
    "Pollak Tickets":    "Y&S",
    "Yoni Levine":       "Y&S",
    "YS Katz":           "Y&S",
    "YS Tickets":        "Y&S",
    "YS TL":             "Y&S",
    "YSA":               "Y&S",
    "YSA 2":             "Y&S",
    "YSA 3":             "Y&S",
    "YSM Tickets":       "Y&S",
    "YSS Tickets":       "Y&S",
    "YS-Seatgeek":       "Y&S",
    "YS-Seatgeek2":      "Y&S",
    "YSW":               "Y&S",
    "YS Tickets Spec":   "Y&S",
}

def get_main_company(company):
    c = company.lower() if company else ""
    if c in ("ys-seatgeek", "ys-seatgeek2", "ys-seatgeek2", "ys tickets spec"):
        return "YS Tickets"
    if c in ("ysa 2", "ysa 3"):
        return "YSA"
    return company

# Lowercase lookup for case-insensitive matching
COMPANY_MAPPING_LOWER = {k.lower(): v for k, v in COMPANY_MAPPING.items()}

def get_ys_mapping(company):
    return COMPANY_MAPPING_LOWER.get(company.lower() if company else "", None)

MONEY_COLS = {"Amnt", "Bal.", "Difference"}

def _style_sheet(wb, ws, df):
    hdr_fmt = wb.add_format({
        "bold": True,
        "font_name": "Arial",
        "font_size": 10,
        "font_color": "#FFFFFF",
        "bg_color": "#1F4E79",
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True,
        "border": 0,
    })
    money_fmt = wb.add_format({
        "font_name": "Arial",
        "font_size": 10,
        "num_format": "#,##0.00",
    })
    data_fmt = wb.add_format({
        "font_name": "Arial",
        "font_size": 10,
    })
    headers = list(df.columns)
    col_widths = [max(len(str(h)), 8) for h in headers]
    for ci, h in enumerate(headers):
        ws.write(0, ci, h, hdr_fmt)
        sample = df.iloc[:500, ci].astype(str)
        col_widths[ci] = min(max(col_widths[ci], sample.str.len().max() if len(sample) else 8) + 2, 40)
    for ci, h in enumerate(headers):
        ws.set_column(ci, ci, col_widths[ci], money_fmt if h in MONEY_COLS else data_fmt)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(df), len(headers) - 1)

def build_excel_output(df_all, df_non_sq):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df_all.to_excel(writer, sheet_name="All Invoices", index=False)
        df_non_sq.to_excel(writer, sheet_name="Non-Sales Queue", index=False)
        wb = writer.book
        _style_sheet(wb, writer.sheets["All Invoices"],    df_all)
        _style_sheet(wb, writer.sheets["Non-Sales Queue"], df_non_sq)
    buf.seek(0)
    return buf

@app.errorhandler(Exception)
def handle_exception(e):
    return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():
    try:
        files = request.files.getlist("files")

        if not files or all(f.filename == "" for f in files):
            return jsonify({"error": "No files uploaded."}), 400

        dfs = []
        for f in files:
            if f.filename.endswith(".xlsx"):
                try:
                    df = pd.read_excel(f, dtype=str, engine=EXCEL_ENGINE)
                    dfs.append(df)
                except Exception as e:
                    return jsonify({"error": f"Could not read {f.filename}: {str(e)}"}), 400

        if not dfs:
            return jsonify({"error": "No valid Excel files found."}), 400

        combined = pd.concat(dfs, ignore_index=True)

        # Convert numeric columns
        for col in ["Amnt", "Bal."]:
            if col in combined.columns:
                combined[col] = pd.to_numeric(combined[col], errors="coerce")

        # Parse Created datetime (keep raw for date extraction, then format)
        if "Created" in combined.columns:
            created_dt = pd.to_datetime(combined["Created"], errors="coerce", format="mixed")
            # Derive file name date from the latest Created date in the data
            max_date = created_dt.max()
            month_end_date = f"{max_date.month}-{max_date.day}-{max_date.year}" if pd.notna(max_date) else "Unknown"
            # Format Created column for output
            combined["Created"] = created_dt.apply(
                lambda d: f"{d.month}/{d.day}/{d.year}" if pd.notna(d) else ""
            )
        else:
            month_end_date = "Unknown"

        # Exclude rows where Bal. is 1 cent or less
        if "Bal." in combined.columns:
            combined = combined[combined["Bal."] > 0.01]

        # Exclude Ticket Evolution client rows, except when Company is The Ticket Guy
        if "Client" in combined.columns:
            is_te = combined["Client"].str.strip().str.lower() == "ticket evolution"
            is_ttg = combined["Company"].str.strip().str.lower() == "the ticket guy"
            combined = combined[~is_te | is_ttg]

        # Insert Main Company as first column
        combined.insert(0, "Main Company", combined["Company"].apply(get_main_company))

        # Tag Y&S vs Non Y&S; companies not in mapping (N/A in master list) are excluded
        combined["_ys_flag"] = combined["Company"].apply(get_ys_mapping)
        combined = combined[combined["_ys_flag"].isin(["Y&S", "Non Y&S"])]

        # Keep only required columns — Amnt before Bal.
        OUTPUT_COLS = ["Main Company", "Company", "Inv#", "Client", "Ext Order #", "Amnt", "Bal.", "Status", "Created", "User"]
        available_cols = [c for c in OUTPUT_COLS if c in combined.columns] + ["_ys_flag"]
        combined = combined[available_cols].copy()

        # Add Difference = Amnt - Bal., inserted right after Bal.
        combined["Difference"] = combined["Amnt"] - combined["Bal."]
        cols = list(combined.columns)
        cols.remove("Difference")
        cols.insert(cols.index("Bal.") + 1, "Difference")
        combined = combined[cols]

        df_ys     = combined[combined["_ys_flag"] == "Y&S"].drop(columns=["_ys_flag"])
        df_non_ys = combined[combined["_ys_flag"] == "Non Y&S"].drop(columns=["_ys_flag"])

        # Non-Sales Queue tabs: rows where User != "Sales Queue"
        df_ys_nsq     = df_ys[df_ys["User"].str.strip() != "Sales Queue"]
        df_non_ys_nsq = df_non_ys[df_non_ys["User"].str.strip() != "Sales Queue"]

        fname_ys     = f"Invoices {month_end_date} (YS).xlsx"
        fname_non_ys = f"Invoices {month_end_date} (Non YS).xlsx"

        buf_ys     = build_excel_output(df_ys,     df_ys_nsq)
        buf_non_ys = build_excel_output(df_non_ys, df_non_ys_nsq)

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(fname_ys,     buf_ys.read())
            zf.writestr(fname_non_ys, buf_non_ys.read())
        zip_buf.seek(0)

        response = send_file(
            zip_buf,
            mimetype="application/zip",
            as_attachment=True,
            download_name=f"Invoices {month_end_date}.zip"
        )
        response.headers["X-YS-Rows"]       = str(len(df_ys))
        response.headers["X-NonYS-Rows"]    = str(len(df_non_ys))
        response.headers["X-Total-Rows"]    = str(len(combined))
        response.headers["X-Month-End-Date"] = month_end_date
        response.headers["Access-Control-Expose-Headers"] = "X-YS-Rows, X-NonYS-Rows, X-Total-Rows, X-Month-End-Date"
        return response

    except Exception as e:
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
