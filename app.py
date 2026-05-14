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

COMPANY_MAPPING = {
    "Damon and Crew": "Non Y&S",
    "The Ticket Guy": "Non Y&S",
    "YourTickets": "Non Y&S",
    "GK LLC": "Y&S",
    "Jacks YS": "Y&S",
    "Levovitz": "Y&S",
    "Needle Tickets LLC": "Y&S",
    "Pollak Tickets": "Y&S",
    "Yoni Levine": "Y&S",
    "YS Katz": "Y&S",
    "YS Tickets": "Y&S",
    "YS TL": "Y&S",
    "YSA": "Y&S",
    "YSA 2": "Y&S",
    "YSA 3": "Y&S",
    "YSM Tickets": "Y&S",
    "YSS Tickets": "Y&S",
    "YS-Seatgeek": "Y&S",
    "YS-Seatgeek2": "Y&S",
    "YSW": "Y&S",
    "YS Tickets Spec": "Y&S",
}

def get_main_company(company):
    if company in ("YS-Seatgeek", "YS-Seatgeek2", "YS Tickets Spec"):
        return "YS Tickets"
    if company in ("YSA 2", "YSA 3"):
        return "YSA"
    return company

def get_ys_mapping(company):
    return COMPANY_MAPPING.get(company, "Non Y&S")

def build_excel_output(df, sheet_name="Invoices"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        wb = writer.book
        ws = writer.sheets[sheet_name]

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

        MONEY_COLS = {"Amnt", "Bal."}

        headers = list(df.columns)
        col_widths = [max(len(str(h)), 8) for h in headers]

        for ci, h in enumerate(headers):
            ws.write(0, ci, h, hdr_fmt)
            sample = df.iloc[:500, ci].astype(str)
            col_widths[ci] = min(
                max(col_widths[ci], sample.str.len().max() if len(sample) else 8) + 2,
                40
            )

        for ci, h in enumerate(headers):
            fmt = money_fmt if h in MONEY_COLS else data_fmt
            ws.set_column(ci, ci, col_widths[ci], fmt)

        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df), len(headers) - 1)

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
        month_end_date = request.form.get("month_end_date", "").strip()

        if not files or all(f.filename == "" for f in files):
            return jsonify({"error": "No files uploaded."}), 400
        if not month_end_date:
            return jsonify({"error": "Please enter a month end date."}), 400

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

        # Format Created as mm/dd/yyyy
        if "Created" in combined.columns:
            combined["Created"] = pd.to_datetime(combined["Created"], errors="coerce").dt.strftime("%-m/%-d/%Y")

        # Insert Main Company as first column
        combined.insert(0, "Main Company", combined["Company"].apply(get_main_company))

        # Tag Y&S vs Non Y&S
        combined["_ys_flag"] = combined["Company"].apply(get_ys_mapping)

        # Keep only the required output columns
        OUTPUT_COLS = ["Main Company", "Company", "Inv#", "Client", "Ext Order #", "Bal.", "Status", "Created", "User", "Amnt"]
        available_cols = [c for c in OUTPUT_COLS if c in combined.columns] + ["_ys_flag"]
        combined = combined[available_cols]

        df_ys     = combined[combined["_ys_flag"] == "Y&S"].drop(columns=["_ys_flag"])
        df_non_ys = combined[combined["_ys_flag"] == "Non Y&S"].drop(columns=["_ys_flag"])

        fname_ys     = f"Invoices {month_end_date} (YS).xlsx"
        fname_non_ys = f"Invoices {month_end_date} (Non YS).xlsx"

        buf_ys     = build_excel_output(df_ys,     sheet_name="Invoices")
        buf_non_ys = build_excel_output(df_non_ys, sheet_name="Invoices")

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
        response.headers["X-YS-Rows"]    = str(len(df_ys))
        response.headers["X-NonYS-Rows"] = str(len(df_non_ys))
        response.headers["X-Total-Rows"] = str(len(combined))
        response.headers["Access-Control-Expose-Headers"] = "X-YS-Rows, X-NonYS-Rows, X-Total-Rows"
        return response

    except Exception as e:
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
