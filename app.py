# app.py
import streamlit as st
import pandas as pd
import os
import json
import uuid
import urllib.parse
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re
import threading
import time

# Flask for serving static HTML form pages + submit endpoint
from flask import Flask, render_template_string, request, redirect, url_for

# ----------------------------
# Setup
# ----------------------------
st.set_page_config(page_title="📄 Excel → Web Form (Admin)", layout="centered")
st.title("📄 Auto Form Creator from Excel (Admin)")

DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)
META_PATH = os.path.join(DATA_DIR, "meta.json")

# ----------------------------
# Helper Functions (shared)
# ----------------------------
def load_meta():
    if os.path.exists(META_PATH):
        with open(META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_meta(meta):
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

def detect_dropdowns(excel_file, df_columns):
    # seeks file start (uploaded is BytesIO-like)
    excel_file.seek(0)
    wb = load_workbook(excel_file, data_only=True)
    ws = wb.active
    dropdowns = {}
    if ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            try:
                if dv.type == "list" and dv.formula1:
                    formula = str(dv.formula1).strip('"')
                    if "," in formula:
                        options = [x.strip() for x in formula.split(",")]
                    else:
                        options = []
                        try:
                            rng = formula.split("!")[-1].replace("$", "")
                            a, b = rng.split(":")
                            col_letters = re.match(r"([A-Za-z]+)", a).group(1)
                            start_row = int(re.match(r"[A-Za-z]+([0-9]+)", a).group(1))
                            end_row = int(re.match(r"[A-Za-z]+([0-9]+)", b).group(1))
                            col_idx = column_index_from_string(col_letters)
                            for r in range(start_row, end_row + 1):
                                v = ws.cell(row=r, column=col_idx).value
                                if v is not None:
                                    options.append(str(v))
                        except Exception:
                            options = []
                    # map to first column in dv.cells if possible
                    for cell_range in getattr(dv, "cells", []):
                        try:
                            cidx = cell_range.min_col - 1
                            if 0 <= cidx < len(df_columns):
                                dropdowns[df_columns[cidx]] = options
                        except Exception:
                            continue
            except Exception:
                continue
    return dropdowns

def read_submissions(form_id):
    path = os.path.join(DATA_DIR, f"submissions_{form_id}.xlsx")
    if os.path.exists(path):
        try:
            return pd.read_excel(path)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def save_submissions(form_id, df):
    path = os.path.join(DATA_DIR, f"submissions_{form_id}.xlsx")
    df.to_excel(path, index=False)
    return path

# ----------------------------
# Flask server (serves HTML form pages)
# ----------------------------
flask_app = Flask(__name__)

# Basic CSS + template for the form page (simple, clean)
FORM_HTML_TEMPLATE = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>{{ form_name }}</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body { font-family: Arial, Helvetica, sans-serif; background:#f4f7fb; padding:30px; }
    .card { background: #fff; max-width:720px; margin:0 auto; padding:24px; border-radius:12px; box-shadow:0 8px 24px rgba(20,30,60,0.08); }
    h1 { margin-top:0; font-size:22px; color:#0b3b66; }
    label { display:block; font-weight:600; margin-top:14px; margin-bottom:6px; color:#07426A; }
    input[type="text"], select, textarea { width:100%; padding:10px 12px; border:1px solid #d9e2ef; border-radius:8px; font-size:14px; }
    textarea { min-height:90px; resize:vertical; }
    .btn { margin-top:18px; display:inline-block; padding:10px 18px; background:#0b66ff; color:#fff; border-radius:10px; text-decoration:none; border:none; cursor:pointer; font-weight:700;}
    .small { color:#6b7a90; font-size:13px; margin-top:8px; }
    .success { text-align:center; padding:40px; }
  </style>
</head>
<body>
  <div class="card">
    {% if not success %}
      <h1>{{ form_name }}</h1>
      <p class="small">Please fill the form below and press Submit.</p>
      <form method="post" action="/submit/{{ form_id }}">
        {% for col in columns %}
          <label>{{ col }}</label>
          {% if col in dropdowns %}
            <select name="{{ col|replace(' ', '_') }}">
              <option value="">-- select --</option>
              {% for opt in dropdowns[col] %}
                <option value="{{ opt|e }}">{{ opt }}</option>
              {% endfor %}
            </select>
          {% else %}
            <input type="text" name="{{ col|replace(' ', '_') }}" />
          {% endif %}
        {% endfor %}
        <button class="btn" type="submit">Submit</button>
      </form>
    {% else %}
      <div class="success">
        <h1>Thank you — response saved ✅</h1>
        <p class="small">Your submission was recorded on {{ submitted_at }}</p>
      </div>
    {% endif %}
  </div>
</body>
</html>
"""

@flask_app.route("/form/<form_id>", methods=["GET"])
def serve_form(form_id):
    meta = load_meta()
    forms = meta.get("forms", {})
    if form_id not in forms:
        return "Invalid form ID", 404
    info = forms[form_id]
    return render_template_string(FORM_HTML_TEMPLATE,
                                  form_name=info.get("form_name", "Form"),
                                  form_id=form_id,
                                  columns=info.get("columns", []),
                                  dropdowns=info.get("dropdowns", {}),
                                  success=False)

@flask_app.route("/submit/<form_id>", methods=["POST"])
def handle_submit(form_id):
    meta = load_meta()
    forms = meta.get("forms", {})
    if form_id not in forms:
        return "Invalid form ID", 404
    info = forms[form_id]
    cols = info.get("columns", [])
    # Build a row mapping column names to posted values (underscores reversed)
    row = {"SubmittedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    for col in cols:
        key = col.replace(" ", "_")
        val = request.form.get(key, "").strip()
        row[col] = val
    # Append to submissions file
    subs = read_submissions(form_id)
    try:
        subs = pd.concat([subs, pd.DataFrame([row])], ignore_index=True) if not subs.empty else pd.DataFrame([row])
    except Exception:
        subs = pd.DataFrame([row])
    save_submissions(form_id, subs)
    # Optionally, you can render a success page
    return render_template_string(FORM_HTML_TEMPLATE,
                                  form_name=info.get("form_name", "Form"),
                                  form_id=form_id,
                                  columns=cols,
                                  dropdowns=info.get("dropdowns", {}),
                                  success=True,
                                  submitted_at=row["SubmittedAt"])

def start_flask_in_thread(host="0.0.0.0", port=5001):
    # Run Flask in a separate thread so Streamlit continues to run
    def run():
        # disable reloader to avoid duplicate threads
        flask_app.run(host=host, port=port, debug=False, use_reloader=False)
    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    # small sleep to let server boot (optional)
    time.sleep(0.5)
    return thread

# Start Flask server (only once)
if "flask_started" not in st.session_state:
    try:
        start_flask_in_thread()
        st.session_state["flask_started"] = True
        st.info("HTML form server started at port 5001 (http://localhost:5001/form/<form_id>).")
    except Exception as e:
        st.warning(f"Could not start embedded HTML server: {e}")

# ----------------------------
# Streamlit Admin UI
# ----------------------------
params = st.experimental_get_query_params()
mode = params.get("mode", ["admin"])[0]
# form_id param ignored in admin mode

meta = load_meta()

if mode != "form":
    st.header("🧑‍💼 Admin Panel")
    st.write("Upload an Excel file — its columns will automatically become form fields and you will get a public HTML link.")

    uploaded = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

    if uploaded:
        try:
            df = pd.read_excel(uploaded)
            df.columns = [str(c).strip().replace("_", " ").title() for c in df.columns if pd.notna(c)]
            st.success(f"Detected columns: {len(df.columns)}")
            st.write(df.columns.tolist())

            # detect dropdowns (reads uploaded BytesIO)
            dropdowns = detect_dropdowns(uploaded, list(df.columns))
            if dropdowns:
                st.info("Detected dropdowns:")
                st.table(pd.DataFrame([{"Field": k, "Options": ", ".join(v)} for k, v in dropdowns.items()]))

            form_name = st.text_input("Form name:", value=f"My Form {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            # Base public host for HTML form pages (where Flask will be reachable)
            # For local use: http://localhost:5001
            public_host = st.text_input("Public host for HTML form pages (example: http://localhost:5001)", value="http://localhost:5001")

            if st.button("🚀 Create HTML Form Link"):
                if not public_host:
                    st.error("Please enter the public host where the HTML form will be served (e.g., http://localhost:5001).")
                else:
                    new_form_id = str(uuid.uuid4())[:10]
                    forms = meta.get("forms", {})
                    forms[new_form_id] = {
                        "form_name": form_name,
                        "columns": list(df.columns),
                        "dropdowns": dropdowns,
                        "created_at": datetime.now().isoformat(),
                    }
                    meta["forms"] = forms
                    save_meta(meta)

                    # Create empty submissions file
                    pd.DataFrame(columns=["SubmittedAt"] + list(df.columns)).to_excel(
                        os.path.join(DATA_DIR, f"submissions_{new_form_id}.xlsx"), index=False
                    )

                    # Show public HTML link
                    html_link = public_host.rstrip("/") + f"/form/{new_form_id}"
                    st.success("✅ HTML Form created successfully!")
                    st.info("Share this HTML link with others to open the form (any browser):")
                    st.code(html_link)
        except Exception as e:
            st.error(f"Error processing file: {e}")

    st.markdown("---")
    st.subheader("📊 Existing Forms")
    forms = meta.get("forms", {})
    if forms:
        df_forms = pd.DataFrame([
            {"Form ID": fid, "Form Name": fdata["form_name"], "Created": fdata["created_at"]}
            for fid, fdata in forms.items()
        ])
        st.dataframe(df_forms)
    else:
        st.info("No forms created yet.")
