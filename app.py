import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Excel Smart Form", page_icon="📄", layout="centered")
st.title("📄 Dynamic Excel Form with Auto Dropdown Detection")

uploaded = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded:
    # Step 1️⃣: Detect header row
    df_raw = pd.read_excel(uploaded, header=None)
    header_row_index = None
    for i in range(len(df_raw)):
        row = df_raw.iloc[i]
        if row.notna().sum() > 2:
            header_row_index = i
            break

    # Step 2️⃣: Read DataFrame
    df = pd.read_excel(uploaded, header=header_row_index)
    df.columns = [str(c).strip().replace("_", " ").title() for c in df.columns if pd.notna(c)]

    st.success("✅ Columns Detected Successfully!")
    st.write("**Detected Columns:**", list(df.columns))

    # Step 3️⃣: Detect dropdowns (data validation)
    uploaded.seek(0)
    wb = load_workbook(uploaded, data_only=True)
    ws = wb.active

    dropdown_dict = {}

    if ws.data_validations is not None:
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1:
                formula = dv.formula1.strip('"')
                if "," in formula:
                    values = [v.strip() for v in formula.split(",")]
                    for cell_range in dv.cells:
                        col_letter = ''.join([c for c in cell_range if c.isalpha()])
                        col_index = ord(col_letter.upper()) - 65
                        if col_index < len(df.columns):
                            dropdown_dict[df.columns[col_index]] = values

    # Step 4️⃣: Show dropdown detection results
    if dropdown_dict:
        st.info("🎯 **Detected Dropdown Columns:**")
        st.table(pd.DataFrame([
            {"Column": col, "Dropdown Options": ", ".join(vals)} 
            for col, vals in dropdown_dict.items()
        ]))
    else:
        st.warning("⚠️ No dropdown lists detected in this Excel file.")

    # Step 5️⃣: Dynamic Form
    st.subheader("🧾 Fill the Form Below")
    data = {}
    for col in df.columns:
        if col in dropdown_dict:
            data[col] = st.selectbox(f"{col}", dropdown_dict[col])
        else:
            data[col] = st.text_input(f"{col}")

    # Step 6️⃣: Submit + Download Updated File
    if st.button("✅ Submit"):
        new_row = pd.DataFrame([data])
        updated_df = pd.concat([df, new_row], ignore_index=True)

        output = BytesIO()
        updated_df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="⬇️ Download Updated Excel",
            data=output,
            file_name="updated_form_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("🎉 Data added successfully!")
