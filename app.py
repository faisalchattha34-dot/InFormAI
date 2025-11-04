import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📄 Dynamic Form from Excel")

uploaded = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded:
    df_raw = pd.read_excel(uploaded, header=None)
    # detect header row
    header_row_index = None
    for i in range(len(df_raw)):
        row = df_raw.iloc[i]
        if row.notna().sum() > 2:
            header_row_index = i
            break
    df = pd.read_excel(uploaded, header=header_row_index)
    columns = [str(c).strip().replace("_", " ").title() for c in df.columns if pd.notna(c)]
    st.success(f"Detected Columns: {columns}")

    st.subheader("🧾 Fill the form below")
    data = {}
    for col in columns:
        data[col] = st.text_input(f"{col}")

    if st.button("Submit"):
        new_row = pd.DataFrame([data])
        # append to the original dataframe
        updated_df = pd.concat([df, new_row], ignore_index=True)

        # download updated Excel file
        output = BytesIO()
        updated_df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("⬇️ Download Updated Excel", output, file_name="updated_form_data.xlsx")
        st.success("✅ Data added successfully!")
