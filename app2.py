# ----------------------------
# Responses Dashboard
# ----------------------------
st.markdown("---")
st.subheader("📊 Responses Dashboard")
responses = load_responses()

if responses.empty:
    st.info("No responses submitted yet.")
else:
    form_filter = st.selectbox("Select Form to View Responses:", ["All"] + [f["form_name"] for f in meta.get("forms", {}).values()])
    if form_filter != "All":
        form_id_list = [fid for fid, f in meta["forms"].items() if f["form_name"] == form_filter]
        responses_display = responses[responses["FormID"] == form_id_list[0]] if form_id_list else pd.DataFrame()
    else:
        responses_display = responses.copy()

    if not responses_display.empty:
        hidden_cols = ["FormID", "FormName", "UserSession", "SubmittedAt"]
        display_df = responses_display.drop(columns=[c for c in hidden_cols if c in responses_display.columns])

        # ---------------- Edit/Delete/Add/Restore Functionality ----------------
        st.write("### ✏️ Edit Responses")
        edited_df = st.data_editor(
            display_df,
            use_container_width=True,
            num_rows="dynamic",
            key="responses_editor",
        )

        # Save changes
        if st.button("💾 Save Changes to Responses"):
            # Merge metadata back
            for col in hidden_cols:
                if col in responses_display.columns:
                    edited_df[col] = responses_display[col].values
            save_responses(edited_df)
            st.success("✅ Response data updated successfully!")

        # ---------------- Download Responses ----------------
        to_download = BytesIO()
        edited_df.to_excel(to_download, index=False)
        to_download.seek(0)
        st.download_button(
            label="📥 Download Responses",
            data=to_download,
            file_name="responses.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
