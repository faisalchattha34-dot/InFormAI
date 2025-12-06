# ----------------------------
# Responses Dashboard with Form Editing
# ----------------------------
st.markdown("---")
st.subheader("📊 Responses Dashboard")
responses = load_responses()
if responses.empty:
    st.info("No responses submitted yet.")
else:
    form_filter = st.selectbox(
        "Select Form to View Responses:",
        ["All"] + [f["form_name"] for f in meta.get("forms", {}).values()]
    )

    if form_filter != "All":
        form_id_list = [fid for fid, f in meta["forms"].items() if f["form_name"] == form_filter]
        responses_display = responses[responses["FormID"] == form_id_list[0]] if form_id_list else pd.DataFrame()
    else:
        responses_display = responses.copy()

    if not responses_display.empty:
        hidden_cols = ["FormID", "FormName", "UserSession", "SubmittedAt"]
        display_df = responses_display.drop(columns=[c for c in hidden_cols if c in responses_display.columns])
        st.dataframe(display_df, use_container_width=True)

        # ----------------------------
        # Select a response to edit
        # ----------------------------
        st.markdown("### ✏️ Edit a Response")
        response_idx = st.number_input(
            "Select Row Number to Edit:",
            min_value=0,
            max_value=len(responses_display) - 1,
            step=1
        )

        if st.button("✍️ Edit Selected Response"):
            row_to_edit = responses_display.iloc[response_idx]
            selected_form_id = row_to_edit["FormID"]
            form_info = meta["forms"].get(selected_form_id, {})
            if form_info:
                st.markdown(f"#### Editing Response for Form: {form_info['form_name']}")
                dropdowns = form_info.get("dropdowns", {})
                columns = form_info.get("columns", [])

                # Show form with current response values
                with st.form("edit_response_form"):
                    edited_values = {}
                    for col in columns:
                        if col in dropdowns:
                            edited_values[col] = st.selectbox(
                                col,
                                dropdowns[col],
                                index=dropdowns[col].index(row_to_edit[col]) if row_to_edit[col] in dropdowns[col] else 0
                            )
                        else:
                            edited_values[col] = st.text_input(col, value=row_to_edit[col])
                    submitted_edit = st.form_submit_button("💾 Save Edited Response")

                if submitted_edit:
                    for col, val in edited_values.items():
                        responses.at[row_to_edit.name, col] = val
                    save_responses(responses)
                    st.success("✅ Response updated successfully!")
                    st.experimental_rerun()
