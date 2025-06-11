import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Excel Subtable Editor", layout="wide")
st.title("📊 Excel Subtable Editor")

# --- File Upload Section ---
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Cache the openpyxl workbook object
        @st.cache_resource(ttl=3600) # Cache for 1 hour (or adjust as needed)
        def load_workbook_from_bytesio(file_buffer):
            # Seek to the beginning before loading
            file_buffer.seek(0)
            return openpyxl.load_workbook(file_buffer, data_only=True)

        # Cache the initial subtable extraction and DataFrame creation
        @st.cache_data(ttl=3600) # Cache for 1 hour
        def get_initial_dataframe(workbook, sheet_name, start_row, end_row, start_col, end_col, use_first_row_as_header):
            ws = workbook[sheet_name]

            data = [
                list(row)
                for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col, values_only=True)
            ]

            if use_first_row_as_header and len(data) > 0:
                raw_headers = list(data[0])
                rows = data[1:]
            else:
                raw_headers = [f"Column_{i+1}" for i in range(end_col - start_col + 1)]
                rows = data

            headers = []
            seen = {}
            for h in raw_headers:
                h_str = str(h) if h is not None and str(h).strip() != "" else "Unnamed"
                if h_str in seen:
                    seen[h_str] += 1
                    h_str = f"{h_str}_{seen[h_str]}"
                else:
                    seen[h_str] = 0
                headers.append(h_str)

            df_result = pd.DataFrame(rows, columns=headers)
            return df_result.dropna(how="all")


        # Load the workbook using the cached function
        wb = load_workbook_from_bytesio(uploaded_file)
        sheet_names = wb.sheetnames

        st.success("File loaded successfully!")

        selected_sheet = st.selectbox("Select a sheet", sheet_names)
        ws = wb[selected_sheet] # Get worksheet from the loaded workbook

        max_row = ws.max_row
        max_col = ws.max_column
        st.write(f"Sheet dimensions: {max_row} rows, {max_col} columns")

        st.markdown("### 🔍 Choose Subtable Selection Method")
        selection_method = st.radio(
            "How do you want to select the subtable?",
            ("Manual Range Input", "Auto-Detect by Blank Rows"),
            index=0
        )

        # Define default df as an empty DataFrame
        df_initial = pd.DataFrame()
        start_row_manual = 1
        end_row_manual = min(start_row_manual + 10, max_row)
        start_col_manual = 1
        end_col_manual = min(start_col_manual + 5, max_col)
        use_first_row_as_header_manual = True

        if selection_method == "Manual Range Input":
            st.markdown("#### Manual Subtable Range Selection")
            start_row_manual = st.number_input("Start Row (from Excel file)", min_value=1, max_value=max_row, value=1, key="start_row_manual")
            end_row_manual = st.number_input("End Row (from Excel file)", min_value=start_row_manual, max_value=max_row, value=min(start_row_manual + 10, max_row), key="end_row_manual")
            start_col_manual = st.number_input("Start Column (A=1)", min_value=1, max_value=max_col, value=1, key="start_col_manual")
            end_col_manual = st.number_input("End Column", min_value=start_col_manual, max_value=max_col, value=min(start_col_manual + 5, max_col), key="end_col_manual")
            use_first_row_as_header_manual = st.checkbox("Use first row of selection as header", value=True, key="use_header_manual")

            # Get the initial DataFrame using the cached function
            df_initial = get_initial_dataframe(wb, selected_sheet, start_row_manual, end_row_manual, start_col_manual, end_col_manual, use_first_row_as_header_manual)

        elif selection_method == "Auto-Detect by Blank Rows":
            st.markdown("#### Auto-Detecting Subtables")
            # For auto-detect, we need to pass the raw uploaded_file BytesIO object
            # to read_excel, as openpyxl might have moved the pointer.
            uploaded_file.seek(0)
            full_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)

            non_empty_rows = full_df.dropna(how='all').index.tolist()
            df_auto_detected = pd.DataFrame() # Initialize auto-detected df

            if non_empty_rows:
                first_data_row_index = non_empty_rows[0]
                last_data_row_index = first_data_row_index
                for i in range(first_data_row_index + 1, len(full_df)):
                    if (i in non_empty_rows) and ((i - 1) in non_empty_rows):
                        last_data_row_index = i
                    else:
                        break

                sub_df_raw = full_df.iloc[first_data_row_index : last_data_row_index + 1]
                non_empty_cols = sub_df_raw.dropna(axis=1, how='all').columns.tolist()

                if non_empty_cols:
                    st.info(f"Auto-detected range: Rows {first_data_row_index + 1} to {last_data_row_index + 1}, Columns {non_empty_cols[0] + 1} to {non_empty_cols[-1] + 1}")

                    auto_headers = sub_df_raw.iloc[0].tolist()
                    auto_rows = sub_df_raw.iloc[1:].values.tolist()

                    headers = []
                    seen = {}
                    for h in auto_headers:
                        h_str = str(h) if h is not None and str(h).strip() != "" else "Unnamed"
                        if h_str in seen:
                            seen[h_str] += 1
                            h_str = f"{h_str}_{seen[h_str]}"
                        else:
                            seen[h_str] = 0
                        headers.append(h_str)

                    df_auto_detected = pd.DataFrame(auto_rows, columns=headers)
                    df_auto_detected = df_auto_detected.dropna(how="all")
                else:
                    st.warning("No contiguous data block found for auto-detection in columns.")
            else:
                st.warning("No non-empty rows found for auto-detection.")

            df_initial = df_auto_detected # Assign the auto-detected df

        # --- Session State Management ---
        # The ID now needs to incorporate the uploaded file's hash or ID to ensure cache invalidation
        # when a new file is uploaded, even if its name is the same.
        # uploaded_file.file_id is unique per upload.
        current_data_selection_id = f"{uploaded_file.file_id}-{selected_sheet}-{selection_method}-{start_row_manual}-{end_row_manual}-{start_col_manual}-{end_col_manual}-{use_first_row_as_header_manual}"

        if "last_processed_file_id" not in st.session_state or st.session_state.last_processed_file_id != current_data_selection_id:
            st.session_state.current_df = df_initial.copy()
            st.session_state.history = []
            st.session_state.last_processed_file_id = current_data_selection_id
            st.info("New file, sheet, or selection parameters detected. Table and history reset.")
        elif st.session_state.current_df.empty and not df_initial.empty:
            st.session_state.current_df = df_initial.copy()
            st.session_state.history = []
            st.session_state.last_processed_file_id = current_data_selection_id
            st.info("Re-initializing table from file as previous data was empty.")

        # --- Display and Editing UI ---
        if not st.session_state.current_df.empty:
            st.subheader("✏️ Edit Table")
            edited_df = st.data_editor(st.session_state.current_df, num_rows="dynamic", use_container_width=True)

            if st.button("Save Changes"):
                if not edited_df.equals(st.session_state.current_df):
                    st.session_state.history.append(st.session_state.current_df.copy())
                    st.session_state.current_df = edited_df.copy()
                    st.success("Changes saved.")
                else:
                    st.info("No changes to save.")

            st.subheader("🔗 Combine Rows")
            st.write("Current table row indices:")
            st.dataframe(st.session_state.current_df.index.to_frame(name='Index'), use_container_width=True)
            st.info("Please select rows using the indices displayed above for the *current table*.")

            selected_rows = st.multiselect(
                "Select rows to combine (by current table index)",
                st.session_state.current_df.index.tolist()
            )
            custom_name = st.text_input("Custom name for the new combined row", value="Combined Row")

            if st.button("Combine Selected Rows"):
                if selected_rows:
                    st.session_state.history.append(st.session_state.current_df.copy())

                    combined_row_data = {}
                    selected_df = st.session_state.current_df.loc[selected_rows]

                    for col in st.session_state.current_df.columns:
                        if pd.api.types.is_numeric_dtype(st.session_state.current_df[col]):
                            combined_row_data[col] = selected_df[col].sum()
                        else:
                            combined_row_data[col] = " / ".join(selected_df[col].astype(str).fillna(''))

                    if st.session_state.current_df.columns.size > 0:
                        first_col_name = st.session_state.current_df.columns[0]
                        combined_row_data[first_col_name] = custom_name

                    combined_df = pd.DataFrame([combined_row_data], columns=st.session_state.current_df.columns)
                    remaining_df = st.session_state.current_df.drop(index=selected_rows)
                    st.session_state.current_df = pd.concat([remaining_df, combined_df], ignore_index=True)
                    st.success("Rows combined successfully.")
                    st.rerun()

                else:
                    st.warning("No rows selected to combine.")

            st.subheader("🧬 Merge Columns")
            selected_cols = st.multiselect("Select columns to merge", st.session_state.current_df.columns.tolist(), key="merge_cols")
            new_col_name = st.text_input("New column name", value="MergedColumn")
            if st.button("Merge Selected Columns"):
                if selected_cols and len(selected_cols) >= 2:
                    if new_col_name in st.session_state.current_df.columns and new_col_name not in selected_cols:
                        st.error(f"Column '{new_col_name}' already exists. Please choose a different name or include it in columns to merge.")
                    else:
                        st.session_state.history.append(st.session_state.current_df.copy())
                        st.session_state.current_df[new_col_name] = st.session_state.current_df[selected_cols].astype(str).agg(" / ".join, axis=1)
                        st.session_state.current_df.drop(columns=selected_cols, inplace=True)
                        st.success(f"Columns merged into '{new_col_name}'")
                else:
                    st.warning("Please select at least two columns to merge.")

            if st.button("Undo Last Action"):
                if st.session_state.history:
                    st.session_state.current_df = st.session_state.history.pop()
                    st.success("Undo successful.")
                    st.rerun()
                else:
                    st.warning("No previous state to undo.")

            st.subheader("📋 Final Table")
            final_df = st.session_state.current_df.dropna(how="all")
            st.dataframe(final_df, use_container_width=True)

            st.subheader("📥 Download Modified Table")
            def to_excel(df_to_save):
                output = BytesIO()
                if not df_to_save.empty:
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df_to_save.to_excel(writer, index=False, sheet_name="ModifiedTable")
                output.seek(0)
                return output

            excel_data = to_excel(final_df)
            st.download_button(
                "Download as Excel",
                data=excel_data,
                file_name="modified_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.info("No data found for the selected range/auto-detection. Please adjust your selection or upload a different file.")

    except Exception as e:
        st.error(f"An error occurred while processing the Excel file: {e}")
        st.info("Please ensure it's a valid Excel file with readable content and try again.")
else:
    st.info("Please upload an Excel file to begin.")
