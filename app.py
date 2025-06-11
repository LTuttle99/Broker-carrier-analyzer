import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Excel Subtable Editor", layout="wide")
st.title("📊 Excel Subtable Editor")

# --- File Upload Section ---
# This widget allows the user to upload an Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

# Only proceed if a file has been uploaded
if uploaded_file is not None:
    try:
        # Crucial: Reset the pointer of the BytesIO object to the beginning.
        # This ensures that openpyxl and pandas can read the file correctly from the start,
        # especially if Streamlit reruns.
        uploaded_file.seek(0)
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet_names = wb.sheetnames

        st.success("File loaded successfully!")

        # Select sheet
        selected_sheet = st.selectbox("Select a sheet", sheet_names)
        ws = wb[selected_sheet]

        # Get sheet dimensions (using openpyxl's calculated dimensions for robustness)
        max_row = ws.max_row
        max_col = ws.max_column
        st.write(f"Sheet dimensions: {max_row} rows, {max_col} columns")

        # --- Subtable Selection Method ---
        st.markdown("### 🔍 Choose Subtable Selection Method")
        selection_method = st.radio(
            "How do you want to select the subtable?",
            ("Manual Range Input", "Auto-Detect by Blank Rows"),
            index=0 # Default to manual input
        )

        # Initialize df as an empty DataFrame. This ensures 'df' is always defined.
        df = pd.DataFrame()

        if selection_method == "Manual Range Input":
            st.markdown("#### Manual Subtable Range Selection")
            start_row = st.number_input("Start Row", min_value=1, max_value=max_row, value=1)
            end_row = st.number_input("End Row", min_value=start_row, max_value=max_row, value=min(start_row + 10, max_row))
            start_col = st.number_input("Start Column (A=1)", min_value=1, max_value=max_col, value=1)
            end_col = st.number_input("End Column", min_value=start_col, max_value=max_col, value=min(start_col + 5, max_col))
            use_first_row_as_header = st.checkbox("Use first row of selection as header", value=True)

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

            df = pd.DataFrame(rows, columns=headers)
            df = df.dropna(how="all")

        elif selection_method == "Auto-Detect by Blank Rows":
            st.markdown("#### Auto-Detecting Subtables")

            # Crucial: Reset pointer before pandas reads the file again
            # This is important as openpyxl might have left the pointer at the end.
            uploaded_file.seek(0)
            # Read the entire sheet into a pandas DataFrame without headers
            full_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)

            # Identify non-empty rows
            non_empty_rows = full_df.dropna(how='all').index.tolist()

            if non_empty_rows:
                first_data_row_index = non_empty_rows[0]
                last_data_row_index = first_data_row_index

                # Find the end of the first contiguous block of non-empty data
                for i in range(first_data_row_index + 1, len(full_df)):
                    if (i in non_empty_rows) and ((i - 1) in non_empty_rows):
                        last_data_row_index = i
                    else:
                        break # Found a gap or end of data

                # Extract the sub-DataFrame for the identified rows
                sub_df_raw = full_df.iloc[first_data_row_index : last_data_row_index + 1]
                # Identify non-empty columns within this sub-DataFrame
                non_empty_cols = sub_df_raw.dropna(axis=1, how='all').columns.tolist()

                if non_empty_cols:
                    st.info(f"Auto-detected range: Rows {first_data_row_index + 1} to {last_data_row_index + 1}, Columns {non_empty_cols[0] + 1} to {non_empty_cols[-1] + 1}")

                    # Use the first row of the detected sub-DataFrame as potential headers
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

                    df = pd.DataFrame(auto_rows, columns=headers)
                    df = df.dropna(how="all")

                else:
                    st.warning("No contiguous data block found for auto-detection in columns within the identified rows.")
            else:
                st.warning("No non-empty rows found for auto-detection.")

        # --- Session State Management ---
        # A unique identifier for the current state of file, sheet, and selection method
        # This will reset current_df and history if the user uploads a new file,
        # selects a different sheet, or changes the selection method.
        current_file_id = f"{uploaded_file.name}-{selected_sheet}-{selection_method}"

        # Initialize or reset session state if the source data configuration changes
        if "last_processed_file_id" not in st.session_state or st.session_state.last_processed_file_id != current_file_id:
            st.session_state.current_df = df.copy() # Set current_df to the newly extracted df
            st.session_state.history = [] # Clear history
            st.session_state.last_processed_file_id = current_file_id
            st.info("New file, sheet, or selection method detected. Table and history reset.")
        # If current_df is empty but a fresh df was just generated (e.g., after an empty selection)
        # we should update current_df.
        elif st.session_state.current_df.empty and not df.empty:
            st.session_state.current_df = df.copy()
            st.session_state.history = []
            st.session_state.last_processed_file_id = current_file_id
            st.info("Re-initializing table from file as previous data was empty.")

        # Only display editing options if a DataFrame was successfully extracted and is not empty
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
            selected_rows = st.multiselect("Select rows to combine (by index)", st.session_state.current_df.index.tolist())
            custom_name = st.text_input("Custom name for the new combined row", value="Combined Row")
            if st.button("Combine Selected Rows"):
                if selected_rows:
                    combined_row = {}
                    selected_df = st.session_state.current_df.loc[selected_rows]
                    for col in st.session_state.current_df.columns:
                        if pd.api.types.is_numeric_dtype(st.session_state.current_df[col]):
                            combined_row[col] = selected_df[col].sum()
                        else:
                            # Use fillna('') to handle potential NaN values in string columns gracefully
                            combined_row[col] = " / ".join(selected_df[col].astype(str).fillna(''))

                    # Assign custom name to the first column, if the DataFrame has columns
                    if st.session_state.current_df.columns.size > 0:
                        first_col_name = st.session_state.current_df.columns[0]
                        combined_row[first_col_name] = custom_name

                    st.session_state.history.append(st.session_state.current_df.copy())
                    # Drop selected rows and reset index, then concatenate the new combined row
                    st.session_state.current_df = st.session_state.current_df.drop(index=selected_rows).reset_index(drop=True)
                    st.session_state.current_df = pd.concat([st.session_state.current_df, pd.DataFrame([combined_row])], ignore_index=True)
                    st.success("Rows combined successfully.")
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
                else:
                    st.warning("No previous state to undo.")

            st.subheader("📋 Final Table")
            final_df = st.session_state.current_df.dropna(how="all")
            st.dataframe(final_df, use_container_width=True)

            st.subheader("📥 Download Modified Table")
            def to_excel(df_to_save):
                output = BytesIO()
                # Only attempt to write to Excel if the DataFrame is not empty
                if not df_to_save.empty:
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df_to_save.to_excel(writer, index=False, sheet_name="ModifiedTable")
                output.seek(0) # Always reset pointer to the beginning for the download button
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
