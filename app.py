import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import os # Import os for path checks

st.set_page_config(page_title="Excel Subtable Editor", layout="wide")
st.title("📊 Excel Subtable Editor")

# --- Configuration for Hardcoded File ---
# Ensure this file is in the same directory as your Streamlit app or provide a full path
excel_file = "Earned Comm Breakdown (Finance)_Apr 2025(Summary).xlsx"

# --- File Loading and Error Handling ---
# Check if the file exists before attempting to load it
if not os.path.exists(excel_file):
    st.error(f"Error: The Excel file '{excel_file}' was not found.")
    st.info("Please ensure the Excel file is in the same directory as this script.")
    st.stop() # Stop the Streamlit app execution if the file is not found

try:
    # Load the Excel file using openpyxl
    # No need for .seek(0) here as we are loading directly from a file path
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    sheet_names = wb.sheetnames

    st.success(f"File '{excel_file}' loaded successfully!")

    # Select sheet
    selected_sheet = st.selectbox("Select a sheet", sheet_names)
    ws = wb[selected_sheet]

    # Get sheet dimensions
    max_row = ws.max_row
    max_col = ws.max_column
    st.write(f"Sheet dimensions: {max_row} rows, {max_col} columns")

    st.markdown("### 🔍 Select Subtable Range")
    start_row = st.number_input("Start Row", min_value=1, max_value=max_row, value=1)
    end_row = st.number_input("End Row", min_value=start_row, max_value=max_row, value=min(start_row + 10, max_row))
    start_col = st.number_input("Start Column (A=1)", min_value=1, max_value=max_col, value=1)
    end_col = st.number_input("End Column", min_value=start_col, max_value=max_col, value=min(start_col + 5, max_col))
    use_first_row_as_header = st.checkbox("Use first row of selection as header", value=True)

    # Extract subtable
    data = [
        list(row)
        for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col, values_only=True)
    ]

    # Handle headers and rows for DataFrame creation
    if use_first_row_as_header and len(data) > 0: # Changed from > 1 to > 0 for robustness with single-row data
        raw_headers = list(data[0])
        rows = data[1:]
    else:
        # Fallback for empty data or if first row is not header
        raw_headers = [f"Column_{i+1}" for i in range(end_col - start_col + 1)]
        rows = data

    # Ensure headers are unique strings
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

    # Create initial DataFrame
    df = pd.DataFrame(rows, columns=headers)
    df = df.dropna(how="all") # Drop rows that are entirely NaN after initial extraction

    # --- Session State Management (Crucial for preserving edits with fixed file) ---
    # Create a unique ID based on the file, sheet, and range parameters.
    # This ID determines when the 'current_df' and 'history' should be reset.
    current_data_selection_id = f"{excel_file}-{selected_sheet}-{start_row}-{end_row}-{start_col}-{end_col}-{use_first_row_as_header}"

    # Only initialize/reset if the selected sheet or range has changed
    if "last_data_selection_id" not in st.session_state or st.session_state.last_data_selection_id != current_data_selection_id:
        st.session_state.current_df = df.copy() # Load fresh data based on new selection
        st.session_state.history = [] # Clear history for new selection
        st.session_state.last_data_selection_id = current_data_selection_id
        st.info("New sheet or range selected. Table and history reset to original data.")
    elif st.session_state.current_df.empty and not df.empty:
        # This case handles scenarios where a previous selection resulted in an empty df,
        # but the new selection now has data. We re-initialize.
        st.session_state.current_df = df.copy()
        st.session_state.history = []
        st.session_state.last_data_selection_id = current_data_selection_id
        st.info("Re-initializing table from file as previous data was empty.")

    # --- Display and Editing UI ---
    if not st.session_state.current_df.empty:
        st.subheader("✏️ Edit Table")
        # Use st.session_state.current_df for the data_editor
        edited_df = st.data_editor(st.session_state.current_df, num_rows="dynamic", use_container_width=True)

        if st.button("Save Changes"):
            if not edited_df.equals(st.session_state.current_df):
                st.session_state.history.append(st.session_state.current_df.copy()) # Save current state before updating
                st.session_state.current_df = edited_df.copy()
                st.success("Changes saved.")
            else:
                st.info("No changes to save.")

        st.subheader("🔗 Combine Rows")
        # Ensure the multiselect operates on the current dataframe's index
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
                # Prevent overwriting an existing column unless it's part of the merge
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
            if not df_to_save.empty: # Only attempt to write to Excel if the DataFrame is not empty
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
        st.info("No data available to display or edit based on the current selection. Please adjust the range or check the Excel file content.")

except Exception as e:
    st.error(f"An unexpected error occurred while processing the Excel file: {e}")
    st.info("Please ensure the Excel file is valid and readable.")
