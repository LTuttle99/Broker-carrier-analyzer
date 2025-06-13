import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Excel Subtable Editor", layout="wide")
st.title("📊 Excel Subtable Editor")

# --- Configuration Constants ---
AUTO_FILL_START_ROW = 23
AUTO_FILL_END_ROW = 36
AUTO_FILL_START_COL = 2
AUTO_FILL_END_COL = 5
ROWS_TO_AUTO_REMOVE = [8, 10, 13]
ROWS_TO_AUTO_COMBINE = [11, 12]
AUTO_COMBINED_ROW_NAME = "MT - Without FV"

# --- File Upload Section (Moved to Sidebar) ---
with st.sidebar:
    st.header("Upload Excel File")
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])
    st.markdown("---")

# --- Utility Functions (Cached) ---
@st.cache_resource(ttl=3600)
def load_workbook_from_bytesio(file_buffer):
    """Loads an OpenPyXL workbook from a BytesIO object."""
    file_buffer.seek(0)
    return openpyxl.load_workbook(file_buffer, data_only=True)

@st.cache_data(ttl=3600)
def get_initial_dataframe(_workbook, sheet_name, start_row, end_row, start_col, end_col, use_first_row_as_header):
    """
    Extracts a DataFrame from a specified range within an OpenPyXL worksheet.
    Handles header logic and ensures unique column names.
    """
    ws = _workbook[sheet_name]

    start_row = max(1, start_row)
    end_row = max(start_row, end_row)
    start_col = max(1, start_col)
    end_col = max(start_col, end_col)

    data = []
    try:
        for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col, values_only=True):
            data.append(list(row))
    except Exception as e:
        st.error(f"Error reading specified range from sheet: {e}. Please check your row/column inputs.")
        return pd.DataFrame()

    if not data:
        return pd.DataFrame()

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
    
    adjusted_rows = []
    expected_cols = len(headers)
    for row in rows:
        if len(row) < expected_cols:
            adjusted_rows.append(list(row) + [None] * (expected_cols - len(row)))
        else:
            adjusted_rows.append(list(row[:expected_cols]))
            
    df_result = pd.DataFrame(adjusted_rows, columns=headers)
    df_result = df_result.dropna(how="all")

    if 'Order' not in df_result.columns:
        df_result.insert(0, 'Order', range(1, len(df_result) + 1))

    return df_result

# --- Main App Logic ---
if uploaded_file is not None:
    try:
        with st.spinner("Loading Excel file..."):
            wb = load_workbook_from_bytesio(uploaded_file)
        
        sheet_names = wb.sheetnames
        st.success("File loaded successfully!")

        with st.sidebar:
            st.header("Sheet Selection")
            selected_sheet = st.selectbox("Select a sheet", sheet_names, key="selected_sheet_sidebar")
            ws = wb[selected_sheet]
            max_row = ws.max_row
            max_column = ws.max_column
            st.info(f"Sheet dimensions: {max_row} rows, {max_column} columns")
            st.markdown("---")

        st.markdown("### 🔍 Choose Subtable Selection Method")
        selection_method = st.radio(
            "How do you want to select the subtable?",
            ("Manual Range Input", "Auto-Detect by Blank Rows"),
            index=0,
            key="selection_method_radio"
        )

        df_initial = pd.DataFrame()

        if "start_row_manual_val" not in st.session_state:
            st.session_state.start_row_manual_val = 1
        if "end_row_manual_val" not in st.session_state:
            st.session_state.end_row_manual_val = min(st.session_state.start_row_manual_val + 10, max_row)
        if "start_col_manual_val" not in st.session_state:
            st.session_state.start_col_manual_val = 1
        if "end_col_manual_val" not in st.session_state:
            st.session_state.end_col_manual_val = min(st.session_state.start_col_manual_val + 5, max_column)
        if "use_header_manual_val" not in st.session_state:
            st.session_state.use_header_manual_val = True

        auto_fill_toggle = st.toggle(
            f"Auto-fill with predefined range (Rows {AUTO_FILL_START_ROW}-{AUTO_FILL_END_ROW}, Cols {AUTO_FILL_START_COL}-{AUTO_FILL_END_COL})",
            key="auto_fill_toggle_switch"
        )

        if auto_fill_toggle:
            st.session_state.start_row_manual_val = AUTO_FILL_START_ROW
            st.session_state.end_row_manual_val = AUTO_FILL_END_ROW
            st.session_state.start_col_manual_val = AUTO_FILL_START_COL
            st.session_state.end_col_manual_val = AUTO_FILL_END_COL
            st.session_state.use_header_manual_val = True

        if selection_method == "Manual Range Input":
            st.markdown("#### Manual Subtable Range Selection")
            st.info("Enter the row and column numbers as they appear in Excel (1-indexed).")
            
            start_row_manual = st.number_input(
                "Start Row", min_value=1, max_value=max_row, 
                value=st.session_state.start_row_manual_val, key="start_row_manual_input_key"
            )
            end_row_manual = st.number_input(
                "End Row", min_value=start_row_manual, max_value=max_row, 
                value=max(start_row_manual, st.session_state.end_row_manual_val), key="end_row_manual_input_key"
            )
            start_col_manual = st.number_input(
                "Start Column (A=1)", min_value=1, max_value=max_column, 
                value=st.session_state.start_col_manual_val, key="start_col_manual_input_key"
            )
            end_col_manual = st.number_input(
                "End Column", min_value=start_col_manual, max_value=max_column, 
                value=max(start_col_manual, st.session_state.end_col_manual_val), key="end_col_manual_input_key"
            )
            use_first_row_as_header_manual = st.checkbox(
                "Use first row of selection as header", 
                value=st.session_state.use_header_manual_val, key="use_header_manual_input_key"
            )

            st.session_state.start_row_manual_val = start_row_manual
            st.session_state.end_row_manual_val = end_row_manual
            st.session_state.start_col_manual_val = start_col_manual
            st.session_state.end_col_manual_val = end_col_manual
            st.session_state.use_header_manual_val = use_first_row_as_header_manual

            df_initial = get_initial_dataframe(wb, selected_sheet, 
                                                start_row_manual, end_row_manual, 
                                                start_col_manual, end_col_manual, 
                                                use_first_row_as_header_manual)

        elif selection_method == "Auto-Detect by Blank Rows":
            st.markdown("#### Auto-Detecting Subtables")
            uploaded_file.seek(0)
            full_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)

            non_empty_rows_indices = full_df.dropna(how='all').index.tolist()
            df_auto_detected = pd.DataFrame()

            if non_empty_rows_indices:
                first_data_row_idx_0based = non_empty_rows_indices[0]
                last_data_row_idx_0based = first_data_row_idx_0based
                for i in range(first_data_row_idx_0based + 1, len(full_df)):
                    if (i in non_empty_rows_indices) and ((i - 1) in non_empty_rows_indices):
                        last_data_row_idx_0based = i
                    else:
                        break
                
                sub_df_raw = full_df.iloc[first_data_row_idx_0based : last_data_row_idx_0based + 1].copy()
                non_empty_cols_indices = sub_df_raw.dropna(axis=1, how='all').columns.tolist()

                if non_empty_cols_indices:
                    detected_start_row = first_data_row_idx_0based + 1
                    detected_end_row = last_data_row_idx_0based + 1
                    detected_start_col = non_empty_cols_indices[0] + 1
                    detected_end_col = non_empty_cols_indices[-1] + 1
                    
                    st.info(f"Auto-detected range: Rows {detected_start_row} to {detected_end_row}, Columns {detected_start_col} to {detected_end_col}")

                    use_auto_detected_header = st.checkbox("Use the first row of the auto-detected selection as header?", value=True, key="use_auto_header")

                    if use_auto_detected_header:
                        auto_headers = sub_df_raw.iloc[0, non_empty_cols_indices].tolist()
                        auto_rows = sub_df_raw.iloc[1:, non_empty_cols_indices].values.tolist()
                    else:
                        auto_headers = [f"Column_{i+1}" for i in range(len(non_empty_cols_indices))]
                        auto_rows = sub_df_raw.iloc[:, non_empty_cols_indices].values.tolist()

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

                    adjusted_auto_rows = []
                    expected_cols_auto = len(headers)
                    for row in auto_rows:
                        if len(row) < expected_cols_auto:
                            adjusted_auto_rows.append(list(row) + [None] * (expected_cols_auto - len(row)))
                        else:
                            adjusted_auto_rows.append(list(row[:expected_cols_auto]))
                            
                    df_auto_detected = pd.DataFrame(adjusted_auto_rows, columns=headers)
                    df_auto_detected = df_auto_detected.dropna(how="all")

                    if 'Order' not in df_auto_detected.columns:
                        df_auto_detected.insert(0, 'Order', range(1, len(df_auto_detected) + 1))

                else:
                    st.warning("No contiguous data block found for auto-detection in columns within the detected rows.")
            else:
                st.warning("No non-empty rows found for auto-detection. The sheet might be entirely blank or formatted unusually.")

            df_initial = df_auto_detected

        # --- Session State Management for current_df and history ---
        current_data_selection_id = (
            f"{uploaded_file.file_id}-"
            f"{selected_sheet}-"
            f"{selection_method}-"
            f"{st.session_state.get('start_row_manual_val', '')}-"
            f"{st.session_state.get('end_row_manual_val', '')}-"
            f"{st.session_state.get('start_col_manual_val', '')}-"
            f"{st.session_state.get('end_col_manual_val', '')}-"
            f"{st.session_state.get('use_header_manual_val', '')}"
        )

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

        # --- Auto-remove specific rows ---
        if not st.session_state.current_df.empty and 'Order' in st.session_state.current_df.columns:
            st.markdown("### 🗑️ Automatic Row Filtering")
            auto_remove_toggle = st.checkbox(
                f"Automatically remove rows with 'Order' numbers: {', '.join(map(str, ROWS_TO_AUTO_REMOVE))}",
                key="auto_remove_rows_toggle"
            )

            if auto_remove_toggle:
                original_row_count = len(st.session_state.current_df)
                
                df_temp = st.session_state.current_df.copy()
                df_temp['Order_numeric'] = pd.to_numeric(df_temp['Order'], errors='coerce')

                rows_to_keep_mask = ~df_temp['Order_numeric'].isin(ROWS_TO_AUTO_REMOVE)
                
                if not rows_to_keep_mask.all():
                    st.session_state.history.append(st.session_state.current_df.copy())
                    
                    st.session_state.current_df = df_temp[rows_to_keep_mask].drop(columns=['Order_numeric']).reset_index(drop=True)
                    
                    removed_count = original_row_count - len(st.session_state.current_df)
                    st.success(f"Automatically removed {removed_count} row(s) based on predefined order numbers.")
                    st.rerun()
            st.markdown("---")

        # --- Auto-combine specific rows ---
        if not st.session_state.current_df.empty and 'Order' in st.session_state.current_df.columns and len(ROWS_TO_AUTO_COMBINE) > 1:
            st.markdown("### 🔗 Automatic Row Combination")
            auto_combine_toggle = st.checkbox(
                f"Automatically combine rows with 'Order' numbers: {', '.join(map(str, ROWS_TO_AUTO_COMBINE))} and rename to '{AUTO_COMBINED_ROW_NAME}'",
                key="auto_combine_rows_toggle"
            )

            if auto_combine_toggle:
                df_temp_combine = st.session_state.current_df.copy()
                df_temp_combine['Order_numeric_combine'] = pd.to_numeric(df_temp_combine['Order'], errors='coerce')
                indices_to_combine = df_temp_combine[df_temp_combine['Order_numeric_combine'].isin(ROWS_TO_AUTO_COMBINE)].index.tolist()

                if len(indices_to_combine) >= 2:
                    st.session_state.history.append(st.session_state.current_df.copy())

                    combined_row_data = {}
                    selected_df_for_auto_combine = st.session_state.current_df.loc[indices_to_combine]

                    for col_idx, col in enumerate(st.session_state.current_df.columns):
                        if pd.api.types.is_numeric_dtype(st.session_state.current_df[col]):
                            combined_row_data[col] = selected_df_for_auto_combine[col].sum()
                        else:
                            joined_value = " / ".join(selected_df_for_auto_combine[col].dropna().astype(str).tolist())
                            combined_row_data[col] = joined_value

                    target_name_col = None
                    if 'Order' in st.session_state.current_df.columns and len(st.session_state.current_df.columns) > 1:
                        first_non_order_col = next((col for col in st.session_state.current_df.columns if col != 'Order'), None)
                        if first_non_order_col:
                            target_name_col = first_non_order_col
                    elif len(st.session_state.current_df.columns) > 0:
                        target_name_col = st.session_state.current_df.columns[0]
                    
                    if target_name_col:
                        combined_row_data[target_name_col] = AUTO_COMBINED_ROW_NAME

                    combined_df_new = pd.DataFrame([combined_row_data], columns=st.session_state.current_df.columns)
                    
                    remaining_df = st.session_state.current_df.drop(index=indices_to_combine).reset_index(drop=True)
                    st.session_state.current_df = pd.concat([remaining_df, combined_df_new], ignore_index=True)

                    st.success(f"Automatically combined rows with Order {', '.join(map(str, ROWS_TO_AUTO_COMBINE))} into '{AUTO_COMBINED_ROW_NAME}'.")
                    st.rerun()
                else:
                    st.warning(f"Could not auto-combine. Found {len(indices_to_combine)} row(s) with order numbers {', '.join(map(str, ROWS_TO_AUTO_COMBINE))}. At least 2 are required.")
            st.markdown("---")


        # --- Display and Editing UI ---
        if not st.session_state.current_df.empty:
            st.subheader("✏️ Edit Table and Prepare Operations")
            st.info("Directly edit values or delete rows in the table below. Use the selections to prepare for combined operations. Click 'Perform Selected Operations' to apply all changes.")

            # --- Using st.form for a single, consolidated submission ---
            # All widgets inside this form will update their values only when the submit button is pressed.
            with st.form(key="operations_form"):
                # Always start with the DataFrame from current session state for the editor display
                # Ensure 'Order' column is numeric for proper sorting for the data editor
                st.session_state.current_df['Order'] = pd.to_numeric(st.session_state.current_df['Order'], errors='coerce').fillna(0).astype(int)

                # The data editor. Its output `edited_df` will contain the user's manual edits
                # at the time of form submission.
                edited_df = st.data_editor(
                    st.session_state.current_df, # This is the source for the editor display
                    num_rows="dynamic",
                    use_container_width=True,
                    column_config={
                        "Order": st.column_config.NumberColumn(
                            "Order",
                            help="Assign a number to reorder rows.",
                            default=0,
                            step=1,
                            format="%d"
                        )
                    },
                    key="main_data_editor_in_form" # Unique key for editor inside form
                )

                # --- Row Combination Options (inside form) ---
                st.subheader("🔗 Row Combination Options")
                st.write("Current table row indices for selection:")
                # Display current indices from the edited_df for user reference
                # This will reflect the state at the time of form submission
                st.dataframe(edited_df.index.to_frame(name='Index'), use_container_width=True)
                st.info("Select rows by their current table index (leftmost column) for combination.")

                selected_rows_to_combine = st.multiselect(
                    "Select rows to combine (by current table index)",
                    edited_df.index.tolist(), # Populate based on editor's current state on submission
                    key="combine_rows_multiselect_in_form"
                )
                custom_name_for_combined_row = st.text_input("Custom name for the new combined row", value="Combined Row", key="custom_combined_row_name_in_form")

                # --- Column Merging Options (inside form) ---
                st.subheader("🧬 Column Merging Options")
                selected_cols_to_merge = st.multiselect("Select columns to merge", edited_df.columns.tolist(), key="merge_cols_multiselect_in_form")
                new_merged_col_name = st.text_input("New column name for merged data", value="MergedColumn", key="new_merged_col_name_input_in_form")
                
                st.markdown("---") # Separator before the button

                # The submit button for the form. This is the "Perform Selected Operations" button.
                submit_button = st.form_submit_button("Perform Selected Operations")

                if submit_button:
                    # All operations here will be triggered only on form submission
                    # Use edited_df directly as the starting point for processing
                    df_to_process = edited_df.copy() 
                    
                    # Save the current_df (state BEFORE this combined operation) to history
                    st.session_state.history.append(st.session_state.current_df.copy()) 
                    
                    messages = []

                    # 1. Apply New Row Order (based on 'Order' column edited in data_editor)
                    if 'Order' in df_to_process.columns:
                        temp_df = df_to_process.copy()
                        temp_df['Order_temp_sort'] = pd.to_numeric(temp_df['Order'], errors='coerce').fillna(0).astype(int)
                        
                        if temp_df['Order_temp_sort'].duplicated().any():
                            temp_df['Order_temp_sort_with_tiebreaker'] = temp_df['Order_temp_sort'].astype(str) + '.' + temp_df.groupby('Order_temp_sort').cumcount().astype(str)
                            temp_df['Order_temp_sort_with_tiebreaker'] = pd.to_numeric(temp_df['Order_temp_sort_with_tiebreaker'], errors='coerce')
                            df_to_process = temp_df.sort_values(by='Order_temp_sort_with_tiebreaker', ascending=True).drop(columns=['Order_temp_sort', 'Order_temp_sort_with_tiebreaker']).reset_index(drop=True)
                        else:
                            df_to_process = temp_df.sort_values(by='Order_temp_sort', ascending=True).drop(columns=['Order_temp_sort']).reset_index(drop=True)

                        messages.append("Rows reordered successfully (based on 'Order' column).")
                    else:
                        messages.append("Skipped row reordering: No 'Order' column found.")

                    # 2. Combine Selected Rows (operate on df_to_process which might be reordered)
                    if selected_rows_to_combine:
                        valid_indices_to_combine = [idx for idx in selected_rows_to_combine if idx in df_to_process.index]

                        if len(valid_indices_to_combine) >= 2:
                            combined_row_data = {}
                            selected_df_for_combine = df_to_process.loc[valid_indices_to_combine]

                            for col_name in df_to_process.columns:
                                if pd.api.types.is_numeric_dtype(df_to_process[col_name]):
                                    combined_row_data[col_name] = selected_df_for_combine[col_name].sum()
                                else:
                                    joined_value = " / ".join(selected_df_for_combine[col_name].dropna().astype(str).tolist())
                                    combined_row_data[col_name] = joined_value

                            target_name_col = None
                            if 'Order' in df_to_process.columns and len(df_to_process.columns) > 1:
                                target_name_col = next((col for col in df_to_process.columns if col != 'Order'), None)
                            elif len(df_to_process.columns) > 0:
                                target_name_col = df_to_process.columns[0]
                            
                            if target_name_col:
                                combined_row_data[target_name_col] = custom_name_for_combined_row

                            combined_df = pd.DataFrame([combined_row_data], columns=df_to_process.columns)
                            
                            remaining_df = df_to_process.drop(index=valid_indices_to_combine).reset_index(drop=True)
                            df_to_process = pd.concat([remaining_df, combined_df], ignore_index=True)
                            messages.append(f"Rows combined successfully into '{custom_name_for_combined_row}'.")
                        else:
                            messages.append(f"Skipped row combination: Need at least 2 selected rows, found {len(valid_indices_to_combine)} valid selected rows.")
                    else:
                        messages.append("Skipped row combination: No rows selected to combine.")


                    # 3. Merge Selected Columns (operate on df_to_process which might be reordered/combined)
                    if selected_cols_to_merge and len(selected_cols_to_merge) >= 2:
                        existing_non_selected_cols = [col for col in df_to_process.columns if col not in selected_cols_to_merge]
                        if new_merged_col_name in existing_non_selected_cols:
                            messages.append(f"Skipped column merge: Column '{new_merged_col_name}' already exists and is not part of the merge selection. Please choose a different name.")
                        else:
                            df_to_process[new_merged_col_name] = (
                                df_to_process[selected_cols_to_merge]
                                .astype(str)
                                .agg(lambda x: " / ".join(x.dropna()), axis=1)
                            )
                            df_to_process.drop(columns=selected_cols_to_merge, inplace=True)
                            messages.append(f"Columns merged into '{new_merged_col_name}'.")
                    else:
                        messages.append("Skipped column merge: Please select at least two columns to merge.")
                    
                    # Update the main DataFrame in session state with the final result of all operations
                    st.session_state.current_df = df_to_process 
                    
                    for msg in messages: # Display all messages
                        st.info(msg)
                    st.success("Operations completed!") # Success message
                    st.rerun() # Rerun once after all operations
            # End of st.form

            # Undo button is outside the form, operating on the main session state
            if st.button("Undo Last Action", key="undo_button_outside_form"):
                if st.session_state.history:
                    st.session_state.current_df = st.session_state.history.pop()
                    st.success("Undo successful. Table restored to previous state.")
                    st.rerun()
                else:
                    st.warning("No previous state to undo. History is empty.")

            st.subheader("📋 Final Edited Table")
            final_df = st.session_state.current_df.dropna(how="all").reset_index(drop=True)
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
                file_name="modified_subtable.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.info("No data found for the selected range/auto-detection. Please adjust your selection or upload a different file.")

    except Exception as e:
        st.error(f"An unexpected error occurred while processing the Excel file: {e}")
        st.exception(e)
        st.info("Please ensure it's a valid Excel file with readable content and try again.")
else:
    st.info("Please upload an Excel file to begin.")
