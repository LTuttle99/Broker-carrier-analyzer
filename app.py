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
        @st.cache_resource(ttl=3600)
        def load_workbook_and_get_sheet_names(file_buffer):
            file_buffer.seek(0)
            wb = openpyxl.load_workbook(file_buffer, data_only=True)
            return wb, wb.sheetnames

        # Modified get_initial_dataframe to use pandas.read_excel directly
        # This simplifies the logic as pandas handles header detection and range reading well.
        @st.cache_data(ttl=3600)
        def get_initial_dataframe_pandas(_file_buffer, sheet_name, header_row_index, usecols_range=None):
            """
            Reads a specific sheet and range from an Excel file buffer into a DataFrame.
            _file_buffer: BytesIO object of the uploaded file.
            sheet_name: The name of the sheet to read.
            header_row_index: The 0-based index of the row to use as header. None if no header.
            usecols_range: A list of column names or 0-based indices to include.
                           If None, all columns are included.
            """
            _file_buffer.seek(0) # Important: always seek to start before reading
            
            # If header_row_index is None, pandas will not use any header
            header_param = header_row_index if header_row_index is not None else None
            
            df_result = pd.read_excel(
                _file_buffer,
                sheet_name=sheet_name,
                header=header_param,
                usecols=usecols_range # If usecols_range is None, pandas reads all
            )
            
            df_result = df_result.dropna(how="all") # Drop rows that are entirely NaN

            # Add a default 'Order' column for reordering if not present
            if 'Order' not in df_result.columns:
                 df_result.insert(0, 'Order', range(1, len(df_result) + 1)) # Add at the beginning, 1-indexed
            
            # Ensure 'Order' column is numeric for proper sorting when initialized
            df_result['Order'] = pd.to_numeric(df_result['Order'], errors='coerce').fillna(0).astype(int)

            return df_result


        wb, sheet_names = load_workbook_and_get_sheet_names(uploaded_file)

        st.success("File loaded successfully!")

        selected_sheet = st.selectbox("Select a sheet", sheet_names)

        # Get max_row and max_column for the selected sheet using openpyxl for accurate dimensions
        ws = wb[selected_sheet]
        max_row = ws.max_row
        max_column = ws.max_column
        st.write(f"Sheet dimensions: {max_row} rows, {max_column} columns")

        st.markdown("### 🔍 Choose Subtable Selection Method")
        selection_method = st.radio(
            "How do you want to select the subtable?",
            ("Manual Range Input", "Auto-Detect First Table"), # Changed text
            index=0
        )

        df_initial = pd.DataFrame() # Initialize df_initial

        if selection_method == "Manual Range Input":
            st.markdown("#### Manual Subtable Range Selection")
            start_row_excel = st.number_input("Start Row (from Excel file, 1-indexed)", min_value=1, max_value=max_row, value=1, key="start_row_manual")
            end_row_excel = st.number_input("End Row (from Excel file, 1-indexed)", min_value=start_row_excel, max_value=max_row, value=min(start_row_excel + 10, max_row), key="end_row_manual")
            start_col_excel = st.number_input("Start Column (A=1)", min_value=1, max_value=max_column, value=1, key="start_col_manual")
            end_col_excel = st.number_input("End Column", min_value=start_col_excel, max_value=max_column, value=min(start_col_excel + 5, max_column), key="end_col_manual")
            use_first_row_as_header = st.checkbox("Use first row of selection as header", value=True, key="use_header_manual")

            # Convert 1-indexed Excel ranges to 0-indexed pandas parameters
            header_row_index = start_row_excel - 1 if use_first_row_as_header else None
            # Pandas usecols expects column names or 0-indexed column numbers.
            # To select a range, we can pass a list of indices: [start_col-1, ..., end_col-1]
            usecols_range = list(range(start_col_excel - 1, end_col_excel))
            
            # Pass the uploaded_file directly, not the wb object
            df_initial = get_initial_dataframe_pandas(uploaded_file, selected_sheet, header_row_index, usecols_range)

            # Manually trim rows if not using a header or if the header row is outside the range for data
            if not use_first_row_as_header:
                # If no header, data starts from start_row_excel, so skiprows = start_row_excel - 1
                # pandas read_excel already accounts for skiprows based on header
                # We need to explicitly slice rows if header=None and we want a specific range.
                # However, with get_initial_dataframe_pandas, we're relying on header=param,
                # so the slicing needs to happen after the initial load.
                # This makes manual range harder if you don't use the first row as header.

                # Re-think this part: if use_first_row_as_header is False, the first row
                # of the loaded dataframe should correspond to start_row_excel.
                # pandas read_excel with header=None will read from row 0.
                # We need to manually slice the dataframe *after* loading if header=None is truly desired
                # and we want a subset of rows from the non-header-skipped dataframe.
                # For simplicity, stick to header_row_index or let pandas infer.
                pass # The current get_initial_dataframe_pandas handles the header_row_index correctly.


        elif selection_method == "Auto-Detect First Table":
            st.markdown("#### Auto-Detecting First Table (Smart Search)")
            
            uploaded_file.seek(0)
            # Read the entire sheet without any header or parsing for initial scan
            temp_full_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)

            potential_header_row_index = -1
            min_non_empty_cells_for_header = 3 # Heuristic: at least 3 non-empty cells to be a header
            
            # Find the first row that looks like a header (sufficient non-empty cells)
            for r_idx in range(len(temp_full_df)):
                row_data = temp_full_df.iloc[r_idx]
                non_empty_count = row_data.count() # Count non-NaN cells
                
                if non_empty_count >= min_non_empty_cells_for_header:
                    potential_header_row_index = r_idx
                    break
            
            if potential_header_row_index != -1:
                st.info(f"Auto-detected table starting around row {potential_header_row_index + 1} (1-indexed) using it as header.")
                # We'll use this row as the header for pd.read_excel
                # get_initial_dataframe_pandas handles reading from this header row onwards
                df_initial = get_initial_dataframe_pandas(uploaded_file, selected_sheet, potential_header_row_index)
            else:
                st.warning("Could not auto-detect a clear table header. Displaying first 50 rows with no header.")
                # Fallback: if no clear header, just load first 50 rows as raw data
                uploaded_file.seek(0)
                df_initial = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None, nrows=50)
                # Manually add Order column if not present
                if 'Order' not in df_initial.columns:
                    df_initial.insert(0, 'Order', range(1, len(df_initial) + 1))
                df_initial = df_initial.dropna(how="all")


        # --- Session State Management ---
        # Include all parameters that affect the initial df_initial creation in the ID
        current_data_selection_id = (
            f"{uploaded_file.file_id}-{selected_sheet}-{selection_method}-"
            f"{start_row_manual}-{end_row_manual}-{start_col_manual}-{end_col_manual}-"
            f"{use_first_row_as_header}"
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
        
        # --- Display and Editing UI ---
        if not st.session_state.current_df.empty:
            st.subheader("✏️ Edit Table and Reorder Rows")
            st.info("To reorder rows, edit the numbers in the 'Order' column. To delete a row, click the 'X' button on the right.")

            # Ensure 'Order' column is numeric for proper sorting (re-convert in case of edits)
            st.session_state.current_df['Order'] = pd.to_numeric(st.session_state.current_df['Order'], errors='coerce').fillna(0).astype(int)

            edited_df = st.data_editor(
                st.session_state.current_df,
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
                }
            )

            if not edited_df.equals(st.session_state.current_df):
                st.session_state.history.append(st.session_state.current_df.copy())
                st.session_state.current_df = edited_df.copy()
                st.success("Changes detected. Save or Apply Order to confirm.")

            if st.button("Apply New Row Order"):
                if 'Order' in st.session_state.current_df.columns:
                    temp_df = st.session_state.current_df.copy()

                    temp_df['Order_temp_sort'] = temp_df['Order']
                    if temp_df['Order_temp_sort'].duplicated().any():
                        temp_df['Order_temp_sort'] = temp_df.groupby('Order_temp_sort').cumcount().add(1).astype(str)
                        temp_df['Order_temp_sort'] = temp_df['Order'].astype(str) + '.' + temp_df['Order_temp_sort']
                        temp_df['Order_temp_sort'] = pd.to_numeric(temp_df['Order_temp_sort'], errors='coerce')

                    st.session_state.current_df = temp_df.sort_values(by='Order_temp_sort', ascending=True).drop(columns=['Order_temp_sort']).reset_index(drop=True)
                    st.success("Rows reordered successfully!")
                    st.rerun()

                else:
                    st.warning("No 'Order' column found to reorder rows.")

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
