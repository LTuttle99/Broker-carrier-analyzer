import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import copy

st.set_page_config(page_title="Excel Table Editor", layout="wide")
st.title("📊 Excel Named Table Editor")

# Session state for undo functionality
if "history" not in st.session_state:
    st.session_state.history = []

def combine_rows(df, selected_indices, custom_name):
    if not selected_indices:
        return df

    selected_rows = df.loc[selected_indices]
    combined_row = {}

    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            combined_row[col] = selected_rows[col].sum()
        else:
            combined_row[col] = " / ".join(selected_rows[col].astype(str))

    # Replace first column value with custom name
    if df.columns.size > 0:
        combined_row[df.columns[0]] = custom_name

    df = df.drop(index=selected_indices)
    df = pd.concat([df, pd.DataFrame([combined_row])], ignore_index=True)
    return df

def merge_columns(df, selected_columns, new_column_name):
    if not selected_columns or len(selected_columns) < 2:
        return df

    df[new_column_name] = df[selected_columns].astype(str).agg(" / ".join, axis=1)
    df = df.drop(columns=selected_columns)
    return df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='ModifiedTable')
    output.seek(0)
    return output

def remove_empty_rows(df):
    return df.dropna(how='all')

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet_names = wb.sheetnames
    selected_sheet = st.selectbox("Select a sheet", sheet_names)
    ws = wb[selected_sheet]
    table_names = list(ws.tables.keys())

    if table_names:
        selected_table = st.selectbox("Select a named table", table_names)
        table = ws.tables[selected_table]
        table_range = table.ref
        data = ws[table_range]
        data = [[cell.value for cell in row] for row in data]
        df = pd.DataFrame(data[1:], columns=data[0])
        df = remove_empty_rows(df)

        if "current_df" not in st.session_state:
            st.session_state.current_df = df.copy()

        st.subheader("✏️ Edit Table")
        edited_df = st.data_editor(st.session_state.current_df, num_rows="dynamic", use_container_width=True)

        # Save current state for undo
        if st.button("Save Changes"):
            st.session_state.history.append(copy.deepcopy(st.session_state.current_df))
            st.session_state.current_df = edited_df.copy()
            st.success("Changes saved.")

        st.subheader("🔗 Combine Rows")
        selected_rows = st.multiselect("Select rows to combine (by index)", st.session_state.current_df.index.tolist())
        custom_name = st.text_input("Custom name for the new combined row", value="Combined Row")
        if st.button("Combine Selected Rows"):
            st.session_state.history.append(copy.deepcopy(st.session_state.current_df))
            st.session_state.current_df = combine_rows(st.session_state.current_df, selected_rows, custom_name)
            st.success("Rows combined successfully!")

        st.subheader("🧬 Merge Columns")
        selected_cols = st.multiselect("Select columns to merge", st.session_state.current_df.columns.tolist(), key="merge_cols")
        new_col_name = st.text_input("New column name", value="MergedColumn")
        if st.button("Merge Selected Columns"):
            st.session_state.history.append(copy.deepcopy(st.session_state.current_df))
            st.session_state.current_df = merge_columns(st.session_state.current_df, selected_cols, new_col_name)
            st.success(f"Columns merged into '{new_col_name}'")

        if st.button("Undo Last Action"):
            if st.session_state.history:
                st.session_state.current_df = st.session_state.history.pop()
                st.success("Undo successful.")
            else:
                st.warning("No previous state to undo.")

        st.subheader("📋 Final Table")
        final_df = remove_empty_rows(st.session_state.current_df)
        st.dataframe(final_df, use_container_width=True)

        st.subheader("📥 Download Modified Table")
        excel_data = to_excel(final_df)
        st.download_button("Download as Excel", data=excel_data, file_name="modified_table.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("No named tables found in the selected sheet.")
else:
    st.info("Please upload an Excel file to begin.")
