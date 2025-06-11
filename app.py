import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Excel Subtable Editor", layout="wide")
st.title("📊 Excel Subtable Editor")

# Load the Excel file
excel_file = "Earned Comm Breakdown (Finance)_Apr 2025(Summary).xlsx"
wb = openpyxl.load_workbook(excel_file, data_only=True)
sheet_names = wb.sheetnames

# Select sheet
selected_sheet = st.selectbox("Select a sheet", sheet_names)
ws = wb[selected_sheet]

# Get sheet dimensions
max_row = ws.max_row
max_col = ws.max_column

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

if use_first_row_as_header and len(data) > 1:
    raw_headers = list(data[0])
    rows = data[1:]
else:
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

df = pd.DataFrame(rows, columns=headers)
df = df.dropna(how="all")

# Initialize session state
if "current_df" not in st.session_state:
    st.session_state.current_df = df.copy()
if "history" not in st.session_state:
    st.session_state.history = []

st.subheader("✏️ Edit Table")
edited_df = st.data_editor(st.session_state.current_df, num_rows="dynamic", use_container_width=True)

if st.button("Save Changes"):
    st.session_state.history.append(st.session_state.current_df.copy())
    st.session_state.current_df = edited_df.copy()
    st.success("Changes saved.")

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
                combined_row[col] = " / ".join(selected_df[col].astype(str))
        if st.session_state.current_df.columns.size > 0:
            combined_row[st.session_state.current_df.columns[0]] = custom_name
        st.session_state.history.append(st.session_state.current_df.copy())
        st.session_state.current_df = st.session_state.current_df.drop(index=selected_rows)
        st.session_state.current_df = pd.concat([st.session_state.current_df, pd.DataFrame([combined_row])], ignore_index=True)
        st.success("Rows combined successfully.")

st.subheader("🧬 Merge Columns")
selected_cols = st.multiselect("Select columns to merge", st.session_state.current_df.columns.tolist(), key="merge_cols")
new_col_name = st.text_input("New column name", value="MergedColumn")
if st.button("Merge Selected Columns"):
    if selected_cols and len(selected_cols) >= 2:
        st.session_state.history.append(st.session_state.current_df.copy())
        st.session_state.current_df[new_col_name] = st.session_state.current_df[selected_cols].astype(str).agg(" / ".join, axis=1)
        st.session_state.current_df.drop(columns=selected_cols, inplace=True)
        st.success(f"Columns merged into '{new_col_name}'")

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
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ModifiedTable")
    output.seek(0)
    return output

excel_data = to_excel(final_df)
st.download_button("Download as Excel", data=excel_data, file_name="modified_table.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
