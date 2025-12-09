# Import Libraries 
import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(
    page_title="Aftermath App",
    page_icon="⎘",
    layout="wide",
)

# -----------------------------
# Core functions
# -----------------------------
@st.cache_data
def cleanInvoice(invoice_excel, row, num_cols, skip_last_rows=False, skip_num=0):
    """
    Generalized Excel cleaning function for Aftermath invoices.

    Parameters:
    - invoice_excel: Excel file object.
    - row: The row number in Excel where you want the first heading.
    - num_cols: The number of columns there should be.
    - skip_last_rows: True or False. Are there any rows to skip at the end during import? (optional)
    - skip_num: The row number you want the Excel to end on. (optional)
    """
    
    if skip_last_rows == True:

        xls = pd.ExcelFile(invoice_excel)
        sheet_name = xls.sheet_names[0]
        
        wb = openpyxl.load_workbook(invoice_excel, read_only=True)
        ws = wb[sheet_name]
        total_rows = ws.max_row

        last_row_of_data = skip_num + 1

        skipfooter = total_rows - last_row_of_data
        
        clean_excel = pd.read_excel(invoice_excel, skiprows=row - 1, skipfooter=skipfooter) 

    else:

        clean_excel = pd.read_excel(invoice_excel, skiprows=row - 1) 
    
    # Standardize column headers
    clean_excel.columns = clean_excel.columns.str.strip()
    clean_excel.columns = clean_excel.columns.str.replace(r'\s+', ' ', regex=True)
    clean_excel.columns = clean_excel.columns.astype(str)
    clean_excel.columns = clean_excel.columns.str.replace(r'\xa0|\t', ' ', regex=True).str.strip()

    column_names = clean_excel.columns.tolist()
    filtered_columns = [col for col in column_names if "Unnamed" not in col]  

    # Drop rows that contain "Totals" anywhere
    clean_excel = clean_excel[~clean_excel.apply(
        lambda row: row.astype(str).str.contains('Totals', case=False, na=False)
    ).any(axis=1)]

    # Drop fully empty rows
    clean_excel = clean_excel.dropna(axis=0, how='all')

    # Forward fill date column
    clean_excel.loc[:, 'Date:'] = clean_excel['Date:'].ffill()
    
    # Parse and filter valid dates
    clean_excel.loc[:, 'Date:'] = pd.to_datetime(
        clean_excel['Date:'], format='%Y/%m/%d', errors='coerce'
    ).dt.date
    clean_excel = clean_excel.dropna(subset=['Date:'])
    clean_excel.reset_index(drop=True, inplace=True)

    # Realign row values to the left
    def realign_row(a_row):
        non_empty_values = a_row.dropna().values
        aligned_row = pd.Series([None] * len(clean_excel.columns), index=clean_excel.columns)
        aligned_row[:len(non_empty_values)] = non_empty_values
        return aligned_row

    clean_excel = clean_excel.apply(realign_row, axis=1)

    # Limit to expected number of columns
    clean_excel = clean_excel.drop(clean_excel.columns[num_cols:], axis=1)

    clean_excel.columns = filtered_columns

    return clean_excel

@st.cache_data
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

@st.cache_data
def cleanMonitorData(monitor_excel):
    """
    Excel cleaning function for monitor data headers.

    Parameters:
    - monitor_excel: Excel file object.
    """

    excel_2 = pd.read_excel(monitor_excel)

    excel_2.columns = excel_2.columns.str.strip()
    excel_2.columns = excel_2.columns.str.replace(r'\s+', ' ', regex=True)
    excel_2.columns = excel_2.columns.astype(str)

    return excel_2

@st.cache_data
def compare_dfs(excel_1, pair1, compare1, excel_2, pair2, compare2):
    """
    Excel comparison function for Aftermath workflow.

    Parameters:
    - excel_1, excel_2: DataFrames to compare.
    - pair1, pair2: Columns to match between DataFrames.
    - compare1, compare2: Columns to compare values for matched IDs. (optional)
    
    Returns:
    - missing_from_excel_1: Rows in excel_1 not found in excel_2 based on pair columns.
    - missing_from_excel_2: Rows in excel_2 not found in excel_1 based on pair columns.
    - diff_qty_df: DataFrame of mismatched values in compare1/compare2.
    - combo: DataFrame of unmatched IDs between both Excels.
    """

    filter1 = excel_1[pair1].isin(excel_2[pair2])
    missing_from_excel_2 = excel_1[~filter1]

    filter2 = excel_2[pair2].isin(excel_1[pair1])
    missing_from_excel_1 = excel_2[~filter2]

    part1 = excel_2[pair2].unique().tolist()
    part2 = excel_1[pair1].unique().tolist()
    id_list = set(part1 + part2)

    diff_qty = {}
    combined_list = []

    for i in id_list:
        excel_1_match = excel_1.loc[excel_1[pair1] == i, compare1]
        excel_2_match = excel_2.loc[excel_2[pair2] == i, compare2]
    
        if not excel_1_match.empty and not excel_2_match.empty:
            qty1 = excel_1_match.values[0]
            qty2 = excel_2_match.values[0]
        
            if qty1 != qty2:
                diff_qty[i] = [qty1, qty2]
        else:
            combined_list.append(i)

    diff_qty_df = pd.DataFrame(diff_qty)
    diff_qty_df.index = ['excel_1', 'excel_2']

    combo = pd.DataFrame(combined_list, columns=['Missing list'])
    combo = combo.sort_values(by='Missing list', ascending=False)

    return missing_from_excel_1, missing_from_excel_2, diff_qty_df, combo

# -----------------------------
# Header and sidebar
# -----------------------------
col1, col2, col3 = st.columns([1, 6, 1])

with col2:
    st.image("logo.png")

st.markdown(
    "<h1 style='text-align: center'>Aftermath Excel Cleaning and Comparison App</h1>",
    unsafe_allow_html=True,
)
st.write("")

with st.sidebar:
    st.markdown("### Navigation")
    st.caption("Use the menu above to switch between pages.")
    st.caption("You are on the **Aftermath App** page. Use this for Aftermath invoices and monitor data.")

    with st.expander("ℹ️ How this page works", expanded=False):
        st.markdown(
            """
1. Upload the **invoice Excel** that was converted from PDF.  
2. Answer a few questions so the headings and rows line up correctly.  
3. Upload the **monitor data Excel**.  
4. Review missing tickets and any differences in quantities.  
5. Download the combined Excel file at the bottom of the page.
            """
        )

# -----------------------------
# Intro text
# -----------------------------
st.markdown(
    """
This web application is tailored for Aftermath Disaster Recovery. It processes invoice and monitor Excel files, identifies missing ticket numbers in each file, and highlights ticket numbers that appear in both files but have different values in a key column.

One Excel file comes from an invoice that was originally a PDF. This file often needs cleaning to remove extra rows, align headers, and standardize dates. The other Excel file is the monitor data, which is usually clean apart from minor header adjustments.

👈👈 Use the sidebar if you would like to open the **:rainbow[tutorial]** or switch to the more general app.
"""
)

st.write("")
st.markdown(
    "<h2 style='text-align: center'>Excel uploads and questions</h2>",
    unsafe_allow_html=True,
)
st.write("")

st.markdown(
    """
In this section, you will upload the invoice Excel file and answer a few questions to help the app clean it. After cleaning, you can download the invoice Excel file. Next, you will upload the monitor data Excel file and can also download the cleaned version. Finally, you will review the comparison results and download an aggregated Excel file that contains all outputs.
"""
)

# -----------------------------
# Main content
# -----------------------------
col1, col2, col3 = st.columns([1, 6, 1])

with col1:
    st.write("")

with col2:

    st.markdown("#### Questions about the invoice Excel (from PDF)")

    # Ask if the Excel is already clean
    isClean = st.radio(
        "Does this invoice Excel already look clean and ready to compare?",
        [":rainbow[Yes]", "No"],
        index=None,
        captions=[
            "The Excel is already clean",
            "It needs to be cleaned",
        ],
    )

    # Stop if not answered yet
    if isClean is None:
        st.warning("Please select one of the options above to continue.")
        st.stop()

    # --- If it is clean ---
    if isClean == ":rainbow[Yes]":
        uploaded_file = st.file_uploader("Upload the invoice Excel file", key="file1")

        if uploaded_file is not None:
            df1 = pd.read_excel(uploaded_file)
            st.session_state.df1 = df1
            st.dataframe(df1, hide_index=True)

            st.write("After checking the preview above, confirm that it still looks clean. If not, you can restart and choose the cleaning option instead.")
        else:
            st.warning("Please upload an invoice Excel file.")

    # --- If it needs cleaning ---
    else:
        st.write("Please answer the following questions about the invoice layout.")

        row = st.number_input(
            "What row number contains the first header you want to keep?",
            min_value=0,
            step=1,
            placeholder="Row number..."
        )

        skipEnd = st.radio(
            "Do you want to remove any extra rows at the bottom? For example, rows that contain notes or other unneeded data. \n\n"
            "**Note:** You do not need to remove rows if they only contain totals and the underlying data is the same.",
            [":rainbow[Yes]", "No"],
            index=None,
            captions=[
                "I need to remove rows at the end",
                "I do **not** need to remove rows at the end."
            ]
        )

        # Required variable
        skip_num = None

        if skipEnd is None:
            st.warning("Please select if you want to skip rows at the end.")
            st.stop()

        if skipEnd == ":rainbow[Yes]":
            skip_num = st.number_input(
                "What is the final row number you want to keep in the invoice Excel?",
                min_value=1,
                step=1,
                placeholder="Final row number..."
            )

        # Final checks
        ready = row is not None and row >= 0

        if skipEnd == ":rainbow[Yes]":
            ready = ready and skip_num is not None and skip_num >= 1

        if not ready:
            st.warning("Please answer all questions above before uploading the file.")
            st.stop()

        uploaded_file = st.file_uploader("Upload the invoice Excel file", key="file1")

        if uploaded_file is not None:
            if skipEnd == ":rainbow[Yes]":

                df1 = cleanInvoice(uploaded_file, 
                                   row, 
                                   8, 
                                   skip_last_rows=True, 
                                   skip_num=skip_num)

                excel_bytes = to_excel(df1)

                # Create download button
                st.download_button(
                    label="📥 Download Cleaned Invoice Excel",
                    data=excel_bytes,
                    file_name='cleaned_invoice.xlsx'
                    )

            else:
                df1 = cleanInvoice(uploaded_file, row, 8)

            st.session_state.df1 = df1
            st.dataframe(df1, hide_index=True)

            st.write(
                "After checking the cleaned invoice above, confirm that it looks correct. "
                "If something is off, you can adjust your answers and run the cleaning step again."
            )

            excel_bytes = to_excel(df1)

            st.download_button(
                label="📥 Download cleaned invoice Excel",
                data=excel_bytes,
                file_name="cleaned_invoice.xlsx",
            )
        else:
            st.warning("Please upload an invoice Excel file.")

    st.markdown("#### Upload the monitor data Excel")
    st.write("This assumes the monitor data Excel is already mostly clean and only needs light header cleaning.")

    uploaded_file = st.file_uploader("Upload the monitor data Excel file", key="file2")

    if uploaded_file is not None:
        df2 = cleanMonitorData(uploaded_file)
        st.session_state.df2 = df2
        st.dataframe(df2, hide_index=True)

        st.write("After checking, confirm that the monitor data looks correct. If not, you may need to adjust the source file.")
        excel_bytes = to_excel(df2)

        st.download_button(
            label="📥 Download cleaned monitor Excel",
            data=excel_bytes,
            file_name="cleaned_monitor_data.xlsx",
        )
    else:
        st.warning("Please upload the monitor data Excel file.")

    st.markdown("#### Comparison results")
    with st.expander("Click here to view or hide comparison results", expanded=True):

        if 'df1' in st.session_state and 'df2' in st.session_state:
            df1 = st.session_state.df1
            df2 = st.session_state.df2

            missing_from_excel_1, missing_from_excel_2, diff_qty_df, combo = compare_dfs(
                df1,
                'FEMA Ticket #',
                'Calculated Qty',
                df2,
                'Ticket Number',
                'Quantity'
            )

            combo_sorted = combo.sort_values(by='Missing list')

            st.write("**Entries missing from Excel 1 (invoice Excel)**")
            st.session_state.missing_from_excel_1 = missing_from_excel_1
            st.dataframe(missing_from_excel_1, hide_index=True)

            st.write("**Entries missing from Excel 2 (monitor Excel)**")
            st.session_state.missing_from_excel_2 = missing_from_excel_2
            st.dataframe(missing_from_excel_2, hide_index=True)

            st.write("**Entries with different quantities between the two Excels**")
            st.session_state.diff_qty_df = diff_qty_df
            st.dataframe(diff_qty_df, hide_index=True)

            st.write("**Combined list of ticket IDs missing from one or both Excels**")
            st.session_state.combo = combo
            st.session_state.combo_sorted = combo_sorted
            st.data_editor(combo_sorted, hide_index=True)

        else:
            st.info("Please upload both the invoice Excel and the monitor Excel to run the comparison.")

    st.markdown("#### Download all DataFrames into a single Excel file")

    if (
        'missing_from_excel_1' in st.session_state and
        'missing_from_excel_2' in st.session_state and
        'diff_qty_df' in st.session_state and
        'combo' in st.session_state
    ):
        missing_from_excel_1 = st.session_state.missing_from_excel_1
        missing_from_excel_2 = st.session_state.missing_from_excel_2
        diff_qty_df = st.session_state.diff_qty_df
        combo = st.session_state.combo

        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            placeholder_df = pd.DataFrame({'Notice': ['No data available']})
            placeholder_df.to_excel(writer, sheet_name='No_Data', index=False)

            sheets_written = 0

            if not missing_from_excel_1.empty:
                missing_from_excel_1.to_excel(
                    writer,
                    sheet_name='missing_from_excel_1',
                    index=False
                )
                sheets_written += 1

            if not missing_from_excel_2.empty:
                missing_from_excel_2.to_excel(
                    writer,
                    sheet_name='missing_from_excel_2',
                    index=False
                )
                sheets_written += 1

            if not diff_qty_df.empty:
                diff_qty_df.to_excel(
                    writer,
                    sheet_name='diff_qty_df',
                    index=True
                )
                sheets_written += 1

            if not combo.empty:
                combo.to_excel(
                    writer,
                    sheet_name='combo_missing_id',
                    index=False
                )
                sheets_written += 1

            if sheets_written > 0:
                del writer.book['No_Data']

        output.seek(0)

        st.download_button(
            label="📥 Download combined comparison Excel",
            data=output,
            file_name="aftermath_comparison_results.xlsx",
        )
    else:
        st.info("Run the comparison above to enable the download button.")

with col3:
    st.write("")
