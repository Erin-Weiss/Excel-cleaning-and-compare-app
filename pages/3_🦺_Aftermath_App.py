# Import Libraries 
import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# Create code functions
@st.cache_data
def cleanInvoice(invoice_excel, row, num_cols, skip_last_rows=False, skip_num=0):
    """
    Generalized Excel cleaning function.
    
    Parameters:
    - invoice_excel: Excel file full name.
    - row: The row number in excel where you want the first heading.
    - num_cols: The number of columns there should be.
    - skip_last_rows: True or False -- Are there any rows to skip at the end during import? (optional)
    - skip_num: The row number you want the excel to end on. (optional)
    """
    
    if skip_last_rows == True:

        xls = pd.ExcelFile(invoice_excel)
        sheet_name = xls.sheet_names[0]
        
        wb = openpyxl.load_workbook(invoice_excel, read_only=True)
        ws = wb[sheet_name]
        total_rows = ws.max_row

        last_row_of_data = skip_num + 1

        skipfooter = total_rows - last_row_of_data
        
        clean_excel = pd.read_excel(invoice_excel, skiprows=row-1, skipfooter=skipfooter) 

    else:

        clean_excel = pd.read_excel(invoice_excel, skiprows=row-1) 
    
    
    clean_excel.columns = clean_excel.columns.str.strip()  # Removes any leading/trailing spaces
    clean_excel.columns = clean_excel.columns.str.replace(r'\s+', ' ', regex=True)
    clean_excel.columns = clean_excel.columns.astype(str)
    clean_excel.columns = clean_excel.columns.str.replace(r'\xa0|\t', ' ', regex=True).str.strip()

    column_names = clean_excel.columns.tolist()
    filtered_columns = [col for col in column_names if "Unnamed" not in col]  

    clean_excel = clean_excel[~clean_excel.apply(lambda row: row.astype(str).str.contains('Totals', case=False, na=False)).any(axis=1)]

    clean_excel = clean_excel.dropna(axis = 0, how = 'all')

    clean_excel.loc[:, 'Date:'] = clean_excel['Date:'].ffill()
    
    clean_excel.loc[:, 'Date:'] = pd.to_datetime(clean_excel['Date:'], format='%Y/%m/%d', errors='coerce').dt.date
    clean_excel = clean_excel.dropna(subset=['Date:'])
    clean_excel.reset_index(drop=True, inplace=True)

    def realign_row(a_row):
        non_empty_values = a_row.dropna().values  # Collect non-empty values in the row
        aligned_row = pd.Series([None] * len(clean_excel.columns), index=clean_excel.columns)  # Create an empty aligned row
        aligned_row[:len(non_empty_values)] = non_empty_values  # Place values from start of row
        return aligned_row

    # Apply realignment function to each row
    clean_excel = clean_excel.apply(realign_row, axis=1)

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
    Excel cleaning function for headers.
    
    Parameters:
    - monitor_excel: Excel file full name.
    """

    excel_2 = pd.read_excel(monitor_excel)

    excel_2.columns = excel_2.columns.str.strip()  # Removes any leading/trailing spaces
    excel_2.columns = excel_2.columns.str.replace(r'\s+', ' ', regex=True)
    excel_2.columns = excel_2.columns.astype(str)

    return excel_2

@st.cache_data
def compare_dfs(excel_1, pair1, compare1, excel_2, pair2, compare2):
    """
    Generalized Excel cleaning function.

    Parameters:
    - excel_1, excel_2: DataFrames to compare.
    - pair1, pair2: Columns to match between DataFrames.
    - compare1, compare2: Optional columns to compare values for matched IDs.
    
    Returns:
    - missing_from_excel_1: Rows in excel_1 not found in excel_2 based on pair columns.
    - missing_from_excel_2: Rows in excel_2 not found in excel_1 based on pair columns.
    - diff_qty_df: DataFrame of mismatched values in compare1/compare2 (if provided).
    - combo: DataFrame of unmatched IDs between both excels.
    """

    filter1 = excel_1[pair1].isin(excel_2[pair2])

    missing_from_excel_1 = excel_1[~filter1]

    filter2 = excel_2[pair2].isin(excel_1[pair1])

    missing_from_excel_2 = excel_2[~filter2]

    part1 = excel_2[pair2].unique().tolist()
    part2 = excel_1[pair1].unique().tolist()
    id_list = part1 + part2
    id_list = set(id_list)

    diff_qty = {}
    combined_list = []

    for i in id_list:
        # Check if there's a match in `compare` for `i`
        excel_1_match = excel_1.loc[excel_1[pair1] == i, compare1]
        # Check if there's a match in `aftermath` for `i`
        excel_2_match = excel_2.loc[excel_2[pair2] == i, compare2]
    
        if not excel_1_match.empty and not excel_2_match.empty:
            qty1 = excel_1_match.values[0]
            qty2 = excel_2_match.values[0]
        
            if qty1 != qty2:
                diff_qty[i] = [qty1, qty2]
        else:
            # Optionally, log or handle cases where `i` was not found in one of the DataFrames
            combined_list.append(i)

    diff_qty_df = pd.DataFrame(diff_qty)
    diff_qty_df.index = ['excel_1', 'excel_2']

    combo = pd.DataFrame(combined_list, columns=['Missing list'])
    combo = combo.sort_values(by='Missing list', ascending=False)

    return missing_from_excel_1, missing_from_excel_2, diff_qty_df, combo

# Set up page
st.set_page_config(
    page_title="Aftermath App",
    page_icon="⎘",
)

def wide_space_default():
    st.set_page_config(layout="wide")

wide_space_default()

col1, col2, col3 = st.columns([1, 6, 1])

with col2:
    st.image("logo.png")

st.markdown("<h1 style='text-align: center'>Aftermath Excel Cleaning and Comparison App</h1>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Navigation")
    st.caption("Use the menu above to switch between pages.")
    st.caption("You're on the **Aftermath App** page. Use this for Aftermath invoices and monitor data.")

    with st.expander("ℹ️ How this page works", expanded=False):
        st.markdown(
            """
            1. Upload the **Invoice Excel** (converted from PDF).  
            2. Answer the questions so the headings and rows line up.  
            3. Upload the **Monitor Data Excel**.  
            4. Review the missing tickets and quantity differences.  
            5. Download the combined Excel file at the bottom.
            """
        )


st.markdown(
    """
   This web application is tailored for Aftermath Disaster Recovery to process multiple Excel files, identify missing ticket numbers within each file, and pinpoint any ticket numbers that appear in both files with varying values in specific columns. One Excel file originates from an invoice that was previously converted from a PDF. This Excel file is likely incomplete and requires cleaning to obtain precise data. The other Excel file is the monitor data and is typically clean, except for the standardization of the header. 

   👈👈 Use the sidebar if you would like to see a **:rainbow[tutorial]** or see the app for general use.

"""
)

st.markdown("<br>", unsafe_allow_html=True)

st.markdown("<h2 style='text-align: center'>Excel Uploads and Questions</h2>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

st.markdown(
    """
   In this section, you'll upload the invoice Excel file and answer a few questions to optimize cleaning. You'll have the option to download the cleaned Excel file afterward. Next, you'll upload the already cleaned monitor data and be given the choice to download the dataframe as well. Finally, you'll be able to view the comparison information and download an aggregated Excel file at the end. 

"""
)

col1, col2, col3 = st.columns([1, 6, 1])

with col1:
    st.write("")

with col2:

    st.markdown("#### Answer a few Questions about the Invoice Excel (from PDF):")

    # Ask if the Excel is clean
    isClean = st.radio(
        "Does this excel look clean?",
        [":rainbow[Yes]", "No"],
        index=None,
        captions=[
            "The excel is clean",
            "It needs to be cleaned"
        ],
    )

    # Stop if not answered yet
    if isClean is None:
        st.warning("Please select an option above.")
        st.stop()

    # --- If it's clean ---
    if isClean == ":rainbow[Yes]":
        uploaded_file = st.file_uploader("Upload the file", key="file1")

        if uploaded_file is not None:
            df1 = pd.read_excel(uploaded_file)
            st.session_state.df1 = df1
            st.dataframe(df1)

            st.write("After checking, does it still look clean?")

        else:
            st.warning("Please upload a file.")

    # --- If it needs cleaning ---
    else:
        st.write("Please answer the following questions:")

        row = st.number_input(
            "What row number does the first heading you want start?",
            min_value=0,
            step=1,
            placeholder="Row number..."
        )

        skipEnd = st.radio(
            "Do you want to skip any rows at the end? For example, there is extra unneeded data. \n\n"
            "**Note**--You do not need to remove rows if it is only totals, but it is the same data.",
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
                "What is the final row you want in the excel?",
                min_value=1,
                step=1,
                placeholder="Final row number..."
            )

        # Final checks
        ready = row is not None and row >= 0

        if skipEnd == ":rainbow[Yes]":
            ready = ready and skip_num is not None and skip_num >= 1

        if not ready:
            st.warning("Please answer all questions before uploading the file.")
            st.stop()

        uploaded_file = st.file_uploader("Upload the file", key="file1")

        if uploaded_file is not None:
            
            if skipEnd == ":rainbow[Yes]":

                df1 = cleanInvoice(uploaded_file, row, 8, skip_last_rows=True, skip_num=skip_num)

                st.session_state.df1 = df1
                st.dataframe(df1)

                st.write("After checking, does it still look clean? If not, please redo until it looks clean.")

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

                st.write("After checking, does it still look clean? If not, please redo until it looks clean.")

                excel_bytes = to_excel(df1)

                # Create download button
                st.download_button(
                    label="📥 Download Cleaned Invoice Excel",
                    data=excel_bytes,
                    file_name='cleaned_invoice.xlsx'
                    )

        else:
            st.warning("Please upload a file.")


    st.markdown("#### Download the Monitor Data Excel:")
    st.write("This is assuming the Monitor Data Excel is clean.")

    uploaded_file = st.file_uploader("Upload the file", key="file2")

    if uploaded_file is not None:
        df2 = cleanMonitorData(uploaded_file)
        st.session_state.df2 = df2
        st.dataframe(df2, hide_index=True)

        st.write("After checking, does it still look clean?")

        excel_bytes = to_excel(df2)

        # Create download button
        st.download_button(
            label="📥 Download Cleaned Monitor Excel",
            data=excel_bytes,
            file_name='cleaned_monitor_data.xlsx'
            )

    else:
        st.warning("Please upload a file.")

    st.markdown("#### Comparison Excels:")
    with st.expander("**Click here to expand or contract**", expanded=True):

        if 'df1' in st.session_state and 'df2' in st.session_state:
            df1 = st.session_state.df1
            df2 = st.session_state.df2

            missing_from_excel_1, missing_from_excel_2, diff_qty_df, combo = compare_dfs(df1, 'FEMA Ticket #', 'Calculated Qty', df2, 'Ticket Number', 'Quantity')

            combo_sorted = combo.sort_values(by='Missing list')

            st.write("**Here Are the Entries Missing from Excel 1 (Invoice Excel)**")
            st.session_state.missing_from_excel_1 = missing_from_excel_1
            st.dataframe(missing_from_excel_1, hide_index=True)

            st.write("**Here Are the Entries Missing from Excel 2 (Monitor Excel)**")
            st.session_state.missing_from_excel_2 = missing_from_excel_2
            st.dataframe(missing_from_excel_2, hide_index=True)

            st.write("**Here Are the Entries with Different Quantities Between the Two Excels**")
            st.session_state.diff_qty_df = diff_qty_df
            st.dataframe(diff_qty_df)

            st.write("**Here Are the Entries Missing from Both Excels in One List**")
            st.session_state.combo = combo
            st.session_state.combo_sorted = combo_sorted
            st.data_editor(combo_sorted, hide_index=True)

        else:
            st.info("Please upload both files to compare.")

    st.markdown("#### Download All Dataframes into Excel File:")

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

        # Create an in-memory buffer
        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Always write a dummy sheet immediately
            placeholder_df = pd.DataFrame({'Notice': ['No data available']})
            placeholder_df.to_excel(writer, sheet_name='No_Data', index=False)

            sheets_written = 0

            if not missing_from_excel_1.empty:
                missing_from_excel_1.to_excel(writer, sheet_name='missing_from_excel_1', index=False)
                sheets_written += 1

            if not missing_from_excel_2.empty:
                missing_from_excel_2.to_excel(writer, sheet_name='missing_from_excel_2', index=False)
                sheets_written += 1

            if not diff_qty_df.empty:
                diff_qty_df.to_excel(writer, sheet_name='diff_qty_df', index=True)
                sheets_written += 1

            if not combo.empty:
                combo.to_excel(writer, sheet_name='combo_missing_id', index=False)
                sheets_written += 1

            if sheets_written > 0:
                del writer.book['No_Data']

        # Move to the beginning of the buffer
        output.seek(0)

        # Create download button
        st.download_button(
            label="📥 Download Excel",
            data=output,
            file_name="df_comb_new.xlsx"
            )
        
    else:
        st.info("Run the comparison above to enable the download.")

with col3:
    st.write("")