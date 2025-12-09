# Import Libraries 
import streamlit as st
import base64
from streamlit_pdf_viewer import pdf_viewer

# Set up page
st.set_page_config(
    page_title="Tutorial",
    page_icon="💡",
    layout="wide",
)

# Set up inline images
def img_to_base64(path):
    with open(path, "rb") as f:
        data = f.read()
        return base64.b64encode(data).decode("utf-8")

img64_1 = img_to_base64("pages/invoice-headings.jpg")

img64_2 = img_to_base64("pages/last-line.jpg")

img64_3 = img_to_base64("pages/invoice-1-screenshot.jpg")

img64_4 = img_to_base64("pages/invoice-empty-row.jpg")

img64_5 = img_to_base64("pages/row-fixed.jpg")

img64_6 = img_to_base64("pages/num-cols.jpg")

# Set up collapsible PDF
def display_pdf_collapsible(pdf_file, label="**📄 View Sample File Table of Contents**", zoom_level: float = 1.25):
    with st.expander(label, expanded=False):
        with open(pdf_file, "rb") as f:
            pdf_bytes = f.read()
        pdf_viewer(pdf_bytes, zoom_level=zoom_level)

# Set up downloadable zip file and button
zip_path = "pages/Sample_Excels.zip"

with open(zip_path, "rb") as f:
    zip_bytes = f.read()

st.markdown("""
    <style>
    /* Target Streamlit's download button wrapper */
    div.stDownloadButton > button {
        background-color: #217346 !important;   /* Excel green */
        color: white !important;
        border-radius: 4px !important;
        border: 1px solid #185C37 !important;
        padding: 0.6em 1.6em !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
    }

    div.stDownloadButton > button:hover {
        background-color: #185C37 !important;   /* darker on hover */
        border-color: #0F3B23 !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- Sidebar header ---
st.sidebar.header("Tutorial")

with st.sidebar:
    st.markdown("### Navigation")
    st.caption("Use the menu above to switch between pages.")
    st.caption("You are on the **Tutorial** page. Follow the steps for a guided walkthrough.")

    with st.expander("ℹ️ How this page works", expanded=False):
        st.markdown(
            """
            1. Read each section in order.  
            2. Download sample files to practice.  
            3. When you are comfortable, move to the **General App** or **Aftermath App**.
            """
        )

# --- Hero section ---
col1, col2, col3 = st.columns([1, 1, 1])
with col2:
    st.image("pages/excel-gif.gif")

st.markdown("<h1 style='text-align: center'>Tutorial</h1>", unsafe_allow_html=True)
st.write("")

st.markdown(
    """
This tutorial walks you through how to use both apps and provides downloadable sample data and invoices you can experiment with. Aftermath has granted permission to use any data displayed on this website, and some details in the sample files have been anonymized.

These tutorials, together with the sample files in the ZIP, give users a clear path to:

- See how the cleaning logic behaves in real scenarios.  
- Understand how keys and comparison columns work.  
- Compare the General App’s flexibility to the more streamlined Aftermath App.
"""
)

# --- Sample Excel section ---

st.write("")
st.markdown("<h4 style='text-align: center'>Download Samples</h4>", unsafe_allow_html=True)
st.write("")

display_pdf_collapsible("pages/Tutorial_Table_of_Contents.pdf", zoom_level=1.25)

col_left, col_center, col_right = st.columns([1.5, 1, 1.5])

with col_center:
    st.download_button(
        label="📦 Download Tutorial Files ZIP",
        data=zip_bytes,
        file_name="Sample_Excels.zip",
        mime="application/zip",
        use_container_width=True
    )

# =======================
# General App Section
# =======================

st.markdown("<h2 style='text-align: left'>General App</h2>", unsafe_allow_html=True)
st.write("")

st.markdown("<h3 style='text-align: left'>Information</h3>", unsafe_allow_html=True)

with st.expander("**Click to expand or collapse**", expanded=True):
    st.markdown(
        """
The General App is designed for flexible use across a variety of spreadsheets. It has two main parts. The first part helps turn messy Excel files into structured data that you can either download or pass directly into the comparison step. This is especially helpful when invoices have been converted from PDF to Excel and contain extra spaces, inconsistent formatting, or other issues that can cause errors when analyzing the spreadsheet.

It is important to note that this app does not attempt to automatically clean every possible messy Excel layout. Instead, it focuses on formats that are similar to invoices or related documents. If you encounter an unexpected issue during the cleaning process, you may need to fix that specific part of the file manually before continuing to the comparison step.

The comparison step requires two key pieces of information from each Excel file. First, you choose the column that contains the shared IDs. This is the column that connects the two spreadsheets, such as a SKU, ticket number, or other unique identifier. Second, you select the column in each file that you want to compare. These comparison columns should represent the same type of information, such as the price of a product or the reported weight, so that the tool can highlight any differences between the two files.
        """
    )

st.markdown("<h3 style='text-align: left'>General App tutorial</h3>", unsafe_allow_html=True)

with st.expander("**Click to expand or collapse**", expanded=False):
    st.markdown(
        """
The General App is designed to handle a wide range of Excel layouts, especially invoice-style tables that may have come from PDF conversions. You can use the sample files in the tutorial ZIP to practice three types of tasks:

1. Cleaning an invoice and comparing it with monitor data (Aftermath examples).  
2. Comparing two general invoice files that use different but related columns.  
3. Cleaning invoices without doing any comparison.

Below are suggested walk-throughs for each type.

---

#### A. Aftermath examples with the General App (Pairs 1 and 2)

You can use **Pair 1** and **Pair 2** to run the Aftermath workflow through the **General App** instead of the Aftermath App.

##### Pair 1 – `Dirty-Invoice-Aftermath-1.xlsx` and `Aftermath-Monitor-Data-1.xlsx`

**Step 1: Clean the invoice in the General App**

1. Upload `Dirty-Invoice-Aftermath-1.xlsx` into the **First Excel (Source A)** cleaning section.  
2. In **Step 1 – What is wrong with this sheet?**, select:  
   - “Header row is not on row 1”  
   - “Skip extra rows at the bottom”  
   - “There is a date column to clean/standardize”  
   - “Drop rows containing certain keywords”  
   - “Limit to first N columns”  
3. Set the following options:

   - **Header row is not on row 1**  
     - Header row: `25`  

    *Note*: You can see that the correct header values are in the header row in the picture below. They are between the `unnamed` headings that will be removed later and can be ignored in the preview.
"""
    )

    st.markdown(
        f"""
     <p align="center">
    <img src="data:image/jpeg;base64,{img64_1}" width="700">
    </p>

    <p style="text-align: center; color: gray; font-size: 0.9em;">
        Correct selection of header row
    </p>
    """,
    unsafe_allow_html=True
    )

    st.markdown(
        """
     
        - **Skip extra rows at the bottom**  
            - Final row to keep: `304`  

     *Note*: You can see that the section within the red box is unrelated data to the data above it. This can be removed by selecting the last useful line of data, which in this case is row `304` (underlined in green). Theoretically, you could also select row `309` since the “Drop rows containing certain keywords” will get rid of the rows with those values, but it is easier and more accurate to select `304`.
     """
    )

    st.markdown(
        f"""
     <p align="center">
    <img src="data:image/jpeg;base64,{img64_2}" width="700">
    </p>

    <p style="text-align: center; color: gray; font-size: 0.9em;">
        Last useful line of data (underlined in green) and extraneous data (in red box)
    </p>
    """,
    unsafe_allow_html=True
    )

    st.markdown(
        """
   - **There is a date column to clean/standardize**  
     - Date column: `Auto-detect from columns` or explicitly `Date:`  
     - Date format: `01/31/2024 (MM/DD/YYYY)` or any option that is **not** `Let pandas infer from the data` in this case. 
     - Fill method: `ffill` (forward fill). In this sample there are no missing dates, so either choice works.  

     *Note*: In this file, certain numeric values slide into the date column (see first picture below). By forcing a specific date format, those irregular rows become completely empty in the date column and can be dropped during cleaning. Otherwise, if nothing is done with those numbers, there will be a fake date added and a mostly empty row (see second picture below). Keep situations like this in mind for your own Excel cleaning. Corrected version in final picture below.
     """
    )

    st.write("")

    st.markdown(
        f"""
     <p align="center">
    <img src="data:image/jpeg;base64,{img64_3}" width="700">
    </p>

    <p style="text-align: center; color: gray; font-size: 0.9em;">
        Number spanning all columns that needs to be deleted
    </p>
    """,
    unsafe_allow_html=True
    )

    st.write("")
    st.write("")

    st.markdown(
        f"""
     <p align="center">
    <img src="data:image/jpeg;base64,{img64_4}" width="700">
    </p>

    <p style="text-align: center; color: gray; font-size: 0.9em;">
        Number turned into date if not removed properly
    </p>
    """,
    unsafe_allow_html=True
    )

    st.write("")
    st.write("")

    st.markdown(
        f"""
     <p align="center">
    <img src="data:image/jpeg;base64,{img64_5}" width="700">
    </p>

    <p style="text-align: center; color: gray; font-size: 0.9em;">
        Row removed after number correctly removed using date option
    </p>
    """,
    unsafe_allow_html=True
    )

    st.markdown(
        """
   - **Drop rows containing certain keywords**  
     - Keywords: `totals`  
       This will also catch strings such as “Grand Totals”.

   - **Limit to first N columns**  
     - Number of columns: `8`

4. Run the cleaning step.  
5. The cleaned invoice should look the same as `invoice1-cleaned_excel1.xlsx`.

**Step 2: Clean the monitor data in the General App**

1. Upload `Aftermath-Monitor-Data-1.xlsx` into the **Second Excel (Source B)** cleaning section.  
2. In **Step 1 – What is wrong with this sheet?**, select:  
   - “Do NOT realign rows (keep original alignment)”  
     (This keeps the monitor data in its original column layout.)  
3. Leave other cleaning options off unless you see a specific problem.  
4. Run the cleaning step.  
5. The cleaned file should look the same as `monitor1-cleaned_excel2.xlsx`.

**Step 3: Compare the two files**

1. Scroll to **Step 3: Compare the two Excels**.  
2. Choose the key columns:

   - Key or ID column in First Excel: `FEMA Ticket #`  
   - Key or ID column in Second Excel: `Ticket Number`

3. Choose the columns to compare:

   - From First Excel: `Calculated Qty`  
   - From Second Excel: `Quantity`

4. Leave the “Ignore differences in letter case” option checked.  
5. Run the comparison and review:

   - Entries missing from Excel 1 (invoice).  
   - Entries missing from Excel 2 (monitor).  
   - Entries with different quantities between the two files.  
   - The combined list of ticket IDs missing from one or both files.

6. If desired, download `comparison_results1.xlsx` using the download button.

---

##### Pair 2 – `Dirty-Invoice-Aftermath-2.xlsx` and `Aftermath-Monitor-Data-2.xlsx`

Pair 2 follows the same pattern with slightly different header row and skipping options.

**Step 1: Clean the invoice**

1. Upload `Dirty-Invoice-Aftermath-2.xlsx` into **First Excel (Source A)**.  
2. In **Step 1 – What is wrong with this sheet?**, select:  
   - “Header row is not on row 1”  
   - “There is a date column to clean/standardize”  
   - “Drop rows containing certain keywords”  
   - “Limit to first N columns”  
3. Set the following options:

   - Header row: `21`  
   - Date column: `Auto-detect from columns` or `Date:`  
   - Date format: `01/31/2024 (MM/DD/YYYY)` or any non-inferred option  
   - Date fill: `ffill`  
   - Drop keywords: `totals`  
   - Limit to first N columns: `8`

4. Run the cleaning step and confirm the result looks the same as `invoice2-cleaned_excel1.xlsx`.

**Step 2: Clean the monitor data**

1. Upload `Aftermath-Monitor-Data-2.xlsx` into **Second Excel (Source B)**.  
2. In **Step 1 – What is wrong with this sheet?**, select:  
   - “Do NOT realign rows (keep original alignment)”  
3. Run the cleaning step and confirm it matches `monitor2-cleaned_excel2.xlsx`.

**Step 3: Compare**

Use the same comparison setup as Pair 1:

- Keys  
  - First Excel: `FEMA Ticket #`  
  - Second Excel: `Ticket Number`  
- Compare columns  
  - First Excel: `Calculated Qty`  
  - Second Excel: `Quantity`  

Run the comparison and review the outputs, which should be the same as `comparison_results2.xlsx`.

---

#### B. General comparison example (Pair 3)

This example shows how to compare two invoice files that use different but related column names.

**Files**

- `Comparison-invoice-part1.xlsx`  
- `Comparison-invoice-part2.xlsx`

**Step 1: Clean `Comparison-invoice-part1.xlsx`**

1. Upload `Comparison-invoice-part1.xlsx` into **First Excel (Source A)**.  
2. In **Step 1 – What is wrong with this sheet?**, select:  
   - “Header row is not on row 1”  
   - “Skip extra rows at the bottom”  
   - “Limit to first N columns”  
   - “Fill missing values in non-date columns”  
   - (Optional) “Track which values were filled”  
3. Set the following options:

   - Header row: `16`  
   - Final row to keep: `23`  
   - Limit to first N columns: `8`  
   - Fill missing values in non-date columns:  
     - Column `UNIT OF MEASURE`  
       - Method: “Use a constant value” set to `Unknown` if you want to match example. 
     - Column `TAX`  
       - Method: “Use a constant value” set to `No` if you want to match example. 

4. Run the cleaning step. The result should match `comparison1-cleaned_excel1.xlsx`.

**Step 2: Clean `Comparison-invoice-part2.xlsx`**

1. Upload `Comparison-invoice-part2.xlsx` into **Second Excel (Source B)**.  
2. In **Step 1 – What is wrong with this sheet?**, select:  
   - “Header row is not on row 1”  
   - “Skip extra rows at the bottom”  
   - “Start from a specific column letter (drop left-side junk)”  
   - “Limit to first N columns”  
3. Set the following options:

   - Header row: `18`  
   - Final row to keep: `23`  
   - Start from column letter: `B` (uppercase or lowercase is fine)  
   - Limit to first N columns: `8` 
   
   *Note*: The headers we are trying to isolate are the ones in the red box. Below, you can see the letters that correspond to the columns in the Excel sheet. You select `B` as the column that you want to start, and from that starting point, you count how many total columns you want in your dataframe. In this case that number is `8`.
   """
    )

    st.write("")

    st.markdown(
        f"""
     <p align="center">
    <img src="data:image/jpeg;base64,{img64_6}" width="700">
    </p>

    <p style="text-align: center; color: gray; font-size: 0.9em;">
        Dirty invoice with column labels and desired columns numbered
    </p>
    """,
    unsafe_allow_html=True
    )

    st.markdown(
        """
   
4. Run the cleaning step and confirm it matches `comparison2-cleaned_excel2.xlsx`.

**Step 3: Compare the two cleaned files**

1. In **Step 3: Compare the two Excels**, set:

   - Key or ID column in First Excel: `PART NUMBER`  
   - Key or ID column in Second Excel: `Description`  

2. Choose comparison columns:

   - From First Excel: `Title`, `QTY`, `UNIT PRICE`  
   - From Second Excel: `Title`, `Quantity`, `Rate`  

3. Make sure **“Ignore differences in letter case”** is checked to match example.  
4. Run the comparison and review the results. They should match the structure of `comparison_results3.xlsx`.

---

#### C. Cleaning only examples (Tests 1 and 2)

These examples are intended to practice the cleaning step without comparison.

##### Test 1 – `Dirty-invoice-test-1.xlsx`

1. Upload `Dirty-invoice-test-1.xlsx` into **First Excel (Source A)**.  
2. In **Step 1 – What is wrong with this sheet?**, select:  
   - “Header row is not on row 1”  
   - “Skip extra rows at the bottom”  
   - “There is a date column to clean/standardize”  
   - “Limit to first N columns”  
   - “Fill missing values in non-date columns”  
   - (Optional) “Track which values were filled”  
   - (Optional) “Track which dates were filled”  

3. Set the options:

   - Header row: `16`  
   - Final row to keep: `23`  
   - Date column: `Auto-detect from columns` or `DATE ORDERED`  
   - Date format:  
     - `Let pandas infer from the data` **or**  
     - `01/31/2024 (MM/DD/YYYY)`  
   - Date fill: `ffill`  
   - Limit to first N columns: `8`  
   - Fill missing values in non-date columns:  
     - `UNIT OF MEASURE` → constant value `Unknown`  
     - `TAX` → constant value `No`

4. Run the cleaning step. The result should match `General1-cleaned_excel1.xlsx`.

---

##### Test 2 – `Dirty-invoice-test-2.xlsx`

1. Upload `Dirty-invoice-test-2.xlsx` into **First Excel (Source A)**.  
2. In **Step 1 – What is wrong with this sheet?**, select:  
   - “Header row is not on row 1”  
   - “Skip extra rows at the bottom”  
   - “Start from a specific column letter (drop left-side junk)”  
   - “There is a date column to clean/standardize”  
   - “Limit to first N columns”  
   - (Optional) “Track which dates were filled”  

3. Set the options:

   - Header row: `18`  
   - Final row to keep: `21`  
   - Start from column letter: `B`  
   - Date column: `Auto-detect from columns` or `Date`  
   - Date format:  
     - `Let pandas infer from the data` **or**  
     - `01/31/2024 (MM/DD/YYYY)`  
   - Date fill: `ffill`  
   - Limit to first N columns: `8`

4. Run the cleaning step. The result should match `General2-cleaned_excel1.xlsx`.

These cleaning-only scenarios are useful for experimenting with different combinations of options and seeing how the cleaned output changes.

        """
    )

st.write("")
st.markdown("<h2 style='text-align: left'>Aftermath App</h2>", unsafe_allow_html=True)
st.write("")

# =======================
# Aftermath App Section
# =======================

st.markdown("<h3 style='text-align: left'>Information</h3>", unsafe_allow_html=True)

with st.expander("**Click to expand or collapse**", expanded=True):
    st.markdown(
        """
The Aftermath App is a specialized version of the tool created for Aftermath Disaster Recovery. It was designed to be more streamlined, based on the fact that the two Excel files they work with follow a consistent structure. The monitor data has a nearly identical layout each time, and the invoices usually follow a similar format, with only the values changing from file to file.

When invoices are converted from PDF to Excel, occasional formatting issues can still appear, and this app includes logic to handle many of those cases. Because the inputs are well understood, the goal was to ask the user as few questions as possible. The app makes more assumptions about the layout and column names, which speeds up the workflow for users who work with the same types of files repeatedly.
        """
    )

st.markdown("<h3 style='text-align: left'>Aftermath App tutorial</h3>", unsafe_allow_html=True)

with st.expander("**Click to expand or collapse**", expanded=False):
    st.markdown(
        """
The Aftermath App is a streamlined version of the tool that assumes a consistent layout for Aftermath invoices and monitor data. It asks fewer questions and uses fixed column names under the hood.

You can use **Pair 1** and **Pair 2** from the tutorial ZIP to walk through the full workflow.

---

#### A. Pair 1 – Aftermath 1

**Files**

- Invoice (from PDF): `Dirty-Invoice-Aftermath-1.xlsx`  
- Monitor data: `Aftermath-Monitor-Data-1.xlsx`

##### Step 1: Clean the invoice in the Aftermath App

1. In the **invoice section**, when asked whether the Excel looks clean, choose **“No”**.  
2. For the question:  
   - “What row number does the first heading you want start?”  
     - Enter: `25`  
3. For the question about skipping rows at the end, choose **“Yes”**.  
4. When asked:  
   - “What is the final row you want in the Excel?”  
     - Enter: `304`  
5. Upload `Dirty-Invoice-Aftermath-1.xlsx` when prompted.  
6. The app will:

   - Remove unwanted rows above row 25.  
   - Drop extra rows after row 304.  
   - Standardize the `Date:` column using the built-in logic.  
   - Realign values to the left and keep the first 8 important columns.

7. Review the preview. It should look match `invoice1-cleaned_excel1.xlsx`.  
8. Download the cleaned invoice if you want a copy.

##### Step 2: Load the monitor data

1. In the **monitor data section**, upload `Aftermath-Monitor-Data-1.xlsx`.  
2. The app will clean only the headers and keep the rest of the data as is.  
3. Review the preview and optionally download it. It should match `monitor1-cleaned_excel2.xlsx`.

##### Step 3: Run the comparison

1. Open the **Comparison Excels** section.  
2. Once both `df1` (invoice) and `df2` (monitor) are present, the app will automatically:

   - Use `FEMA Ticket #` from the invoice and `Ticket Number` from the monitor as the match keys.  
   - Compare `Calculated Qty` (invoice) to `Quantity` (monitor).

3. Review the results:

   - **Entries missing from Excel 1 (Invoice Excel)**  
     Rows present in monitor data but not found in the invoice.  

   - **Entries missing from Excel 2 (Monitor Excel)**  
     Rows present in the invoice but not found in the monitor file.  

   - **Entries with different quantities between the two Excels**  
     Ticket numbers that appear in both files but have different quantities.  

   - **Entries missing from both Excels in one list**  
     Combined list of ticket IDs that are missing in one or the other.

4. Use the **Download all DataFrames into Excel file** section to create a combined Excel file with all the outputs. This should look like `comparison_results1.xlsx`.

---

#### B. Pair 2 – Aftermath 2

**Files**

- Invoice (from PDF): `Dirty-Invoice-Aftermath-2.xlsx`  
- Monitor data: `Aftermath-Monitor-Data-2.xlsx`

##### Step 1: Clean the invoice

1. In the invoice section, select **“No”** when asked if the Excel looks clean.  
2. For:

   - “What row number does the first heading you want start?”  
     - Enter: `21`  

3. For the question about skipping rows at the end, choose **“No”**.  
4. Upload `Dirty-Invoice-Aftermath-2.xlsx`.  
5. The app will:

   - Start the data at row 21.  
   - Keep all remaining rows.  
   - Clean and standardize the `Date:` column.
   - Realign values to the left and keep the first 8 important columns.

6. Review the cleaned result. It should resemble `invoice2-cleaned_excel1.xlsx`. Download it if you need a copy.

##### Step 2: Load the monitor data

1. Upload `Aftermath-Monitor-Data-2.xlsx` in the monitor section.  
2. The app will adjust only the headers.  
3. The output should match `monitor2-cleaned_excel2.xlsx`.

##### Step 3: Run the comparison

The comparison setup is the same as for Pair 1 and is handled for you automatically inside the Aftermath App:

- Key columns  
  - Invoice: `FEMA Ticket #`  
  - Monitor: `Ticket Number`  

- Comparison columns  
  - Invoice: `Calculated Qty`  
  - Monitor: `Quantity`  

Review the same four result tables and download the combined Excel file at the bottom if you want to save all outputs together. This should look like `comparison_results2.xlsx`.
        """
    )
