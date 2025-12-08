# Import Libraries 
import streamlit as st

# Set up page
st.set_page_config(
    page_title="Tutorial",
    page_icon="💡",
    layout="wide",
)

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
"""
)

st.write("")
st.markdown("<h4 style='text-align: center'>Download Samples</h4>", unsafe_allow_html=True)
st.write("")

# =======================
# General App Section
# =======================

st.markdown("<h4 style='text-align: left'>General App</h4>", unsafe_allow_html=True)
st.write("")

with st.expander("**Information**", expanded=True):
    st.markdown(
        """
The General App is designed for flexible use across a variety of spreadsheets. It has two main parts. The first part helps turn messy Excel files into structured data that you can either download or pass directly into the comparison step. This is especially helpful when invoices have been converted from PDF to Excel and contain extra spaces, inconsistent formatting, or other issues that can cause errors when analyzing the spreadsheet.

It is important to note that this app does not attempt to automatically clean every possible messy Excel layout. Instead, it focuses on formats that are similar to invoices or related documents. If you encounter an unexpected issue during the cleaning process, you may need to fix that specific part of the file manually before continuing to the comparison step.

The comparison step requires two key pieces of information from each Excel file. First, you choose the column that contains the shared IDs. This is the column that connects the two spreadsheets, such as a SKU, ticket number, or other unique identifier. Second, you select the column in each file that you want to compare. These comparison columns should represent the same type of information, such as the price of a product or the reported weight, so that the tool can highlight any differences between the two files.
        """
    )

with st.expander("**Step by Step Guide**", expanded=True):
    st.markdown(
        """
This section will provide a step by step guide, including short videos that show how to:

1. Upload and clean a sample Excel file.  
2. Review and download the cleaned data.  
3. Select matching ID and comparison columns.  
4. Run a comparison and interpret the results.
        """
    )

st.write("")
st.markdown("<h4 style='text-align: left'>Aftermath App</h4>", unsafe_allow_html=True)
st.write("")

# =======================
# Aftermath App Section
# =======================

with st.expander("**Information**", expanded=True):
    st.markdown(
        """
The Aftermath App is a specialized version of the tool created for Aftermath Disaster Recovery. It was designed to be more streamlined, based on the fact that the two Excel files they work with follow a consistent structure. The monitor data has a nearly identical layout each time, and the invoices usually follow a similar format, with only the values changing from file to file.

When invoices are converted from PDF to Excel, occasional formatting issues can still appear, and this app includes logic to handle many of those cases. Because the inputs are well understood, the goal was to ask the user as few questions as possible. The app makes more assumptions about the layout and column names, which speeds up the workflow for users who work with the same types of files repeatedly.
        """
    )

with st.expander("**Step by Step Guide**", expanded=True):
    st.markdown(
        """
This section will provide a step by step guide, including short videos that show how to:

1. Upload the monitor and invoice files used by Aftermath.  
2. Confirm or adjust any preselected settings.  
3. Review how the app cleans and aligns the data.  
4. Run the comparison and interpret the output that highlights differences.
        """
    )
