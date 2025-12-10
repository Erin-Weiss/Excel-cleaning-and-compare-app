# Import Libraries 
import streamlit as st
from PIL import Image
import base64
from io import BytesIO
import streamlit.components.v1 as components

# Set up page
st.set_page_config(
    page_title="Excel Help",
    page_icon="⎘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Sidebar ---
with st.sidebar:
    st.markdown("### Navigation")
    st.caption("Use the menu above to switch between pages.")
    st.caption("You're on the **Introduction** page. Start here for an overview.")

    with st.expander("ℹ️ How this page works", expanded=False):
        st.markdown(
            """
            This page gives you a quick overview of the tools available:

            - **Tutorial** – a guided walkthrough for first-time users  
            - **General App** – a flexible Excel cleaning and comparison tool  
            - **Aftermath App** – optimized for Aftermath Disaster Recovery workflows  
            - **Contact** – get in touch if you need help or have questions  
            """
        )

# --- Hero image / title ---
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.image("clean-compare.png")

st.markdown(
    """
    <div style="
        padding: 18px;
        border-radius: 10px;
        background-color: #F1F5FF;
        border-left: 5px solid #4A6CFF;
        margin-bottom: 20px;
        text-align: center;
    ">
        <h1 style="margin-top: 0;">🚀 Getting Started</h1>
        <p style="font-size: 16px;">
            Begin with the <strong>Tutorial</strong> section in the sidebar.  
            It will guide you through uploading spreadsheets, cleaning the data,  
            and performing your first comparison. The tutorial page also has example spreadsheets to use. Or jump right in and try out the <strong>General App</strong> and follow the instructions there.   
        </p>
    </div>
    """,
    unsafe_allow_html=True
)

# --- Introduction ---
st.markdown("<h2 style='text-align: center'>Introduction</h2>", unsafe_allow_html=True)
st.write("")

col1, col2, col3 = st.columns([0.25, 3, 0.25])
with col2:
    st.markdown(
        """
        This website provides two powerful tools for cleaning and comparing Excel files. The Aftermath App was originally built for Aftermath Disaster Recovery and is optimized for their formatting, workflow, and the kinds of spreadsheet issues that come from converting invoice PDFs into Excel. The General App is designed for a wider audience and can handle a broader range of Excel layouts, offering more flexibility for users with varied data sources.

        Both tools follow the same two-step process:

        1. Clean the spreadsheets to remove noise, structure the data, and standardize the layout.  
        2. Compare two cleaned Excel files that share a common identifying column. It matches rows from two separate spreadsheets based on a shared key and highlights any discrepancies or missing records.

        Users can choose to stop after cleaning the files or continue to the comparison step to see how the two datasets align. Alternatively, users can upload two cleaned files and use the comparison feature alone. 

        👈👈 Use the sidebar to explore the **:rainbow[tutorial]** or jump directly into one of the **:rainbow[two apps]** .
        """
    )

    st.markdown(
        """
        #### Why This Matters:

        Messy or inconsistent Excel files can lead to missed discrepancies, reporting errors, and hours of manual cleanup. By standardizing the data first and offering automated comparison tools, this app helps you catch issues quickly, ensure accuracy, and spend more time analyzing your data rather than fixing it.

        #### Who This Tool Is For:

        These tools are built for anyone who needs to repeatedly clean or compare spreadsheets, including project managers, analysts, auditors, and operational teams who rely on consistent and trustworthy data. They were designed with PDF-to-Excel invoice conversions in mind, so they handle many of the common issues found in those files. However, they are not intended to solve every possible messy Excel scenario. The General App provides broader flexibility for users outside Aftermath, while the Aftermath App is optimized for the specific formatting and workflow needs of that organization.

        """
    )

# =======================
# Image Carousel Section
# =======================

st.write("")
st.markdown("<h2 style='text-align: center'>🧼 Cleaning</h2>", unsafe_allow_html=True)
st.write("")

# === Load & Encode Images ===
def get_image_base64_and_style(path, max_width=900):
    img = Image.open(path)

    # --- Resize calculation for display style ---
    width, height = img.size
    scale_w = min(width, max_width)
    scale_ratio = scale_w / width
    scale_h = int(height * scale_ratio)

    # --- Convert PIL image to bytes once and base64-encode ---
    buffered = BytesIO()
    img.save(buffered, format="PNG")  
    base64_str = base64.b64encode(buffered.getvalue()).decode()

    style = f"width:{scale_w}px; height:{scale_h}px; object-fit:contain;"
    return base64_str, style, scale_h

# === Customize images ===
image_data = [
    {
        "path": "dirty-invoice-pic.png",
        "title": "By using this app you can take an excel that looks like this:",
        "caption": "Original Excel of the invoice before preprocessing.",
        "title_margin_top": "10px",
        "image_margin_top": "12px",
        "caption_margin_top": "15px"
    },
    {
        "path": "dirty_dataframe.png",
        "title": "Which the computer reads in like this:",
        "caption": "Extracted table if read into the computer with no modifications.",
        "title_margin_top": "20px",
        "image_margin_top": "6px",
        "caption_margin_top": "10px"
    },
    {
        "path": "clean_dataframe.png",
        "title": "And turns it into this:",
        "caption": "Structured and cleaned data after using app.",
        "title_margin_top": "5px",
        "image_margin_top": "30px",
        "caption_margin_top": "20px"
    }
]

# === Generate slides and indicators ===
carousel_items = ""
carousel_indicators = ""
max_slide_height = 0

for i, item in enumerate(image_data):
    base64_str, img_style, img_height = get_image_base64_and_style(item["path"])
    is_active = "active" if i == 0 else ""
    max_slide_height = max(max_slide_height, img_height + 140)

    carousel_indicators += f"""
    <button type="button" data-bs-target="#carouselExample" data-bs-slide-to="{i}" {'class="active"' if i == 0 else ''} aria-current="{'true' if i == 0 else 'false'}" aria-label="Slide {i+1}"></button>
    """

    carousel_items += f"""
    <div class="carousel-item {is_active}">
        <div class="d-flex flex-column justify-content-center align-items-center slide-content">
            <h5 style="margin-top:{item['title_margin_top']};">{item['title']}</h5>
            <img src="data:image/png;base64,{base64_str}" style="{img_style}; margin-top:{item['image_margin_top']};" class="d-block">
            <p class="text-muted" style="margin-top:{item['caption_margin_top']};">{item['caption']}</p>
        </div>
    </div>
    """

# === Carousel autoplay + pause config ===
autoplay_interval = 5000  # milliseconds
pause_on_hover = "hover"

# === Final Carousel HTML ===
carousel_html = f"""
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet" />
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>

<style>
  .carousel-inner {{
    background-color: white;
    height: {max_slide_height}px;
    overflow: hidden;
  }}

  .carousel-item {{
    position: relative;
    height: 100%;
    display: none;
    text-align: center;
    padding: 20px 0;
  }}

  .carousel-item.active {{
    display: flex;
    justify-content: center;
    align-items: center;
  }}

  .slide-content {{
    max-width: 900px;
    margin: auto;
  }}

  .carousel-indicators [data-bs-target] {{
    background-color: #333;
  }}

  .carousel-control-prev-icon,
  .carousel-control-next-icon {{
    background-color: rgba(0,0,0,0.5);
    border-radius: 50%;
  }}
</style>

<div id="carouselExample" class="carousel slide" data-bs-ride="carousel" data-bs-interval="{autoplay_interval}" data-bs-pause="{pause_on_hover}" data-bs-touch="true">
  <div class="carousel-indicators">
    {carousel_indicators}
  </div>
  <div class="carousel-inner">
    {carousel_items}
  </div>
  <button class="carousel-control-prev" type="button" data-bs-target="#carouselExample" data-bs-slide="prev">
    <span class="carousel-control-prev-icon" aria-hidden="true"></span>
    <span class="visually-hidden">Previous</span>
  </button>
  <button class="carousel-control-next" type="button" data-bs-target="#carouselExample" data-bs-slide="next">
    <span class="carousel-control-next-icon" aria-hidden="true"></span>
    <span class="visually-hidden">Next</span>
  </button>
</div>
"""

# === Show Carousel in Streamlit ===
components.html(carousel_html, height=max_slide_height + 15)

# === Carousel Explanation ===
col1, col2, col3 = st.columns([0.25, 3, 0.25])
with col2:
    st.markdown(
        """
        Cleaning the data first is essential as it ensures the computer can correctly interpret the rows, columns, and values before running any comparisons. It also makes the cleaned Excel file easy to download and reuse for your own analysis.
        """
    )

# =======================
# Table Section
# =======================

st.markdown("<h2 style='text-align: center'>🧾 Excel Comparison Features</h2>", unsafe_allow_html=True)
st.write("")

col1, col2, col3 = st.columns([0.25, 3, 0.25])
with col2:
    st.markdown("""
    This step in each of the apps compares two Excel files that share a common **ID** column, and it performs the following checks:
    """)    

    st.write("")

    table_html = """
    <table style="width:100%; border-collapse: collapse; text-align: center;">
    <thead>
        <tr>
        <th style="border: 1px solid #ddd; padding: 8px;">Feature</th>
        <th style="border: 1px solid #ddd; padding: 8px;">What it Does</th>
        </tr>
    </thead>
    <tbody>
        <tr>
        <td style="border: 1px solid #ddd; padding: 8px;">Missing ID Detection</td>
        <td style="border: 1px solid #ddd; padding: 8px;">Identifies IDs that appear in one file but not the other.</td>
        </tr>
        <tr>
        <td style="border: 1px solid #ddd; padding: 8px;">Field Value Comparison</td>
        <td style="border: 1px solid #ddd; padding: 8px;">Compares values (such as price or quantity) for matching IDs across both files and flags anything that doesn’t match.</td>
        </tr>
    </tbody>
    </table>
    """

    st.markdown(table_html, unsafe_allow_html=True)
