# ⎘ Excel Cleaning & Comparison App

**Author:** Erin Weiss  
👉 [**Interactive Web App (Streamlit)**](https://excel-cleaning-and-compare-app.streamlit.app) 

---

## 📘 Overview

This project provides an interactive **Streamlit web application** for **cleaning**, **standardizing**, and **comparing Excel files**—especially spreadsheets created from **PDF-to-Excel invoice conversions**.  

The app includes two specialized tools:

- **General App** – a flexible, configurable cleaning + comparison tool for many spreadsheet layouts.  
- **Aftermath App** – a streamlined workflow designed for **Aftermath Disaster Recovery**, where the invoice and monitor data follow consistent structures.  
- A full **Tutorial** with sample files is included to guide new users.  
- A built-in **Contact page** allows users to send questions, feature requests, or bug reports.

Whether your spreadsheets come from messy invoices, monitoring systems, or manually edited Excel files, this app helps you **clean them reliably** and **compare values across files** with minimal manual effort.

---

## ⚙️ Approach

This project was built around **real-world messy spreadsheets first**, and polished features second. The design focuses on three ideas:

1. **Start from the workflow, not the code.**  
   The app was shaped around a concrete use case: comparing **invoice Excel files converted from PDF** against **monitor data**. Instead of assuming clean, rectangular tables, the logic was designed to tolerate header noise, extra rows, shifted columns, and inconsistent date formats. All cleaning options in the General App grew out of specific failure modes seen in these files.

2. **Separate “guided” and “flexible” tools.**  
   To balance convenience and generality, the app is split into two layers:  
   - The **Aftermath App** encodes a known workflow with fixed column names and assumptions about layout. It asks just a few questions (header row, last useful row) and then applies a consistent cleaning and comparison pipeline, making it fast and repeatable for Aftermath Disaster Recovery.  
   - The **General App** reuses the same underlying ideas but exposes more configuration. It turns cleaning into a guided form: users describe what’s wrong with each sheet (header row, extra rows, date issues, missing values), and only then are the corresponding options shown. This keeps the interface approachable while still covering a wide range of layouts.

3. **Make the logic transparent and reproducible.**  
   At each stage, the app aims to show exactly **what it’s doing** and to leave users with reusable artifacts:  
   - Intermediate outputs (cleaned invoices, monitor data, and general cleaned files) are displayed in Streamlit and can be **downloaded as Excel**.  
   - Comparison results are broken into separate tables (missing from A, missing from B, mismatched values, combined missing IDs) and bundled into a **multi-sheet Excel file** for auditing and documentation.  
   - A dedicated **Tutorial** page, with sample files and screenshots, mirrors the internal logic step-by-step so users can see how parameter choices (like header row, date format, or keywords to drop) affect the final results.

Together, these choices create a workflow where users can **experiment safely**, understand how their data is being transformed, and then apply the same pattern to their own Excel files—either via the streamlined Aftermath App or the more general, configurable tool.

---

## 📊 Key Features

### 🧼 Excel Cleaning
- Handles extremely messy Excel files, especially those created from PDFs.  
- Cleans headers, drops junk rows, realigns misaligned cells.  
- Standardizes date columns and prevents “fake dates” from slipping through.  
- Supports filling missing values with constants or forward/backward fill.  
- Lets users track which values or dates were filled.  
- Supports auto-detection or manual selection of date columns.  
- Allows cropping to desired rows or columns.  
- Outputs a clean, consistent DataFrame ready for comparison or reuse.

### 🔍 Excel Comparison
- Case-insensitive matching of key/ID values.  
- Flexible cross-file column comparison (e.g., *QTY vs Quantity*, *Rate vs UNIT PRICE*).  
- Highlights rows missing from either source.  
- Produces clean mismatch tables with ID, columns compared, and both values.  
- Supports exporting all comparison results into a multi-sheet Excel file.

### 🧭 Tutorial with sample data
- A full guided walkthrough with screenshots, GIFs, and explanations.  
- Downloadable ZIP file of sample invoices and monitor data.  
- Step-by-step instructions for:
  - **Aftermath workflows**  
  - **General comparison workflows**  
  - **Cleaning-only exercises**

### ⚡ Aftermath-specific tool
- Designed for Aftermath Disaster Recovery’s invoice + monitor format.  
- Minimizes questions—uses fixed assumptions for speed.  
- Automatically selects:
  - ID columns  
  - Quantity columns  
  - Date column logic  
  - Expected column count  
- Ideal for repeated workflows using the same file structures.

### ✉️ Contact Page
- Built-in support form for questions, bug reports, or feedback.  
- Email validation, CAPTCHA, and Formspree integration.  
- Notes that replies may not be immediate, but inquiries are welcome.

---

## 🧰 Tools Used

- **Language:** Python  
- **Framework:** Streamlit  
- **Libraries:**  
  - `pandas` – data cleaning and comparison  
  - `openpyxl` – Excel reading and writing  
  - `email_validator` – email validation  
  - `captcha` – CAPTCHA generation  
  - `streamlit_js_eval` – client-side JS for form reset  
  - `Pillow (PIL)` – image handling  
  - `requests` – Formspree submissions  
- **Deployment:** Streamlit Cloud  

---

## 🔗 Links

- **Live Streamlit App:** [Excel Cleaning & Comparison App](https://excel-cleaning-and-compare-app.streamlit.app)
- **GitHub Repository:** [View Source Code](https://github.com/Erin-Weiss/Excel-cleaning-and-compare-app)
- **Portfolio Page:** Will insert when available 
- **Sample Files (Tutorial ZIP):** Included directly in the app’s [Tutorial page](https://excel-cleaning-and-compare-app.streamlit.app/Tutorial)

