# Import Libraries 
import streamlit as st
import pandas as pd
import openpyxl
import string
import numpy as np
import re
from io import BytesIO

# -----------------------------
# Core functions
# -----------------------------
@st.cache_data
def clean_excel(
    file,
    header_row_guess=1,
    sheet_name=0,
    date_col=None,
    date_format=None,
    date_fill_method='ffill',
    track_date_fill=False,
    drop_keywords=None,
    realign=True,
    n_cols=None,
    skip_last_rows=False, 
    skip_num=0,
    start_col_letter=None,
    fill_cols=None,
    fill_methods='ffill',
    track_fill=False
):
    """
    Clean and preprocess an Excel sheet with flexible options.
    """

    # Change letter input to index
    def col_letter_to_index(letter):
        letter = letter.upper()
        num = 0
        for c in letter:
            if c in string.ascii_uppercase:
                num = num * 26 + (ord(c) - ord('A')) + 1
            else:
                raise ValueError(f"Invalid column letter: {letter}")
        return num - 1

    # Crop rows at bottom at the row listed in skip_num
    if skip_last_rows:
        xls = pd.ExcelFile(file)
        if isinstance(sheet_name, int):
            sheet_name = xls.sheet_names[sheet_name]
        wb = openpyxl.load_workbook(file, read_only=True)
        ws = wb[sheet_name]
        total_rows = ws.max_row

        skipfooter = total_rows - skip_num
        df = pd.read_excel(file, sheet_name=sheet_name, skiprows=header_row_guess - 1, skipfooter=skipfooter)
    else:
        df = pd.read_excel(file, sheet_name=sheet_name, skiprows=header_row_guess - 1)

    # Drop columns before specified letter
    if start_col_letter is not None:
        start_idx = col_letter_to_index(start_col_letter)
        df = df.iloc[:, start_idx:]

    # Clean up column headers
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace(r'\s+', ' ', regex=True)
        .str.replace(r'\xa0|\t', ' ', regex=True)
    )

    # Remove unnecessary additional columns "Unnamed"
    column_names = df.columns.tolist()
    filtered_columns = [col for col in column_names if "Unnamed" not in col]

    # Map column names to lower for safe matching
    col_map = {col.lower(): col for col in df.columns}

    # Drop rows with keywords
    if drop_keywords:
        safe_keywords = [re.escape(str(k)) for k in drop_keywords]
        pattern = '|'.join(safe_keywords)
        df = df[~df.apply(lambda row: row.astype(str).str.contains(pattern, case=False, na=False)).any(axis=1)]

    df = df.dropna(how='all')

    # Handle fill_cols
    if fill_cols:
        # If user gives a single string, make it a list
        if isinstance(fill_cols, str):
            fill_cols = [fill_cols]
    
        # Normalize fill_methods: ensure it's a dict with lowercase keys
        if not isinstance(fill_methods, dict):
            fill_methods = {col.lower(): fill_methods for col in fill_cols}
        else:
            # Also normalize keys if user passed dict directly
            fill_methods = {col.lower(): method for col, method in fill_methods.items()}

        for col in fill_cols:
            col_lower = col.lower()
            if col_lower not in col_map:
                raise ValueError(f"fill_col '{col}' not found in columns: {df.columns.tolist()}")
            
            real_col = col_map[col_lower]
            method = fill_methods.get(col_lower, 'ffill')

            # Save missing mask before fill
            if track_fill:
                was_na = df[real_col].isna()

            # bfill, ffill, or str based on user input
            if method == 'ffill':
                df[real_col] = df[real_col].ffill()
            elif method == 'bfill':
                df[real_col] = df[real_col].bfill()
            elif isinstance(method, (int, float, str)):
                df[real_col] = df[real_col].fillna(method)
            else:
                raise ValueError(
                    f"Invalid fill_method '{method}' for column '{col}'. "
                    "Use 'ffill', 'bfill', or a literal value."
                )

            # Add indicator column if requested
            if track_fill:
                filled_col_name = f"{real_col}_filled"
                insert_pos = df.columns.get_loc(filtered_columns[-1]) + 1
                fill_values = was_na & df[real_col].notna()
                df.insert(insert_pos, filled_col_name, fill_values)
                filtered_columns.append(filled_col_name)

    # Handle date_col or auto-detect
    real_date_col = None
    cleaned_cols_lower = [col.lower().strip() for col in df.columns]

    if date_col:
        if date_col.lower().strip() in cleaned_cols_lower:
            idx = cleaned_cols_lower.index(date_col.lower().strip())
            real_date_col = df.columns[idx]
        else:
            raise ValueError(f"date_col '{date_col}' not found.")
    else:
        # Auto detect
        possible_date_cols = [col for col in df.columns if 'date' in col.lower()]
        scores = {}
        for col in possible_date_cols:
            try:
                parsed = pd.to_datetime(df[col], format=date_format, errors='coerce').dt.date
                scores[col] = parsed.notna().sum()
            except Exception:
                continue
        if scores:
            best_col = max(scores, key=scores.get)
            if scores[best_col] > 0:
                real_date_col = best_col

    # Then process date column if found
    if real_date_col:

        # Make a boolean marker before filling
        if track_date_fill:
            was_na = df[real_date_col].isna()

        # Fill dates with chosen method
        if date_fill_method == 'ffill':
            df[real_date_col] = df[real_date_col].ffill()
        elif date_fill_method == 'bfill':
            df[real_date_col] = df[real_date_col].bfill()
        else:
            raise ValueError(f"Invalid date_fill_method '{date_fill_method}'")

        # Add marker column if requested
        if track_date_fill:
            marker_col_name = f"{real_date_col}_was_{date_fill_method}ed"
            insert_pos = df.columns.get_loc(filtered_columns[-1]) + 1
            fill_values = was_na & df[real_date_col].notna()
            df.insert(insert_pos, marker_col_name, fill_values)
            filtered_columns.append(marker_col_name)

        # Parse and clean
        df[real_date_col] = pd.to_datetime(df[real_date_col], format=date_format, errors='coerce').dt.date
        df = df.dropna(subset=[real_date_col]).reset_index(drop=True)

    # Realign rows to the left if true
    if realign:
        def realign_row(row):
            non_empty = row.dropna().values
            aligned = pd.Series([None] * len(df.columns), index=df.columns, dtype=object)
            aligned[:len(non_empty)] = non_empty
            return aligned

        df = df.apply(realign_row, axis=1)

    # Adjust n_cols if tracking columns were added
    extra_cols = 0
    if track_fill and fill_cols:
        extra_cols += len(fill_cols)
    if track_date_fill and real_date_col:
        extra_cols += 1

    # Drop columns after specified number 
    if n_cols is not None:
        df = df.iloc[:, :n_cols + extra_cols]

    df.columns = filtered_columns[-len(df.columns):]
    df = df.reset_index(drop=True)

    return df

@st.cache_data
def compare_dfs(
    excel_1,
    pair1,
    excel_2,
    pair2,
    compare1=None,
    compare2=None,
    case_insensitive_match=True,
):
    """
    Generalized Excel comparison function.
    """

    # Work on copies so original DataFrames aren't modified
    excel_1 = excel_1.copy()
    excel_2 = excel_2.copy()

    def resolve_column(df, name):
        lower_map = {col.lower(): col for col in df.columns}
        name_lower = name.lower()
        if name_lower not in lower_map:
            raise ValueError(f"Column '{name}' not found in DataFrame columns: {list(df.columns)}")
        return lower_map[name_lower]

    def resolve_column_list(df, names):
        return [resolve_column(df, name) for name in names]

    # Normalize column names
    pair1 = resolve_column(excel_1, pair1)
    pair2 = resolve_column(excel_2, pair2)

    # Convert single column to list
    if isinstance(compare1, str):
        compare1 = [compare1]
    if isinstance(compare2, str):
        compare2 = [compare2]

    if (compare1 is None) != (compare2 is None):
        raise ValueError("Both compare1 and compare2 must be provided together, or not at all.")
    if compare1 and compare2:
        if not isinstance(compare1, list) or not isinstance(compare2, list):
            raise TypeError("compare1 and compare2 must be lists.")
        if len(compare1) != len(compare2):
            raise ValueError("compare1 and compare2 must be the same length.")

        compare1 = resolve_column_list(excel_1, compare1)
        compare2 = resolve_column_list(excel_2, compare2)

    # Create comparison keys
    if case_insensitive_match:
        excel_1['_key'] = excel_1[pair1].astype(str).str.lower()
        excel_2['_key'] = excel_2[pair2].astype(str).str.lower()
    else:
        excel_1['_key'] = excel_1[pair1].astype(str)
        excel_2['_key'] = excel_2[pair2].astype(str)

    # Find unmatched rows
    filter1 = excel_1['_key'].isin(excel_2['_key'])
    missing_from_excel_2 = excel_1[~filter1]

    filter2 = excel_2['_key'].isin(excel_1['_key'])
    missing_from_excel_1 = excel_2[~filter2]

    # Build unique key list
    id_list = set(excel_1['_key']).union(set(excel_2['_key']))
    diff_qty = {}
    combined_list = []

    for key in id_list:
        row_1 = excel_1[excel_1['_key'] == key]
        row_2 = excel_2[excel_2['_key'] == key]

        if row_1.empty or row_2.empty:
            combined_list.append(key)
            continue

        if compare1 and compare2:
            for col1, col2 in zip(compare1, compare2):
                val1 = row_1.iloc[0][col1]
                val2 = row_2.iloc[0][col2]

                val1_norm = val1.strip().lower() if isinstance(val1, str) else val1
                val2_norm = val2.strip().lower() if isinstance(val2, str) else val2

                if pd.isnull(val1_norm) and pd.isnull(val2_norm):
                    continue
                if val1_norm != val2_norm:
                    original_id = row_1.iloc[0][pair1]  # report original value
                    if original_id not in diff_qty:
                        diff_qty[original_id] = {}
                    diff_qty[original_id][f"{col1} vs {col2}"] = [val1, val2]

    # Build tidy mismatch DataFrame
    if diff_qty:
        records = []
        for key, mismatches in diff_qty.items():
            for col_pair, values in mismatches.items():
                col1_name, col2_name = col_pair.split(" vs ")
                val1, val2 = values
                records.append({
                    "ID": key,
                    "Column (Excel 1 | Excel2)": f"{col1_name} | {col2_name}",
                    "Excel 1": val1,
                    "Excel 2": val2
                })
        diff_qty_df = pd.DataFrame(records)
        diff_qty_df.sort_values(by=["ID", "Column (Excel 1 | Excel2)"], inplace=True)
    else:
        diff_qty_df = pd.DataFrame()

    # Convert unmatched keys to original IDs
    original_ids = []
    for key in combined_list:
        match = excel_1[excel_1['_key'] == key]
        if not match.empty:
            original_ids.append(match.iloc[0][pair1])
        else:
            match = excel_2[excel_2['_key'] == key]
            original_ids.append(match.iloc[0][pair2])

    combo = pd.DataFrame(original_ids, columns=['Missing list']).sort_values(
        by='Missing list', ascending=False
    )

    # Cleanup original DataFrames (copies)
    excel_1.drop(columns=['_key'], inplace=True, errors='ignore')
    excel_2.drop(columns=['_key'], inplace=True, errors='ignore')
    
    # Also remove _key from the return values
    missing_from_excel_1 = missing_from_excel_1.drop(columns=['_key'], errors='ignore')
    missing_from_excel_2 = missing_from_excel_2.drop(columns=['_key'], errors='ignore')

    return missing_from_excel_1, missing_from_excel_2, diff_qty_df, combo

@st.cache_data
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# -----------------------------
# Reusable cleaning UI section
# -----------------------------
def excel_cleaning_section(section_title: str, state_prefix: str):
    """
    Render the interactive cleaning UI for one Excel file.
    Returns the cleaned DataFrame (or None).
    Uses `state_prefix` to keep Streamlit keys distinct
    across the first and second Excel.
    """

    # Initialize per-section state
    config_key = f"{state_prefix}_config_expanded"
    cleaned_key = f"{state_prefix}_cleaned_df"

    if config_key not in st.session_state:
        st.session_state[config_key] = True
    if cleaned_key not in st.session_state:
        st.session_state[cleaned_key] = None

    st.markdown(f"### {section_title}")

    # 0. Ask if the Excel is already clean
    isClean = st.radio(
        f"Does this Excel for **{section_title}** look clean?",
        [":rainbow[Yes]", "No"],
        index=None,
        key=f"isClean_{state_prefix}",
        captions=[
            "The Excel is clean",
            "It needs to be cleaned",
        ],
    )

    if isClean is None:
        st.warning("Please select an option above.")
        return st.session_state[cleaned_key]

    # PATH A: Excel is already clean → simple upload & show
    if isClean == ":rainbow[Yes]":
        uploaded_file_clean = st.file_uploader(
            f"Upload the clean Excel for {section_title}",
            type=["xlsx", "xls"],
            key=f"file_clean_direct_{state_prefix}",
        )

        if uploaded_file_clean is not None:
            try:
                df_clean = pd.read_excel(uploaded_file_clean)
            except Exception as e:
                st.error(f"Could not read the Excel file: {e}")
                return st.session_state[cleaned_key]

            st.session_state[cleaned_key] = df_clean
            st.session_state[f"df_{state_prefix}"] = df_clean

            st.success("Here is your DataFrame:")
            st.dataframe(df_clean.reset_index(drop=True), hide_index=True)

            excel_bytes = to_excel(df_clean)
            st.download_button(
                f"📥 Download Excel for {section_title}",
                data=excel_bytes,
                file_name=f"cleaned_{state_prefix}.xlsx",
                key=f"download_clean_{state_prefix}",
            )
        else:
            st.warning("Please upload a file.")

        return st.session_state[cleaned_key]

    # PATH B: Needs cleaning
    with st.expander(f"⚙️ Cleaning setup – {section_title}", expanded=st.session_state[config_key]):

        # Step 0: upload
        uploaded_file = st.file_uploader(
            f"Upload the Excel file to clean for {section_title}",
            type=["xlsx", "xls"],
            key=f"clean_excel_file_{state_prefix}",
        )

        if uploaded_file is None:
            st.info("Upload an Excel file to get started.")
            return st.session_state[cleaned_key]

        # Step 1: sheet
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheet_name = st.selectbox(
                "Which sheet should be cleaned?",
                xls.sheet_names,
                key=f"sheet_{state_prefix}",
            )
        except Exception as e:
            st.error(f"Could not read sheet names: {e}")
            return st.session_state[cleaned_key]

        st.markdown("#### Step 1 – What’s wrong with this sheet?")
        st.caption("Check everything that applies. I’ll only ask follow-up questions for the items you select.")

        issue_options = [
            "Header row is not on row 1",
            "Skip extra rows at the bottom",
            "Start from a specific column letter (drop left-side junk)",
            "There is a date column to clean/standardize",
            "Drop rows containing certain keywords",
            "Limit to first N columns",
            "Fill missing values in non-date columns",
            "Track which values were filled",
            "Track which dates were filled",
            "Do NOT realign rows (keep original alignment)",
        ]

        issues = st.multiselect(
            "Select all that apply:",
            options=issue_options,
            default=[],
            key=f"issues_{state_prefix}",
        )

        # Step 2: header row
        st.markdown("#### Step 2 – Header row")
        st.caption(
            "Tell me where the column headers are. "
            "If they’re already on the first row of the sheet, leave this as 1."
        )

        header_row_guess = st.number_input(
            "Which row contains the column headers? (1-based)",
            min_value=1,
            value=1,
            step=1,
            key=f"header_row_{state_prefix}",
        )

        # Realign rows?
        realign = "Do NOT realign rows (keep original alignment)" not in issues

        # Step 3: preview
        st.markdown("#### Step 3 – Preview")

        st.markdown(
            """
**How to use this preview**

- **Tip** - When answering questions about the row number or column number, look at the **original excel instead of preview** for most accurate number. 
- If a **date column** shows anything other than dates (for example rows with words like `Totals`, `Grand Totals`, or other notes), you can:
  - Remove those rows later using **Drop rows by keywords** in Step 4, or  
  - Enforce a specific **date format** in the **Date column options** so that non-date values become empty and get dropped.
  - You may have to **go back to the previous step** to add those options in the drop down menu if they were not previously selected.
- You can **ignore any `Unnamed:` columns** in this preview – those extra columns will be cleaned up later in the process.
- **Any blank cells** in columns you want to **keep** must be filled with something for the cleaning process. Select **Fill missing values...** from the dropdown menu to accomplish this if needed. 
- The preview is mainly for:
  1. Confirming you picked the correct **header row**, and  
  2. Spot-checking that the **date column**, other column values, and other headings look reasonable (no obvious junk or missing values).
"""
        )

        columns_preview = []
        show_preview = st.checkbox(
            "Show preview using this header row",
            key=f"show_preview_{state_prefix}",
        )

        if show_preview:
            try:
                preview_df = pd.read_excel(
                    uploaded_file,
                    sheet_name=sheet_name,
                    skiprows=header_row_guess - 1,
                    nrows=100,
                )
                st.caption("Preview (first ~100 rows) using the chosen header row:")
                st.dataframe(preview_df.reset_index(drop=True), hide_index=True)
                columns_preview = list(preview_df.columns)
            except Exception as e:
                st.error(f"Could not preview the sheet with these settings: {e}")
                columns_preview = []
        else:
            st.info(
                "After setting the header row, you can turn on the preview checkbox above "
                "to see a sample of the data."
            )

        # Defaults for options
        date_col = None
        date_format = None
        date_fill_method = "ffill"
        track_date_fill = False
        drop_keywords_list = None
        n_cols = None
        skip_last_rows = False
        skip_num = 0
        start_col_letter = None
        fill_cols_list = None
        fill_methods_arg = None
        track_fill = False

        # 4a. Extra rows at bottom
        if "Skip extra rows at the bottom" in issues:
            st.markdown("##### Skip extra rows at the bottom")
            st.caption(
                "Use this if there are totals, notes, or other junk rows after the useful data."
            )
            skip_num = st.number_input(
                "What is the final row number you want to KEEP in the sheet?",
                min_value=1,
                step=1,
                help="For example, if the useful data ends on row 200, enter 200.",
                key=f"skip_num_{state_prefix}",
            )
            skip_last_rows = True

        # 4b. Start from specific column letter
        if "Start from a specific column letter (drop left-side junk)" in issues:
            st.markdown("##### Start from a specific column letter")
            start_col_letter = st.text_input(
                "Enter column letter to start from (e.g., C, D, F)",
                max_chars=3,
                help="Columns before this will be dropped.",
                key=f"start_col_{state_prefix}",
            ) or None

        # 4c. Date column handling
        if "There is a date column to clean/standardize" in issues:
            st.markdown("##### Date column options")

            # Choose date column via dropdown if preview is available
            if columns_preview:
                date_col_choice = st.selectbox(
                    "Which column is the date column?",
                    options=["Auto-detect from columns"] + columns_preview,
                    index=0,
                    key=f"date_col_choice_{state_prefix}",
                    help="Choose a column or let the function try to auto-detect one with 'date' in its name.",
                )
                if date_col_choice != "Auto-detect from columns":
                    date_col = date_col_choice
                else:
                    date_col = None
            else:
                date_col = st.text_input(
                    "Date column name (or leave blank to auto-detect)",
                    help="This should exactly match the column name after preview cleaning.",
                    key=f"date_col_text_{state_prefix}",
                ) or None

            # Date format dropdown
            st.caption(
                """
If the date column contains text like `Totals`, `Grand Totals`, or other non-date values, you have two tools:

1. **Drop rows by keywords** (in a later section) to remove those rows completely.
2. **Force a specific date format** here. Any value that doesn't match this format will be coerced to an empty date, and those rows can then be safely dropped.

If you leave this as **“Let pandas infer from the data”**, it will try its best automatically, but you may end up with a few extra blank or strange dates if the column is very messy. In those cases, manually selecting the correct format is usually best.
"""
            )
            date_format_choice = st.selectbox(
                "Date format example",
                options=[
                    "Let pandas infer from the data",
                    "2024-01-31 (YYYY-MM-DD)",
                    "01/31/2024 (MM/DD/YYYY)",
                    "31/01/2024 (DD/MM/YYYY)",
                    "December 2, 2025 (Month D, YYYY)",
                    "Custom format...",
                ],
                index=0,
                key=f"date_format_choice_{state_prefix}",
            )

            if date_format_choice == "Let pandas infer from the data":
                date_format = None
            elif date_format_choice == "2024-01-31 (YYYY-MM-DD)":
                date_format = "%Y-%m-%d"
            elif date_format_choice == "01/31/2024 (MM/DD/YYYY)":
                date_format = "%m/%d/%Y"
            elif date_format_choice == "31/01/2024 (DD/MM/YYYY)":
                date_format = "%d/%m/%Y"
            elif date_format_choice == "December 2, 2025 (Month D, YYYY)":
                date_format = "%B %d, %Y"
                st.caption(
                    "⚠️ If your data has ordinals like 'December 2nd 2025', "
                    "it’s usually best to choose 'Let pandas infer from the data' instead, "
                    "since explicit formats can’t represent the 'nd/th' parts."
                )
            else:
                custom_fmt = st.text_input(
                    "Enter custom date format",
                    placeholder="e.g. %Y/%m/%d or %d-%b-%Y",
                    help=(
                        "Use Python datetime format codes, like %Y-%m-%d. "
                        "Helpful source [here](https://docs.python.org/3/library/datetime.html#strftime-and-strptime-behavior)."
                    ),
                    key=f"custom_date_fmt_{state_prefix}",
                )
                date_format = custom_fmt or None

            date_fill_method = st.radio(
                "How should missing dates be filled?",
                options=["ffill", "bfill"],
                horizontal=True,
                key=f"date_fill_method_{state_prefix}",
            )

            st.caption(
                """
            - **ffill** (forward fill): fill each empty date with the last non-empty date **above** it.  
            - **bfill** (backward fill): fill each empty date with the next non-empty date **below** it.
            """
            )

        # 4d. Track which dates were filled
        if "Track which dates were filled" in issues:
            track_date_fill = True

        # 4e. Drop rows with keywords
        if "Drop rows containing certain keywords" in issues:
            st.markdown("##### Drop rows by keywords")
            st.caption(
                """
Use this when the date column (or any other column) has extra rows like subtotals, grand totals, notes, or headers.

- Matching is **case-insensitive**.
- It looks for the keyword **anywhere in the row**.
- Typing just `totals` will remove rows containing **`Totals`**, **`Grand Totals`**, **`Project Totals`**, etc.
"""
            )
            kw_text = st.text_area(
                "Enter keywords to drop (comma-separated)",
                placeholder="e.g. totals, notes, header",
                key=f"drop_keywords_{state_prefix}",
            )
            drop_keywords_list = [
                kw.strip() for kw in kw_text.split(",") if kw.strip()
            ] or None

        # 4f. Limit to first N columns
        if "Limit to first N columns" in issues:
            st.markdown("##### Limit to first N columns")
            n_cols_val = st.number_input(
                "How many columns do you want to keep (from the left)?",
                min_value=1,
                step=1,
                key=f"n_cols_{state_prefix}",
            )
            n_cols = int(n_cols_val)

        # 4g. Fill missing values in non-date columns
        if "Fill missing values in non-date columns" in issues:
            st.markdown("##### Fill missing values in non-date columns")

            if columns_preview:
                fill_cols_list = st.multiselect(
                    "Which columns should be filled?",
                    options=columns_preview,
                    help="These columns will have missing values filled.",
                    key=f"fill_cols_{state_prefix}",
                )
            else:
                cols_text = st.text_input(
                    "Column names to fill (comma-separated)",
                    placeholder="e.g. UNIT OF MEASURE, tax",
                    key=f"fill_cols_text_{state_prefix}",
                )
                fill_cols_list = [
                    c.strip() for c in cols_text.split(",") if c.strip()
                ]

            fill_methods_dict = {}

            if fill_cols_list:
                st.caption("Configure how each selected column should be filled:")
                for col in fill_cols_list:
                    with st.expander(f"Fill options for '{col}'", expanded=False):
                        method_choice = st.radio(
                            f"Method for '{col}'",
                            options=[
                                "Forward fill (ffill)",
                                "Backward fill (bfill)",
                                "Use a constant value",
                            ],
                            key=f"fill_method_{state_prefix}_{col}",
                        )

                        st.caption(
                            """
                        - **ffill** (forward fill): fill each empty date with the last non-empty date **above** it.  
                        - **bfill** (backward fill): fill each empty date with the next non-empty date **below** it.
                        """
                        )

                        if method_choice == "Forward fill (ffill)":
                            fill_methods_dict[col] = "ffill"
                        elif method_choice == "Backward fill (bfill)":
                            fill_methods_dict[col] = "bfill"
                        else:
                            const_val = st.text_input(
                                f"Constant value for '{col}'",
                                key=f"fill_const_{state_prefix}_{col}",
                                placeholder="e.g. Unknown, 0, No",
                            )
                            fill_methods_dict[col] = const_val

                fill_methods_arg = fill_methods_dict

        # 4h. Track which non-date values were filled
        if "Track which values were filled" in issues:
            track_fill = True

        # Step 5: Run cleaning
        st.markdown("----")
        run_clean = st.button(
            f"🚀 Run cleaning for {section_title}",
            key=f"run_clean_{state_prefix}",
        )

        if run_clean:
            if "Skip extra rows at the bottom" in issues and skip_last_rows and skip_num <= 0:
                st.error("You selected 'Skip extra rows at the bottom' but did not provide a valid final row.")
                return st.session_state[cleaned_key]

            if "Fill missing values in non-date columns" in issues and (not fill_cols_list):
                st.error("You selected to fill missing values, but no columns were selected.")
                return st.session_state[cleaned_key]

            with st.spinner("Cleaning Excel..."):
                try:
                    cleaned_df = clean_excel(
                        file=uploaded_file,
                        header_row_guess=int(header_row_guess),
                        sheet_name=sheet_name,
                        date_col=date_col,
                        date_format=date_format,
                        date_fill_method=date_fill_method,
                        track_date_fill=track_date_fill,
                        drop_keywords=drop_keywords_list,
                        realign=realign,
                        n_cols=n_cols,
                        skip_last_rows=skip_last_rows,
                        skip_num=int(skip_num),
                        start_col_letter=start_col_letter,
                        fill_cols=fill_cols_list,
                        fill_methods=fill_methods_arg if fill_methods_arg is not None else "ffill",
                        track_fill=track_fill,
                    )
                except Exception as e:
                    st.error(f"Something went wrong while cleaning: {e}")
                    return st.session_state[cleaned_key]

            st.session_state[cleaned_key] = cleaned_df
            st.session_state[f"df_{state_prefix}"] = cleaned_df
            st.session_state[config_key] = False

            # Rerun to collapse the expander and show the final df neatly
            try:
                st.rerun()
            except AttributeError:
                pass

    # Outside the expander: show final df (if any)
    if st.session_state[cleaned_key] is not None:
        st.success(f"Here is your cleaned DataFrame for {section_title}:")
        st.dataframe(st.session_state[cleaned_key].reset_index(drop=True), hide_index=True)

        excel_bytes = to_excel(st.session_state[cleaned_key])
        st.download_button(
            f"📥 Download cleaned Excel for {section_title}",
            data=excel_bytes,
            file_name=f"cleaned_{state_prefix}.xlsx",
            key=f"download_cleaned_{state_prefix}",
        )

    return st.session_state[cleaned_key]

# -----------------------------
# Page setup & main layout
# -----------------------------
st.set_page_config(
    page_title="General Excel App",
    page_icon="⎘",
)

def wide_space_default():
    st.set_page_config(layout="wide")

wide_space_default()

col1, col2, col3 = st.columns([1, 1.5, 1])

with col2:
    st.image("pages/gen-gif.gif")

st.markdown("<h1 style='text-align: center'>General Excel Cleaning and Comparison App</h1>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Navigation")
    st.caption("Use the menu above to switch between pages.")
    st.caption("You're on the **General App** page. Use this tool for standard Excel cleaning and comparison tasks.")

    with st.expander("ℹ️ How this page works", expanded=False):
        st.markdown(
            """
            This page provides a flexible version of the Excel comparison tool.

            **General Workflow**
            1. Upload **Excel File 1**.  
            2. Upload **Excel File 2**.  
            3. Choose the columns to match between the two files.  
            4. (Optional) Select additional columns to compare for differences.  
            5. Review missing or mismatched entries.  
            6. Download results as an Excel file.

            This version is ideal for workflows **not tied to Aftermath formatting**.
            """
        )


st.markdown(
    """
    This General Excel Cleaning and Comparison web application, similar to the Aftermath Disaster Recovery app, processes and cleans multiple Excel files before comparing them. Unlike the Aftermath version, this application is designed to be more flexible and general, catering to a wider range of use cases. This allows users to apply the tool to various scenarios beyond a single specific case, but it requires more inputs from the user.

    While this application can clean a broader spectrum of dirty Excel files compared to the Aftermath version, it is still limited to cleaning Excel files in an invoice format or similar structures. If you encounter any issues or have questions, refer to the tutorial for more information.

    In addition to its cleaning function, this application enables users to compare information between two clean Excel files as long as they share a unique identifier column (e.g., a specific product code). These clean Excel files can be either cleaned through the earlier process or provided by the user. Furthermore, the comparison feature allows users to compare multiple columns between each Excel if desired. 

    👈👈 Use the sidebar if you would like to see a **:rainbow[tutorial]**, have questions, or want to see the app for Aftermath Disaster Recovery.
"""
)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("<h2 style='text-align: center'>Excel Uploads and Questions</h2>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

st.markdown(
    """
   In this section, you can upload either Excel files that require cleaning or already cleaned Excel files to compare. Afterward, you'll have the option to download the cleaned Excel files. Finally, you'll be able to view the comparison information and download an aggregated Excel file.
"""
)

col1, col2, col3 = st.columns([1, 6, 1])

with col1:
    st.write("")

with col2:

    # 1. Clean first Excel
    df1 = excel_cleaning_section("First Excel (Source A)", "excel1")

    st.markdown("---")

    # 2. Clean second Excel
    df2 = excel_cleaning_section("Second Excel (Source B)", "excel2")

    st.markdown("---")
    st.markdown("### Step 3 – Compare the two Excels")

    with st.expander("🔍 Comparison options", expanded=True):
        if df1 is not None and df2 is not None:
            st.write("Select the key columns and (optionally) value columns to compare. You can tweak this later.")

            key_col_1 = st.selectbox(
                "Key / ID column in First Excel",
                df1.columns,
                key="key_col_1",
            )
            key_col_2 = st.selectbox(
                "Key / ID column in Second Excel",
                df2.columns,
                key="key_col_2",
            )

            compare_cols_1 = st.multiselect(
                "Columns to compare from First Excel (optional)",
                df1.columns,
                key="compare_cols_1",
            )
            compare_cols_2 = st.multiselect(
                "Columns to compare from Second Excel (optional, same length as above if used)",
                df2.columns,
                key="compare_cols_2",
            )

            # NEW: toggle for case-insensitive matching of key / ID values
            case_insensitive = st.checkbox(
                "Ignore differences in letter case when matching key / ID values",
                value=True,
                key="case_insensitive_match",
                help="If checked, IDs like 'abc123' and 'ABC123' will be treated as the same.",
            )

            if (compare_cols_1 and not compare_cols_2) or (compare_cols_2 and not compare_cols_1):
                st.warning("If you choose columns to compare, please select columns in both Excels.")
            elif compare_cols_1 and compare_cols_2 and len(compare_cols_1) != len(compare_cols_2):
                st.warning("The number of columns selected in each Excel must match.")
            else:
                compare1 = compare_cols_1 if compare_cols_1 else None
                compare2 = compare_cols_2 if compare_cols_2 else None

                if st.button("Run comparison", key="run_comparison"):
                    with st.spinner("Comparing…"):
                        missing_from_excel_1, missing_from_excel_2, diff_qty_df, combo = compare_dfs(
                            df1,
                            key_col_1,
                            df2,
                            key_col_2,
                            compare1=compare1,
                            compare2=compare2,
                            case_insensitive_match=case_insensitive,  # <-- passes toggle through
                        )

                    combo_sorted = combo.sort_values(by='Missing list')

                    # Store in session so they persist after reruns / downloads
                    st.session_state.missing_from_excel_1 = missing_from_excel_1
                    st.session_state.missing_from_excel_2 = missing_from_excel_2
                    st.session_state.diff_qty_df = diff_qty_df
                    st.session_state.combo = combo
                    st.session_state.combo_sorted = combo_sorted

            # Always show results if present (so they stay visible after downloads)
            if (
                'missing_from_excel_1' in st.session_state and
                'missing_from_excel_2' in st.session_state and
                'diff_qty_df' in st.session_state and
                'combo_sorted' in st.session_state
            ):
                st.write("**Entries missing from First Excel (Source A)**")
                st.dataframe(
                    st.session_state.missing_from_excel_1.reset_index(drop=True),
                    hide_index=True,
                )

                st.write("**Entries missing from Second Excel (Source B)**")
                st.dataframe(
                    st.session_state.missing_from_excel_2.reset_index(drop=True),
                    hide_index=True,
                )

                st.write("**Entries with different values between the two Excels**")
                st.dataframe(
                    st.session_state.diff_qty_df.reset_index(drop=True),
                    hide_index=True,
                )

                st.write("**Combined list of IDs missing from one or both Excels**")
                st.dataframe(
                    st.session_state.combo_sorted.reset_index(drop=True),
                    hide_index=True,
                )

        else:
            st.info("Please finish cleaning / loading both Excels above to enable comparison.")


    st.markdown("#### Download All Comparison Dataframes into Excel File:")

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
                missing_from_excel_1.to_excel(writer, sheet_name='missing_from_excel_1', index=False)
                sheets_written += 1

            if not missing_from_excel_2.empty:
                missing_from_excel_2.to_excel(writer, sheet_name='missing_from_excel_2', index=False)
                sheets_written += 1

            if not diff_qty_df.empty:
                diff_qty_df.to_excel(writer, sheet_name='diff_qty_df', index=False)
                sheets_written += 1

            if not combo.empty:
                combo.to_excel(writer, sheet_name='combo_missing_id', index=False)
                sheets_written += 1

            if sheets_written > 0:
                del writer.book['No_Data']

        output.seek(0)

        st.download_button(
            label="📥 Download Comparison Excel",
            data=output,
            file_name="comparison_results.xlsx"
        )
    else:
        st.info("Run the comparison above to enable the download.")

with col3:
    st.write("")
