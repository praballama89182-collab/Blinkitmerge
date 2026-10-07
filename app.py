import io
import os
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

st.set_page_config(
    page_title="Blinkit Report Merger & Analytics", page_icon="🛍️", layout="wide"
)

# Custom CSS for seamless alignment and clean table structure
st.markdown(
    """
<style>
    [data-testid="stDataFrame"] th, [data-testid="stDataFrame"] td {
        text-align: center !important;
        vertical-align: middle !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

st.title("📊 Blinkit Ad Report Merger & Analytics")
st.write(
    "Upload up to 5 monthly Excel ad campaign spreadsheets (.xlsx, .xls, .xlsb,"
    " .xlsm). Preview raw & consolidated sheets, view Month-on-Month"
    " Comparison tables (arranged latest to earlier), and analyze multi-metric trend comparisons."
)


# Helper function to format numeric values (round to 2 decimals or whole number if .00)
def format_num_val(val):
    if isinstance(val, (int, float, np.number)):
        if pd.isna(val):
            return 0
        if float(val).is_integer():
            return int(round(val))
        return float(round(val, 2))
    return val


# Helper function to apply map safely across pandas versions
def safe_cell_map(df, func):
    if hasattr(df, "map"):
        return df.map(func)
    return df.applymap(func)


# Helper function to style single downloadable Excel Pivot tables cleanly
def style_and_export_pivot(pivot_df, sheet_name="Comparison"):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]

    top_header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    top_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    sec_header_fill = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )
    sec_header_font = Font(name="Calibri", size=10, bold=True, color="1F4E78")

    index_fill = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )
    index_font = Font(name="Calibri", size=10, bold=True, color="000000")

    total_fill = PatternFill(
        start_color="E9ECEF", end_color="E9ECEF", fill_type="solid"
    )
    total_font = Font(name="Calibri", size=10, bold=True, color="000000")

    data_font = Font(name="Calibri", size=10, color="000000")

    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=False
    )
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)

    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )

    if isinstance(pivot_df.columns, pd.MultiIndex):
        metrics = pivot_df.columns.get_level_values(1)
        entity_title = pivot_df.index.name if pivot_df.index.name else ""

        ws.cell(row=1, column=1, value="").fill = top_header_fill
        ws.cell(row=1, column=1).border = thin_border

        current_col = 2
        for month in pivot_df.columns.levels[0]:
            sub_cols = [c for c in pivot_df.columns if c[0] == month]
            num_sub = len(sub_cols)
            if num_sub > 1:
                ws.merge_cells(
                    start_row=1,
                    start_column=current_col,
                    end_row=1,
                    end_column=current_col + num_sub - 1,
                )
            cell = ws.cell(row=1, column=current_col, value=str(month))
            cell.fill = top_header_fill
            cell.font = top_header_font
            cell.alignment = align_center

            for c in range(current_col, current_col + num_sub):
                ws.cell(row=1, column=c).border = thin_border
                ws.cell(row=1, column=c).fill = top_header_fill
            current_col += num_sub

        cell_a2 = ws.cell(row=2, column=1, value=str(entity_title))
        cell_a2.fill = index_fill
        cell_a2.font = index_font
        cell_a2.alignment = align_center
        cell_a2.border = thin_border

        for col_idx, metric in enumerate(metrics, start=2):
            cell = ws.cell(row=2, column=col_idx, value=str(metric))
            cell.fill = sec_header_fill
            cell.font = sec_header_font
            cell.alignment = align_center
            cell.border = thin_border

        start_data_row = 3
        for r_idx, (idx_val, row_data) in enumerate(
            pivot_df.iterrows(), start=start_data_row
        ):
            is_grand_total = str(idx_val).strip().lower() == "grand total"

            idx_cell = ws.cell(row=r_idx, column=1, value=str(idx_val))
            idx_cell.fill = total_fill if is_grand_total else index_fill
            idx_cell.font = total_font if is_grand_total else index_font
            idx_cell.alignment = align_left
            idx_cell.border = thin_border

            for c_idx, (m_col, val) in enumerate(zip(metrics, row_data), start=2):
                val_cell = ws.cell(row=r_idx, column=c_idx)
                val_cell.font = total_font if is_grand_total else data_font
                if is_grand_total:
                    val_cell.fill = total_fill
                val_cell.alignment = align_center
                val_cell.border = thin_border

                if isinstance(val, (int, float, np.number)):
                    clean_v = float(val)
                    if str(m_col).upper() == "ACOS":
                        val_cell.value = clean_v / 100.0 if clean_v > 1 else clean_v
                        val_cell.number_format = "0.00%"
                    elif clean_v.is_integer():
                        val_cell.value = int(round(clean_v))
                        val_cell.number_format = "#,##0"
                    else:
                        val_cell.value = round(clean_v, 2)
                        val_cell.number_format = "#,##0.00"
                else:
                    val_cell.value = val
    else:
        for col_idx, col_name in enumerate(pivot_df.columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=str(col_name))
            cell.fill = top_header_fill
            cell.font = top_header_font
            cell.alignment = align_center
            cell.border = thin_border

        for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=2):
            is_pct_row = str(row_data.iloc[0]).strip().lower() == "percentage %"

            for c_idx, (col_name, val) in enumerate(
                zip(pivot_df.columns, row_data), start=1
            ):
                val_cell = ws.cell(row=r_idx, column=c_idx)
                val_cell.font = total_font if is_pct_row else data_font
                val_cell.alignment = align_center
                val_cell.border = thin_border

                if (
                    str(col_name).upper() == "ACOS"
                    and not is_pct_row
                    and isinstance(val, (int, float, np.number))
                ):
                    clean_v = float(val)
                    val_cell.value = clean_v / 100.0 if clean_v > 1 else clean_v
                    val_cell.number_format = "0.00%"
                elif isinstance(val, (int, float, np.number)) and not is_pct_row:
                    clean_v = float(val)
                    if clean_v.is_integer():
                        val_cell.value = int(round(clean_v))
                        val_cell.number_format = "#,##0"
                    else:
                        val_cell.value = round(clean_v, 2)
                        val_cell.number_format = "#,##0.00"
                else:
                    val_cell.value = val

                if is_pct_row and c_idx > 1:
                    val_str = str(val).replace("%", "").replace("+", "").strip()
                    try:
                        num_v = float(val_str)
                        if num_v < 0:
                            val_cell.fill = PatternFill(
                                start_color="F8D7DA", end_color="F8D7DA", fill_type="solid"
                            )
                            val_cell.font = Font(
                                name="Calibri", size=10, bold=True, color="721C24"
                            )
                        elif num_v > 0:
                            val_cell.fill = PatternFill(
                                start_color="D4EDDA", end_color="D4EDDA", fill_type="solid"
                            )
                            val_cell.font = Font(
                                name="Calibri", size=10, bold=True, color="155724"
                            )
                    except ValueError:
                        pass
                elif is_pct_row:
                    val_cell.fill = sec_header_fill

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# Function to export all pivot tables into a multi-tab Excel Workbook
def convert_all_pivots_to_excel(pivot_dict):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    top_header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    top_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    sec_header_fill = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )
    sec_header_font = Font(name="Calibri", size=10, bold=True, color="1F4E78")
    index_fill = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )
    index_font = Font(name="Calibri", size=10, bold=True, color="000000")
    total_fill = PatternFill(
        start_color="E9ECEF", end_color="E9ECEF", fill_type="solid"
    )
    total_font = Font(name="Calibri", size=10, bold=True, color="000000")
    data_font = Font(name="Calibri", size=10, color="000000")

    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=False
    )
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)
    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )

    for sheet_name, pivot_df in pivot_dict.items():
        if pivot_df is not None and not pivot_df.empty:
            ws = wb.create_sheet(title=sheet_name[:31])

            if isinstance(pivot_df.columns, pd.MultiIndex):
                metrics = pivot_df.columns.get_level_values(1)
                entity_title = pivot_df.index.name if pivot_df.index.name else ""

                ws.cell(row=1, column=1, value="").fill = top_header_fill
                ws.cell(row=1, column=1).border = thin_border

                current_col = 2
                for month in pivot_df.columns.levels[0]:
                    sub_cols = [c for c in pivot_df.columns if c[0] == month]
                    num_sub = len(sub_cols)
                    if num_sub > 1:
                        ws.merge_cells(
                            start_row=1,
                            start_column=current_col,
                            end_row=1,
                            end_column=current_col + num_sub - 1,
                        )
                    cell = ws.cell(row=1, column=current_col, value=str(month))
                    cell.fill = top_header_fill
                    cell.font = top_header_font
                    cell.alignment = align_center

                    for c in range(current_col, current_col + num_sub):
                        ws.cell(row=1, column=c).border = thin_border
                        ws.cell(row=1, column=c).fill = top_header_fill
                    current_col += num_sub

                cell_a2 = ws.cell(row=2, column=1, value=str(entity_title))
                cell_a2.fill = index_fill
                cell_a2.font = index_font
                cell_a2.alignment = align_center
                cell_a2.border = thin_border

                for col_idx, metric in enumerate(metrics, start=2):
                    cell = ws.cell(row=2, column=col_idx, value=str(metric))
                    cell.fill = sec_header_fill
                    cell.font = sec_header_font
                    cell.alignment = align_center
                    cell.border = thin_border

                for r_idx, (idx_val, row_data) in enumerate(
                    pivot_df.iterrows(), start=3
                ):
                    is_grand_total = str(idx_val).strip().lower() == "grand total"

                    idx_cell = ws.cell(row=r_idx, column=1, value=str(idx_val))
                    idx_cell.fill = total_fill if is_grand_total else index_fill
                    idx_cell.font = total_font if is_grand_total else index_font
                    idx_cell.alignment = align_left
                    idx_cell.border = thin_border

                    for c_idx, (m_col, val) in enumerate(zip(metrics, row_data), start=2):
                        val_cell = ws.cell(row=r_idx, column=c_idx)
                        val_cell.font = total_font if is_grand_total else data_font
                        if is_grand_total:
                            val_cell.fill = total_fill
                        val_cell.alignment = align_center
                        val_cell.border = thin_border

                        if isinstance(val, (int, float, np.number)):
                            clean_v = float(val)
                            if str(m_col).upper() == "ACOS":
                                val_cell.value = clean_v / 100.0 if clean_v > 1 else clean_v
                                val_cell.number_format = "0.00%"
                            elif clean_v.is_integer():
                                val_cell.value = int(round(clean_v))
                                val_cell.number_format = "#,##0"
                            else:
                                val_cell.value = round(clean_v, 2)
                                val_cell.number_format = "#,##0.00"
                        else:
                            val_cell.value = val
            else:
                for col_idx, col_name in enumerate(pivot_df.columns, start=1):
                    cell = ws.cell(row=1, column=col_idx, value=str(col_name))
                    cell.fill = top_header_fill
                    cell.font = top_header_font
                    cell.alignment = align_center
                    cell.border = thin_border

                for r_idx, (idx_val, row_data) in enumerate(
                    pivot_df.iterrows(), start=2
                ):
                    is_pct_row = str(row_data.iloc[0]).strip().lower() == "percentage %"
                    for c_idx, (col_name, val) in enumerate(
                        zip(pivot_df.columns, row_data), start=1
                    ):
                        val_cell = ws.cell(row=r_idx, column=c_idx)
                        val_cell.font = total_font if is_pct_row else data_font
                        val_cell.alignment = align_center
                        val_cell.border = thin_border

                        if (
                            str(col_name).upper() == "ACOS"
                            and not is_pct_row
                            and isinstance(val, (int, float, np.number))
                        ):
                            clean_v = float(val)
                            val_cell.value = clean_v / 100.0 if clean_v > 1 else clean_v
                            val_cell.number_format = "0.00%"
                        elif isinstance(val, (int, float, np.number)) and not is_pct_row:
                            clean_v = float(val)
                            if clean_v.is_integer():
                                val_cell.value = int(round(clean_v))
                                val_cell.number_format = "#,##0"
                            else:
                                val_cell.value = round(clean_v, 2)
                                val_cell.number_format = "#,##0.00"
                        else:
                            val_cell.value = val

                        if is_pct_row and c_idx > 1:
                            val_str = str(val).replace("%", "").replace("+", "").strip()
                            try:
                                num_v = float(val_str)
                                if num_v < 0:
                                    val_cell.fill = PatternFill(
                                        start_color="F8D7DA",
                                        end_color="F8D7DA",
                                        fill_type="solid",
                                    )
                                    val_cell.font = Font(
                                        name="Calibri", size=10, bold=True, color="721C24"
                                    )
                                elif num_v > 0:
                                    val_cell.fill = PatternFill(
                                        start_color="D4EDDA",
                                        end_color="D4EDDA",
                                        fill_type="solid",
                                    )
                                    val_cell.font = Font(
                                        name="Calibri", size=10, bold=True, color="155724"
                                    )
                            except ValueError:
                                pass
                        elif is_pct_row:
                            val_cell.fill = sec_header_fill

            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# File Upload Section
col1, col2, col3, col4, col5 = st.columns(5)
uploaded_files = []

with col1:
    f1 = st.file_uploader(
        "Upload File 1", type=["xlsx", "xls", "xlsb", "xlsm"], key="file1"
    )
    if f1:
        uploaded_files.append(f1)

with col2:
    f2 = st.file_uploader(
        "Upload File 2", type=["xlsx", "xls", "xlsb", "xlsm"], key="file2"
    )
    if f2:
        uploaded_files.append(f2)

with col3:
    f3 = st.file_uploader(
        "Upload File 3", type=["xlsx", "xls", "xlsb", "xlsm"], key="file3"
    )
    if f3:
        uploaded_files.append(f3)

with col4:
    f4 = st.file_uploader(
        "Upload File 4", type=["xlsx", "xls", "xlsb", "xlsm"], key="file4"
    )
    if f4:
        uploaded_files.append(f4)

with col5:
    f5 = st.file_uploader(
        "Upload File 5", type=["xlsx", "xls", "xlsb", "xlsm"], key="file5"
    )
    if f5:
        uploaded_files.append(f5)

if uploaded_files:
    consolidated_dfs = []
    raw_files_dict = {}
    consolidated_raw_tabs = {}

    for uploaded_file in uploaded_files:
        fallback_month_name = os.path.splitext(uploaded_file.name)[0].upper()
        raw_files_dict[uploaded_file.name] = {}

        try:
            xls = pd.ExcelFile(uploaded_file)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                raw_files_dict[uploaded_file.name][sheet_name] = df.copy()

                date_col = None
                for col_candidate in ["Date", "date", "Day", "DATE"]:
                    if col_candidate in df.columns:
                        date_col = col_candidate
                        break

                if date_col:
                    dt_series = pd.to_datetime(
                        df[date_col], dayfirst=True, errors="coerce"
                    )
                    month_series = dt_series.dt.strftime("%B").str.upper()
                    df["Month"] = month_series.fillna(fallback_month_name)
                else:
                    df["Month"] = fallback_month_name

                df_raw = df.copy()
                if sheet_name not in consolidated_raw_tabs:
                    consolidated_raw_tabs[sheet_name] = []
                consolidated_raw_tabs[sheet_name].append(df_raw)

                df_consolidated = df.copy()
                if "Month" in df_consolidated.columns:
                    col_month = df_consolidated.pop("Month")
                    df_consolidated.insert(0, "Month", col_month)
                else:
                    df_consolidated.insert(0, "Month", fallback_month_name)

                df_consolidated.insert(1, "Tab Name", sheet_name)
                consolidated_dfs.append(df_consolidated)
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")

    if consolidated_dfs:
        final_df = pd.concat(consolidated_dfs, ignore_index=True)

        base_cols = ["Month", "Tab Name"]
        product_listing_cols = []
        for df in consolidated_dfs:
            if "PRODUCT_LISTING" in df["Tab Name"].values:
                pl_df = df[df["Tab Name"] == "PRODUCT_LISTING"]
                product_listing_cols = [
                    c for c in pl_df.columns if c not in base_cols
                ]
                break

        remaining_cols = [
            c
            for c in final_df.columns
            if c not in base_cols and c not in product_listing_cols
        ]
        final_df = final_df.reindex(
            columns=base_cols + product_listing_cols + remaining_cols
        )

        match_col = "Match Type" if "Match Type" in final_df.columns else None
        target_col = "Targeting Type" if "Targeting Type" in final_df.columns else None
        tab_col = "Tab Name" if "Tab Name" in final_df.columns else None

        def is_na_series(series):
            if series is None or series.empty:
                return pd.Series(True, index=final_df.index)
            cleaned = series.astype(str).str.strip().str.upper()
            return series.isna() | cleaned.isin(["NA", "N/A", "NAN", "NONE", ""])

        final_df["_filled_fallback"] = False

        if match_col:
            na_match = is_na_series(final_df[match_col])
            if target_col:
                valid_target = ~is_na_series(final_df[target_col])
                fill_from_target = na_match & valid_target
                final_df.loc[fill_from_target, match_col] = final_df.loc[
                    fill_from_target, target_col
                ]
                final_df.loc[fill_from_target, "_filled_fallback"] = True
                na_match = is_na_series(final_df[match_col])

            if tab_col:
                fill_from_tab = na_match
                final_df.loc[fill_from_tab, match_col] = final_df.loc[
                    fill_from_tab, tab_col
                ]
                final_df.loc[fill_from_tab, "_filled_fallback"] = True
        elif target_col:
            final_df["Match Type"] = final_df[target_col]
            na_match = is_na_series(final_df["Match Type"])
            if tab_col:
                final_df.loc[na_match, "Match Type"] = final_df.loc[na_match, tab_col]
                final_df.loc[na_match, "_filled_fallback"] = True
            match_col = "Match Type"

        def get_numeric_col(df, possible_cols):
            for col in possible_cols:
                if col in df.columns:
                    return pd.to_numeric(df[col], errors="coerce").fillna(0)
            return pd.Series(0, index=df.index)

        final_df["_impressions"] = get_numeric_col(final_df, ["Impressions"])
        final_df["_direct_atc"] = get_numeric_col(final_df, ["Direct ATC"])
        final_df["_indirect_atc"] = get_numeric_col(final_df, ["Indirect ATC"])
        final_df["_atc"] = final_df["_direct_atc"] + final_df["_indirect_atc"]

        final_df["_direct_orders"] = get_numeric_col(
            final_df, ["Direct Quantities Sold", "Direct Orders"]
        )
        final_df["_indirect_orders"] = get_numeric_col(
            final_df, ["Indirect Quantities Sold", "Indirect Orders"]
        )
        final_df["_orders"] = (
            final_df["_direct_orders"] + final_df["_indirect_orders"]
        )

        final_df["_direct_sales"] = get_numeric_col(final_df, ["Direct Sales"])
        final_df["_indirect_sales"] = get_numeric_col(final_df, ["Indirect Sales"])
        final_df["_sales"] = (
            final_df["_direct_sales"] + final_df["_indirect_sales"]
        )

        final_df["_budget_consumed"] = get_numeric_col(
            final_df, ["Estimated Budget Consumed", "Budget Consumed", "Spend"]
        )

        if match_col and match_col in final_df.columns:
            final_df["Ad Type Combined"] = (
                final_df[match_col].fillna("Other").astype(str).str.title()
            )
        else:
            final_df["Ad Type Combined"] = "Other"

        date_col = None
        for col_candidate in ["Date", "date", "Day", "DATE"]:
            if col_candidate in final_df.columns:
                date_col = col_candidate
                break

        if date_col:
            final_df["_date_dt"] = pd.to_datetime(
                final_df[date_col], dayfirst=True, errors="coerce"
            )

            def assign_week_formatted(row):
                dt = row["_date_dt"]
                if pd.isna(dt):
                    return np.nan
                day = dt.day
                if 1 <= day <= 7:
                    return "Week 1 (1-7)"
                elif 8 <= day <= 14:
                    return "Week 2 (8-14)"
                elif 15 <= day <= 21:
                    return "Week 3 (15-21)"
                elif 22 <= day <= 28:
                    return "Week 4 (22-28)"
                elif day >= 29:
                    return "Week 5 (29-31)"
                return np.nan

            final_df["Week"] = final_df.apply(assign_week_formatted, axis=1)
        else:
            final_df["Week"] = np.nan

        # Creates MoM pivot with descending month order (Latest Month First)
        def create_mom_comparison_table(df_input, entity_col):
            if entity_col not in df_input.columns:
                return pd.DataFrame()

            working_df = df_input.dropna(subset=[entity_col, "Month"]).copy()
            if working_df.empty:
                return pd.DataFrame()

            grouped = (
                working_df.groupby([entity_col, "Month"])
                .agg(
                    Impressions=("_impressions", "sum"),
                    ATC=("_atc", "sum"),
                    Orders=("_orders", "sum"),
                    Spends=("_budget_consumed", "sum"),
                    Sales=("_sales", "sum"),
                )
                .reset_index()
            )

            grouped["CPM"] = grouped.apply(
                lambda r: round((r["Spends"] / r["Impressions"]) * 1000)
                if r["Impressions"] > 0
                else 0,
                axis=1,
            )
            grouped["ROAS"] = grouped.apply(
                lambda r: round(r["Sales"] / r["Spends"], 2)
                if r["Spends"] > 0
                else 0.0,
                axis=1,
            )
            grouped["ACOS"] = grouped.apply(
                lambda r: round((r["Spends"] / r["Sales"]) * 100, 2)
                if r["Sales"] > 0
                else 0.0,
                axis=1,
            )

            pivot_df = grouped.pivot(
                index=entity_col,
                columns="Month",
                values=[
                    "Impressions",
                    "CPM",
                    "ATC",
                    "Orders",
                    "Spends",
                    "Sales",
                    "ROAS",
                    "ACOS",
                ],
            )

            metrics_order = [
                "Impressions",
                "CPM",
                "ATC",
                "Orders",
                "Spends",
                "Sales",
                "ROAS",
                "ACOS",
            ]

            descending_months = list(df_input["Month"].unique())[::-1]

            pivot_df = pivot_df.reorder_levels([1, 0], axis=1)
            sorted_cols = pd.MultiIndex.from_product(
                [descending_months, metrics_order], names=["Month", "Metric"]
            )
            pivot_df = pivot_df.reindex(columns=sorted_cols).fillna(0)

            grand_total_series = {}
            for month in descending_months:
                month_df = working_df[working_df["Month"] == month]
                total_imp = month_df["_impressions"].sum()
                total_atc = month_df["_atc"].sum()
                total_orders = month_df["_orders"].sum()
                total_spends = month_df["_budget_consumed"].sum()
                total_sales = month_df["_sales"].sum()

                total_cpm = (
                    round((total_spends / total_imp) * 1000) if total_imp > 0 else 0
                )
                total_roas = (
                    round(total_sales / total_spends, 2) if total_spends > 0 else 0.0
                )
                total_acos = (
                    round((total_spends / total_sales) * 100, 2)
                    if total_sales > 0
                    else 0.0
                )

                grand_total_series[(month, "Impressions")] = total_imp
                grand_total_series[(month, "CPM")] = total_cpm
                grand_total_series[(month, "ATC")] = total_atc
                grand_total_series[(month, "Orders")] = total_orders
                grand_total_series[(month, "Spends")] = total_spends
                grand_total_series[(month, "Sales")] = total_sales
                grand_total_series[(month, "ROAS")] = total_roas
                grand_total_series[(month, "ACOS")] = total_acos

            pivot_df.loc["Grand Total"] = grand_total_series
            pivot_df.index.name = entity_col

            # Updated map function to fix pandas 2.1+ AttributeError
            pivot_df = safe_cell_map(pivot_df, format_num_val)
            return pivot_df

        # Creates overall monthly summary table with descending month order
        def create_monthly_summary_table(df_input):
            working_df = df_input.dropna(subset=["Month"]).copy()
            if working_df.empty:
                return pd.DataFrame()

            monthly_agg = (
                working_df.groupby("Month")
                .agg(
                    Impressions=("_impressions", "sum"),
                    ATC=("_atc", "sum"),
                    Orders=("_orders", "sum"),
                    Spends=("_budget_consumed", "sum"),
                    Sales=("_sales", "sum"),
                )
                .reset_index()
            )

            monthly_agg["CPM"] = monthly_agg.apply(
                lambda r: round((r["Spends"] / r["Impressions"]) * 1000)
                if r["Impressions"] > 0
                else 0,
                axis=1,
            )
            monthly_agg["ROAS"] = monthly_agg.apply(
                lambda r: round(r["Sales"] / r["Spends"], 2)
                if r["Spends"] > 0
                else 0.0,
                axis=1,
            )
            monthly_agg["ACOS"] = monthly_agg.apply(
                lambda r: round((r["Spends"] / r["Sales"]) * 100, 2)
                if r["Sales"] > 0
                else 0.0,
                axis=1,
            )

            descending_months = list(df_input["Month"].unique())[::-1]
            monthly_agg["month_order"] = monthly_agg["Month"].map(
                lambda x: descending_months.index(x) if x in descending_months else 99
            )
            monthly_agg = monthly_agg.sort_values("month_order").drop(
                columns=["month_order"]
            )

            col_order = [
                "Month",
                "Impressions",
                "CPM",
                "ATC",
                "Orders",
                "Spends",
                "Sales",
                "ROAS",
                "ACOS",
            ]
            monthly_agg = monthly_agg[col_order]

            if len(monthly_agg) >= 2:
                latest_row = monthly_agg.iloc[0]
                prev_row = monthly_agg.iloc[1]

                pct_row = {"Month": "Percentage %"}
                for metric in [
                    "Impressions",
                    "CPM",
                    "ATC",
                    "Orders",
                    "Spends",
                    "Sales",
                    "ROAS",
                    "ACOS",
                ]:
                    prev_val = prev_row[metric]
                    curr_val = latest_row[metric]
                    if prev_val > 0:
                        pct_change = round(((curr_val - prev_val) / prev_val) * 100)
                        pct_row[metric] = (
                            f"{pct_change}%" if pct_change <= 0 else f"+{pct_change}%"
                        )
                    else:
                        pct_row[metric] = "0%"

                monthly_agg = pd.concat(
                    [monthly_agg, pd.DataFrame([pct_row])], ignore_index=True
                )

            for col in monthly_agg.columns:
                if col != "Month":
                    monthly_agg[col] = monthly_agg[col].apply(format_num_val)

            monthly_agg["ACOS"] = monthly_agg["ACOS"].apply(
                lambda v: f"{v}%" if isinstance(v, (int, float)) else str(v)
            )
            return monthly_agg

        # Unified table renderer with FREEZED Grand Total / Percentage % pinned at the bottom during column sorting
        def render_unified_single_table(df_to_show, key_prefix="mom"):
            if df_to_show is None or df_to_show.empty:
                st.info("No data available to display for this view.")
                return

            working_df = df_to_show.copy()

            # Separate dynamic data rows from pinned summary rows
            pinned_rows = None
            if (
                isinstance(working_df.index, pd.Index)
                and "Grand Total" in working_df.index
            ):
                main_df = working_df.drop("Grand Total")
                pinned_rows = working_df.loc[["Grand Total"]]
            elif (
                "Month" in working_df.columns
                and "Percentage %" in working_df["Month"].values
            ):
                main_df = working_df[working_df["Month"] != "Percentage %"]
                pinned_rows = working_df[working_df["Month"] == "Percentage %"]
            else:
                main_df = working_df

            # Filtering & Search
            search_term = st.text_input(
                f"🔍 Search / Filter ({key_prefix})", key=f"{key_prefix}_search"
            )
            if search_term:
                if isinstance(main_df.columns, pd.MultiIndex):
                    mask = (
                        main_df.index.astype(str)
                        .str.contains(search_term, case=False, na=False)
                    )
                    main_df = main_df[mask]
                else:
                    mask = main_df.astype(str).apply(
                        lambda row: row.str.contains(
                            search_term, case=False, na=False
                        ).any(),
                        axis=1,
                    )
                    main_df = main_df[mask]

            if pinned_rows is not None and not pinned_rows.empty:
                final_display = pd.concat([main_df, pinned_rows])
            else:
                final_display = main_df

            st.dataframe(final_display, use_container_width=True)

        # Main Navigation Tabs Interface
        tab1, tab2, tab3 = st.tabs([
            "📋 Consolidated & Raw Data",
            "📊 MoM Comparisons",
            "📈 Analytics & Trends",
        ])

        # TAB 1: Consolidated Master Dataset & Raw Sheet Inspector
        with tab1:
            st.subheader("Consolidated Master Dataset")
            st.write(
                f"Combined Total Rows: **{len(final_df):,}** across"
                f" **{len(uploaded_files)}** uploaded monthly files."
            )

            search_col, dl_col = st.columns([3, 1])
            with search_col:
                filter_text = st.text_input(
                    "Filter Master Data", key="master_search", placeholder="Type campaign, SKU, keyword..."
                )
            with dl_col:
                st.write(" ")
                csv_buffer = final_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="⬇️ Export Consolidated CSV",
                    data=csv_buffer,
                    file_name="Blinkit_Consolidated_Ad_Report.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            disp_df = final_df.copy()
            if filter_text:
                mask = disp_df.astype(str).apply(
                    lambda r: r.str.contains(filter_text, case=False, na=False).any(),
                    axis=1,
                )
                disp_df = disp_df[mask]

            display_cols = [c for c in disp_df.columns if not c.startswith("_")]
            st.dataframe(disp_df[display_cols], use_container_width=True, height=400)

            st.divider()
            st.subheader("Raw Sheet Inspector")
            selected_file = st.selectbox(
                "Select File to Inspect", list(raw_files_dict.keys())
            )
            if selected_file:
                selected_sheet = st.selectbox(
                    "Select Worksheet",
                    list(raw_files_dict[selected_file].keys()),
                    key="sheet_inspect",
                )
                if selected_sheet:
                    raw_df = raw_files_dict[selected_file][selected_sheet]
                    st.dataframe(raw_df, use_container_width=True, height=300)

        # TAB 2: All 4 MoM Pivot Tables & Master Multi-Tab Excel Export
        with tab2:
            st.subheader("Month-on-Month Performance Pivots (Latest Month First)")

            monthly_summary_pivot = create_monthly_summary_table(final_df)
            campaign_mom_pivot = create_mom_comparison_table(
                final_df, "Campaign Name"
            )
            match_mom_pivot = create_mom_comparison_table(
                final_df, "Ad Type Combined"
            )
            week_mom_pivot = create_mom_comparison_table(final_df, "Week")

            all_pivots = {
                "Monthly Summary": monthly_summary_pivot,
                "Campaign MoM": campaign_mom_pivot,
                "Match Type MoM": match_mom_pivot,
                "Weekly MoM": week_mom_pivot,
            }

            excel_data = convert_all_pivots_to_excel(all_pivots)
            st.download_button(
                label="📊 Download Complete Excel MoM Report (.xlsx)",
                data=excel_data,
                file_name="Blinkit_MoM_Analytics_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

            st.divider()

            st.markdown("### 1. Overall Monthly Performance Summary")
            render_unified_single_table(monthly_summary_pivot, key_prefix="monthly_sum")

            st.markdown("### 2. Campaign-Level MoM Comparison")
            render_unified_single_table(campaign_mom_pivot, key_prefix="campaign_mom")

            st.markdown("### 3. Match / Ad Type MoM Comparison")
            render_unified_single_table(match_mom_pivot, key_prefix="match_mom")

            if not week_mom_pivot.empty:
                st.markdown("### 4. Week-Level MoM Comparison")
                render_unified_single_table(week_mom_pivot, key_prefix="week_mom")

        # TAB 3: Interactive Plotly Visual Charts & Filter Controls
        with tab3:
            st.subheader("Trend Analytics & Visual Insights")

            descending_months = list(final_df["Month"].unique())[::-1]
            m_cols = st.columns(3)
            with m_cols[0]:
                selected_month = st.selectbox("Filter Month", ["All"] + descending_months)
            with m_cols[1]:
                campaigns = (
                    ["All"] + sorted(final_df["Campaign Name"].dropna().unique().tolist())
                    if "Campaign Name" in final_df.columns
                    else ["All"]
                )
                selected_campaign = st.selectbox("Filter Campaign", campaigns)
            with m_cols[2]:
                match_types = ["All"] + sorted(
                    final_df["Ad Type Combined"].dropna().unique().tolist()
                )
                selected_match = st.selectbox("Filter Match/Ad Type", match_types)

            df_filtered = final_df.copy()
            if selected_month != "All":
                df_filtered = df_filtered[df_filtered["Month"] == selected_month]
            if selected_campaign != "All" and "Campaign Name" in df_filtered.columns:
                df_filtered = df_filtered[
                    df_filtered["Campaign Name"] == selected_campaign
                ]
            if selected_match != "All":
                df_filtered = df_filtered[
                    df_filtered["Ad Type Combined"] == selected_match
                ]

            monthly_trend = (
                final_df.groupby("Month")
                .agg(
                    Spends=("_budget_consumed", "sum"),
                    Sales=("_sales", "sum"),
                    Orders=("_orders", "sum"),
                    Impressions=("_impressions", "sum"),
                )
                .reset_index()
            )

            monthly_trend["month_order"] = monthly_trend["Month"].map(
                lambda x: descending_months.index(x) if x in descending_months else 99
            )
            monthly_trend = monthly_trend.sort_values("month_order")

            monthly_trend["ROAS"] = monthly_trend.apply(
                lambda r: round(r["Sales"] / r["Spends"], 2) if r["Spends"] > 0 else 0,
                axis=1,
            )
            monthly_trend["ACOS"] = monthly_trend.apply(
                lambda r: round((r["Spends"] / r["Sales"]) * 100, 2)
                if r["Sales"] > 0
                else 0,
                axis=1,
            )

            fig = make_subplots(specs=[[{"secondary_y": True}]])

            fig.add_trace(
                go.Bar(
                    x=monthly_trend["Month"],
                    y=monthly_trend["Spends"],
                    name="Spends (₹)",
                    marker_color="#1F4E78",
                ),
                secondary_y=False,
            )

            fig.add_trace(
                go.Bar(
                    x=monthly_trend["Month"],
                    y=monthly_trend["Sales"],
                    name="Sales (₹)",
                    marker_color="#2CA02C",
                ),
                secondary_y=False,
            )

            fig.add_trace(
                go.Scatter(
                    x=monthly_trend["Month"],
                    y=monthly_trend["ROAS"],
                    name="ROAS",
                    mode="lines+markers",
                    line=dict(color="#FF7F0E", width=3),
                ),
                secondary_y=True,
            )

            fig.update_layout(
                title="Overall MoM Spends, Sales & ROAS Trend (Latest Month First)",
                barmode="group",
                hovermode="x unified",
                height=450,
            )
            fig.update_xaxes(title_text="Month")
            fig.update_yaxes(title_text="Amount (₹)", secondary_y=False)
            fig.update_yaxes(title_text="ROAS (x)", secondary_y=True)

            st.plotly_chart(fig, use_container_width=True)

            st.divider()

            st.markdown("### Performance Breakdown by Ad / Match Type")
            ad_type_summary = (
                df_filtered.groupby("Ad Type Combined")
                .agg(
                    Spends=("_budget_consumed", "sum"),
                    Sales=("_sales", "sum"),
                    Orders=("_orders", "sum"),
                )
                .reset_index()
            )

            ad_type_summary["ROAS"] = ad_type_summary.apply(
                lambda r: round(r["Sales"] / r["Spends"], 2) if r["Spends"] > 0 else 0,
                axis=1,
            )

            fig_ad = go.Figure()
            fig_ad.add_trace(
                go.Bar(
                    x=ad_type_summary["Ad Type Combined"],
                    y=ad_type_summary["Spends"],
                    name="Spends (₹)",
                    marker_color="#5B9BD5",
                )
            )
            fig_ad.add_trace(
                go.Bar(
                    x=ad_type_summary["Ad Type Combined"],
                    y=ad_type_summary["Sales"],
                    name="Sales (₹)",
                    marker_color="#70AD47",
                )
            )

            fig_ad.update_layout(
                title="Spends vs Sales by Match/Ad Type (Filtered View)",
                barmode="group",
                height=400,
            )
            st.plotly_chart(fig_ad, use_container_width=True)
else:
    st.info(
        "👆 Please upload at least one Blinkit ad report spreadsheet above to get"
        " started."
    )
