import os
import io
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(
    page_title="Blinkit Report Merger & Analytics",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for seamless alignment and clean table structure
st.markdown("""
<style>
    [data-testid="stDataFrame"] th, [data-testid="stDataFrame"] td {
        text-align: center !important;
        vertical-align: middle !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Blinkit Ad Report Merger & Analytics")
st.write("Upload up to 5 monthly Excel ad campaign spreadsheets (.xlsx, .xls, .xlsb, .xlsm). Preview raw & consolidated sheets, view Month-on-Month Comparison tables, and analyze multi-metric curve & bar trend comparisons.")

# Helper list for chronological month sorting
MONTH_ORDER = [
    "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
    "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"
]

def sort_months_chronologically(month_list):
    """Sorts a list of month names based on calendar order."""
    def get_month_index(m):
        m_str = str(m).strip().upper()
        if m_str in MONTH_ORDER:
            return MONTH_ORDER.index(m_str)
        return 99  # Fallback for non-standard month strings
    return sorted(list(set(month_list)), key=get_month_index)

# Helper function to style downloadable Excel Pivot tables cleanly
def style_and_export_pivot(pivot_df, sheet_name="Comparison"):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    
    top_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    top_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    sec_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    sec_header_font = Font(name="Calibri", size=10, bold=True, color="1F4E78")
    
    index_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    index_font = Font(name="Calibri", size=10, bold=True, color="000000")
    
    total_fill = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
    total_font = Font(name="Calibri", size=10, bold=True, color="000000")
    
    data_font = Font(name="Calibri", size=10, color="000000")
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)
    
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
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
                ws.merge_cells(start_row=1, start_column=current_col, end_row=1, end_column=current_col + num_sub - 1)
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
        for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=start_data_row):
            is_grand_total = (str(idx_val).strip().lower() == 'grand total')
            
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
                    if str(m_col).upper() == 'ACOS':
                        val_cell.value = round(val / 100.0, 4) if val > 1 else round(val, 4)
                        val_cell.number_format = '0.00%'
                    else:
                        val_cell.value = round(float(val), 2)
                        val_cell.number_format = '#,##0.00'
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
            is_pct_row = (str(row_data.iloc[0]).strip().lower() == 'percentage %')
            
            for c_idx, (col_name, val) in enumerate(zip(pivot_df.columns, row_data), start=1):
                val_cell = ws.cell(row=r_idx, column=c_idx)
                val_cell.font = total_font if is_pct_row else data_font
                val_cell.alignment = align_center
                val_cell.border = thin_border
                
                if str(col_name).upper() == 'ACOS' and not is_pct_row and isinstance(val, (int, float, np.number)):
                    val_cell.value = round(val / 100.0, 4) if val > 1 else round(val, 4)
                    val_cell.number_format = '0.00%'
                elif isinstance(val, (int, float, np.number)):
                    val_cell.value = round(float(val), 2)
                    val_cell.number_format = '#,##0.00'
                else:
                    val_cell.value = val

                if is_pct_row and c_idx > 1:
                    val_str = str(val).replace('%', '').replace('+', '').strip()
                    try:
                        num_v = float(val_str)
                        if num_v < 0:
                            val_cell.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                            val_cell.font = Font(name="Calibri", size=10, bold=True, color="721C24")
                        elif num_v > 0:
                            val_cell.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
                            val_cell.font = Font(name="Calibri", size=10, bold=True, color="155724")
                    except ValueError:
                        pass
                elif is_pct_row:
                    val_cell.fill = sec_header_fill

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

def convert_all_pivots_to_excel(pivot_dict):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    top_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    top_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    sec_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    sec_header_font = Font(name="Calibri", size=10, bold=True, color="1F4E78")
    index_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    index_font = Font(name="Calibri", size=10, bold=True, color="000000")
    total_fill = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
    total_font = Font(name="Calibri", size=10, bold=True, color="000000")
    data_font = Font(name="Calibri", size=10, color="000000")
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
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
                        ws.merge_cells(start_row=1, start_column=current_col, end_row=1, end_column=current_col + num_sub - 1)
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

                for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=3):
                    is_grand_total = (str(idx_val).strip().lower() == 'grand total')
                    
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
                            if str(m_col).upper() == 'ACOS':
                                val_cell.value = round(val / 100.0, 4) if val > 1 else round(val, 4)
                                val_cell.number_format = '0.00%'
                            else:
                                val_cell.value = round(float(val), 2)
                                val_cell.number_format = '#,##0.00'
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
                    is_pct_row = (str(row_data.iloc[0]).strip().lower() == 'percentage %')
                    for c_idx, (col_name, val) in enumerate(zip(pivot_df.columns, row_data), start=1):
                        val_cell = ws.cell(row=r_idx, column=c_idx)
                        val_cell.font = total_font if is_pct_row else data_font
                        val_cell.alignment = align_center
                        val_cell.border = thin_border
                        
                        if str(col_name).upper() == 'ACOS' and not is_pct_row and isinstance(val, (int, float, np.number)):
                            val_cell.value = round(val / 100.0, 4) if val > 1 else round(val, 4)
                            val_cell.number_format = '0.00%'
                        elif isinstance(val, (int, float, np.number)):
                            val_cell.value = round(float(val), 2)
                            val_cell.number_format = '#,##0.00'
                        else:
                            val_cell.value = val

                        if is_pct_row and c_idx > 1:
                            val_str = str(val).replace('%', '').replace('+', '').strip()
                            try:
                                num_v = float(val_str)
                                if num_v < 0:
                                    val_cell.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                                    val_cell.font = Font(name="Calibri", size=10, bold=True, color="721C24")
                                elif num_v > 0:
                                    val_cell.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
                                    val_cell.font = Font(name="Calibri", size=10, bold=True, color="155724")
                            except ValueError:
                                pass
                        elif is_pct_row:
                            val_cell.fill = sec_header_fill

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# File Upload Section
col1, col2, col3, col4, col5 = st.columns(5)
uploaded_files = []

with col1:
    f1 = st.file_uploader("Upload File 1", type=["xlsx", "xls", "xlsb", "xlsm"], key="file1")
    if f1: uploaded_files.append(f1)

with col2:
    f2 = st.file_uploader("Upload File 2", type=["xlsx", "xls", "xlsb", "xlsm"], key="file2")
    if f2: uploaded_files.append(f2)

with col3:
    f3 = st.file_uploader("Upload File 3", type=["xlsx", "xls", "xlsb", "xlsm"], key="file3")
    if f3: uploaded_files.append(f3)

with col4:
    f4 = st.file_uploader("Upload File 4", type=["xlsx", "xls", "xlsb", "xlsm"], key="file4")
    if f4: uploaded_files.append(f4)

with col5:
    f5 = st.file_uploader("Upload File 5", type=["xlsx", "xls", "xlsb", "xlsm"], key="file5")
    if f5: uploaded_files.append(f5)

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
                for col_candidate in ['Date', 'date', 'Day', 'DATE']:
                    if col_candidate in df.columns:
                        date_col = col_candidate
                        break

                if date_col:
                    dt_series = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
                    month_series = dt_series.dt.strftime('%B').str.upper()
                    df['Month'] = month_series.fillna(fallback_month_name)
                else:
                    df['Month'] = fallback_month_name

                df_raw = df.copy()
                if sheet_name not in consolidated_raw_tabs:
                    consolidated_raw_tabs[sheet_name] = []
                consolidated_raw_tabs[sheet_name].append(df_raw)
                
                df_consolidated = df.copy()
                if 'Month' in df_consolidated.columns:
                    col_month = df_consolidated.pop('Month')
                    df_consolidated.insert(0, 'Month', col_month)
                else:
                    df_consolidated.insert(0, 'Month', fallback_month_name)
                    
                df_consolidated.insert(1, 'Tab Name', sheet_name)
                consolidated_dfs.append(df_consolidated)
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")

    if consolidated_dfs:
        final_df = pd.concat(consolidated_dfs, ignore_index=True)
        
        base_cols = ['Month', 'Tab Name']
        product_listing_cols = []
        for df in consolidated_dfs:
            if 'PRODUCT_LISTING' in df['Tab Name'].values:
                pl_df = df[df['Tab Name'] == 'PRODUCT_LISTING']
                product_listing_cols = [c for c in pl_df.columns if c not in base_cols]
                break
        
        remaining_cols = [c for c in final_df.columns if c not in base_cols and c not in product_listing_cols]
        final_df = final_df.reindex(columns=base_cols + product_listing_cols + remaining_cols)

        match_col = 'Match Type' if 'Match Type' in final_df.columns else None
        target_col = 'Targeting Type' if 'Targeting Type' in final_df.columns else None
        tab_col = 'Tab Name' if 'Tab Name' in final_df.columns else None

        def is_na_series(series):
            if series is None or series.empty:
                return pd.Series(True, index=final_df.index)
            cleaned = series.astype(str).str.strip().str.upper()
            return series.isna() | cleaned.isin(['NA', 'N/A', 'NAN', 'NONE', ''])

        final_df['_filled_fallback'] = False

        if match_col:
            na_match = is_na_series(final_df[match_col])
            if target_col:
                valid_target = ~is_na_series(final_df[target_col])
                fill_from_target = na_match & valid_target
                final_df.loc[fill_from_target, match_col] = final_df.loc[fill_from_target, target_col]
                final_df.loc[fill_from_target, '_filled_fallback'] = True
                na_match = is_na_series(final_df[match_col])

            if tab_col:
                fill_from_tab = na_match
                final_df.loc[fill_from_tab, match_col] = final_df.loc[fill_from_tab, tab_col]
                final_df.loc[fill_from_tab, '_filled_fallback'] = True
        elif target_col:
            final_df['Match Type'] = final_df[target_col]
            na_match = is_na_series(final_df['Match Type'])
            if tab_col:
                final_df.loc[na_match, 'Match Type'] = final_df.loc[na_match, tab_col]
                final_df.loc[na_match, '_filled_fallback'] = True
            match_col = 'Match Type'

        def get_numeric_col(df, possible_cols):
            for col in possible_cols:
                if col in df.columns:
                    return pd.to_numeric(df[col], errors='coerce').fillna(0)
            return pd.Series(0, index=df.index)

        final_df['_impressions'] = get_numeric_col(final_df, ['Impressions'])
        final_df['_direct_atc'] = get_numeric_col(final_df, ['Direct ATC'])
        final_df['_indirect_atc'] = get_numeric_col(final_df, ['Indirect ATC'])
        final_df['_atc'] = final_df['_direct_atc'] + final_df['_indirect_atc']

        final_df['_direct_orders'] = get_numeric_col(final_df, ['Direct Quantities Sold', 'Direct Orders'])
        final_df['_indirect_orders'] = get_numeric_col(final_df, ['Indirect Quantities Sold', 'Indirect Orders'])
        final_df['_orders'] = final_df['_direct_orders'] + final_df['_indirect_orders']

        final_df['_direct_sales'] = get_numeric_col(final_df, ['Direct Sales'])
        final_df['_indirect_sales'] = get_numeric_col(final_df, ['Indirect Sales'])
        final_df['_sales'] = final_df['_direct_sales'] + final_df['_indirect_sales']

        final_df['_budget_consumed'] = get_numeric_col(final_df, ['Estimated Budget Consumed', 'Budget Consumed', 'Spend'])

        if match_col and match_col in final_df.columns:
            final_df['Ad Type Combined'] = final_df[match_col].fillna("Other").astype(str).str.title()
        else:
            final_df['Ad Type Combined'] = "Other"

        date_col = None
        for col_candidate in ['Date', 'date', 'Day', 'DATE']:
            if col_candidate in final_df.columns:
                date_col = col_candidate
                break

        if date_col:
            final_df['_date_dt'] = pd.to_datetime(final_df[date_col], dayfirst=True, errors='coerce')
            def assign_week_formatted(row):
                dt = row['_date_dt']
                if pd.isna(dt): return np.nan
                day = dt.day
                if 1 <= day <= 7: return "Week 1 (1-7)"
                elif 8 <= day <= 14: return "Week 2 (8-14)"
                elif 15 <= day <= 21: return "Week 3 (15-21)"
                elif 22 <= day <= 28: return "Week 4 (22-28)"
                elif day >= 29: return "Week 5 (29-31)"
                return np.nan
            final_df['Week'] = final_df.apply(assign_week_formatted, axis=1)
        else:
            final_df['Week'] = np.nan

        def compute_grouped_table(df_subset, group_col, selected_item="All"):
            if group_col not in df_subset.columns:
                return pd.DataFrame()
            
            df_working = df_subset.dropna(subset=[group_col]).copy()
            df_working[group_col] = df_working[group_col].astype(str)
            
            if selected_item and selected_item != "All":
                df_working = df_working[df_working[group_col] == selected_item]
            
            if df_working.empty:
                return pd.DataFrame()

            grouped = df_working.groupby(group_col).agg(
                IMPRESSIONS=('_impressions', 'sum'),
                ATC=('_atc', 'sum'),
                ORDERS=('_orders', 'sum'),
                SPENDS=('_budget_consumed', 'sum'),
                SALES=('_sales', 'sum')
            ).reset_index()

            grouped['CPM'] = grouped.apply(lambda r: round((r['SPENDS'] / r['IMPRESSIONS']) * 1000, 2) if r['IMPRESSIONS'] > 0 else 0.00, axis=1)
            grouped['ROAS'] = grouped.apply(lambda r: round(r['SALES'] / r['SPENDS'], 2) if r['SPENDS'] > 0 else 0.00, axis=1)
            grouped['ACOS'] = grouped.apply(lambda r: round((r['SPENDS'] / r['SALES']) * 100, 2) if r['SALES'] > 0 else 0.00, axis=1)

            display_name = group_col.upper()
            if group_col == 'Campaign Name': display_name = 'CAMPAIGN NAME'
            elif group_col == 'Ad Type Combined': display_name = 'MATCH / AD TYPE'

            grouped = grouped.rename(columns={group_col: display_name})
            col_order = [display_name, 'IMPRESSIONS', 'CPM', 'ATC', 'ORDERS', 'SPENDS', 'SALES', 'ROAS', 'ACOS']
            
            res_df = grouped.reindex(columns=col_order)
            res_df['IMPRESSIONS'] = res_df['IMPRESSIONS'].round(2)
            res_df['ATC'] = res_df['ATC'].round(2)
            res_df['ORDERS'] = res_df['ORDERS'].round(2)
            res_df['SPENDS'] = res_df['SPENDS'].round(2)
            res_df['SALES'] = res_df['SALES'].round(2)
            return res_df

        def create_mom_comparison_table(df_input, entity_col):
            if entity_col not in df_input.columns:
                return pd.DataFrame()
            
            working_df = df_input.dropna(subset=[entity_col, 'Month']).copy()
            if working_df.empty:
                return pd.DataFrame()
            
            grouped = working_df.groupby([entity_col, 'Month']).agg(
                Impressions=('_impressions', 'sum'),
                ATC=('_atc', 'sum'),
                Orders=('_orders', 'sum'),
                Spends=('_budget_consumed', 'sum'),
                Sales=('_sales', 'sum')
            ).reset_index()

            grouped['CPM'] = grouped.apply(lambda r: round((r['Spends'] / r['Impressions']) * 1000, 2) if r['Impressions'] > 0 else 0.00, axis=1)
            grouped['ROAS'] = grouped.apply(lambda r: round(r['Sales'] / r['Spends'], 2) if r['Spends'] > 0 else 0.00, axis=1)
            grouped['ACOS'] = grouped.apply(lambda r: round((r['Spends'] / r['Sales']) * 100, 2) if r['Sales'] > 0 else 0.00, axis=1)

            pivot_df = grouped.pivot(
                index=entity_col,
                columns='Month',
                values=['Impressions', 'CPM', 'ATC', 'Orders', 'Spends', 'Sales', 'ROAS', 'ACOS']
            )

            metrics_order = ['Impressions', 'CPM', 'ATC', 'Orders', 'Spends', 'Sales', 'ROAS', 'ACOS']
            all_months = sort_months_chronologically(df_input['Month'].unique())
            
            pivot_df = pivot_df.reorder_levels([1, 0], axis=1)
            sorted_cols = pd.MultiIndex.from_product([all_months, metrics_order], names=['Month', 'Metric'])
            pivot_df = pivot_df.reindex(columns=sorted_cols).fillna(0.00)

            grand_total_series = {}
            for month in all_months:
                month_df = working_df[working_df['Month'] == month]
                total_imp = round(month_df['_impressions'].sum(), 2)
                total_atc = round(month_df['_atc'].sum(), 2)
                total_orders = round(month_df['_orders'].sum(), 2)
                total_spends = round(month_df['_budget_consumed'].sum(), 2)
                total_sales = round(month_df['_sales'].sum(), 2)

                total_cpm = round((total_spends / total_imp) * 1000, 2) if total_imp > 0 else 0.00
                total_roas = round(total_sales / total_spends, 2) if total_spends > 0 else 0.00
                total_acos = round((total_spends / total_sales) * 100, 2) if total_sales > 0 else 0.00

                grand_total_series[(month, 'Impressions')] = total_imp
                grand_total_series[(month, 'CPM')] = total_cpm
                grand_total_series[(month, 'ATC')] = total_atc
                grand_total_series[(month, 'Orders')] = total_orders
                grand_total_series[(month, 'Spends')] = total_spends
                grand_total_series[(month, 'Sales')] = total_sales
                grand_total_series[(month, 'ROAS')] = total_roas
                grand_total_series[(month, 'ACOS')] = total_acos

            pivot_df.loc['Grand Total'] = grand_total_series
            pivot_df.index.name = entity_col
            return pivot_df

        def create_monthly_summary_table(df_input):
            working_df = df_input.dropna(subset=['Month']).copy()
            if working_df.empty:
                return pd.DataFrame()

            monthly_agg = working_df.groupby('Month').agg(
                Impressions=('_impressions', 'sum'),
                ATC=('_atc', 'sum'),
                Orders=('_orders', 'sum'),
                Spends=('_budget_consumed', 'sum'),
                Sales=('_sales', 'sum')
            ).reset_index()

            monthly_agg['CPM'] = monthly_agg.apply(lambda r: round((r['Spends'] / r['Impressions']) * 1000, 2) if r['Impressions'] > 0 else 0.00, axis=1)
            monthly_agg['ROAS'] = monthly_agg.apply(lambda r: round(r['Sales'] / r['Spends'], 2) if r['Spends'] > 0 else 0.00, axis=1)
            monthly_agg['ACOS'] = monthly_agg.apply(lambda r: round((r['Spends'] / r['Sales']) * 100, 2) if r['Sales'] > 0 else 0.00, axis=1)

            monthly_agg['Impressions'] = monthly_agg['Impressions'].round(2)
            monthly_agg['ATC'] = monthly_agg['ATC'].round(2)
            monthly_agg['Orders'] = monthly_agg['Orders'].round(2)
            monthly_agg['Spends'] = monthly_agg['Spends'].round(2)
            monthly_agg['Sales'] = monthly_agg['Sales'].round(2)

            all_months = sort_months_chronologically(df_input['Month'].unique())
            monthly_agg['month_order'] = monthly_agg['Month'].map(lambda x: all_months.index(x) if x in all_months else 99)
            monthly_agg = monthly_agg.sort_values('month_order').drop(columns=['month_order'])

            col_order = ['Month', 'Impressions', 'CPM', 'ATC', 'Orders', 'Spends', 'Sales', 'ROAS', 'ACOS']
            monthly_agg = monthly_agg[col_order]

            if len(monthly_agg) >= 2:
                prev_row = monthly_agg.iloc[-2]
                curr_row = monthly_agg.iloc[-1]

                pct_row = {'Month': 'Percentage %'}
                for metric in ['Impressions', 'CPM', 'ATC', 'Orders', 'Spends', 'Sales', 'ROAS', 'ACOS']:
                    prev_val = float(str(prev_row[metric]).replace('%', '').strip()) if isinstance(prev_row[metric], str) else prev_row[metric]
                    curr_val = float(str(curr_row[metric]).replace('%', '').strip()) if isinstance(curr_row[metric], str) else curr_row[metric]
                    if prev_val > 0:
                        pct_change = round(((curr_val - prev_val) / prev_val) * 100, 2)
                        pct_row[metric] = f"{pct_change:.2f}%" if pct_change <= 0 else f"+{pct_change:.2f}%"
                    else:
                        pct_row[metric] = "0.00%"
                
                monthly_agg = pd.concat([monthly_agg, pd.DataFrame([pct_row])], ignore_index=True)

            monthly_agg['ACOS'] = monthly_agg['ACOS'].apply(lambda v: f"{v:.2f}%" if isinstance(v, (int, float)) else str(v))
            return monthly_agg

        def render_unified_single_table(df_to_show, key_prefix="mom", expandable_col=None):
            if df_to_show is None or df_to_show.empty:
                return

            working_df = df_to_show.copy()

            bottom_rows = None
            if isinstance(working_df.index, pd.Index) and 'Grand Total' in working_df.index:
                main_df = working_df.drop('Grand Total')
                bottom_rows = working_df.loc[['Grand Total']]
            elif 'Month' in working_df.columns and 'Percentage %' in working_df['Month'].values:
                main_df = working_df[working_df['Month'] != 'Percentage %'].copy()
                bottom_rows = working_df[working_df['Month'] == 'Percentage %'].copy()
            else:
                main_df = working_df.copy()

            # Column-based sorting controls
            col_sort, col_order_dir = st.columns([3, 1])
            with col_sort:
                if isinstance(main_df.columns, pd.MultiIndex):
                    sort_options = ["Default Order"] + [f"{month} - {metric}" for month, metric in main_df.columns]
                else:
                    sort_options = ["Default Order"] + list(main_df.columns)
                selected_sort_col = st.selectbox("Sort Table Column:", sort_options, key=f"{key_prefix}_sort_col")

            with col_order_dir:
                sort_direction = st.radio("Order:", ["Ascending", "Descending"], horizontal=True, key=f"{key_prefix}_sort_dir")

            if selected_sort_col != "Default Order":
                ascending_flag = (sort_direction == "Ascending")
                if isinstance(main_df.columns, pd.MultiIndex):
                    parts = selected_sort_col.split(" - ")
                    sort_key = (parts[0], parts[1])
                    main_df = main_df.sort_values(by=sort_key, ascending=ascending_flag)
                else:
                    main_df = main_df.sort_values(by=selected_sort_col, ascending=ascending_flag)

            if bottom_rows is not None:
                final_display_df = pd.concat([main_df, bottom_rows])
            else:
                final_display_df = main_df

            # Explicit column formatting (2 decimal places for all numeric metrics)
            if isinstance(final_display_df.columns, pd.MultiIndex):
                format_dict = {}
                for col in final_display_df.columns:
                    metric_name = str(col[1]).upper()
                    if metric_name in ['IMPRESSIONS', 'ATC', 'ORDERS']:
                        format_dict[col] = "{:,.2f}"
                    elif metric_name == 'CPM':
                        format_dict[col] = "{:,.2f}"
                    elif metric_name in ['SPENDS', 'SALES']:
                        format_dict[col] = "₹{:,.2f}"
                    elif metric_name == 'ROAS':
                        format_dict[col] = "{:,.2f}"
                    elif metric_name == 'ACOS':
                        format_dict[col] = "{:,.2f}%"
                
                styled_df = final_display_df.style.format(format_dict)
                st.dataframe(styled_df, use_container_width=True)
            else:
                format_dict = {}
                for col in final_display_df.columns:
                    col_upper = str(col).upper()
                    if col_upper in ['IMPRESSIONS', 'ATC', 'ORDERS', 'CPM', 'ROAS']:
                        format_dict[col] = "{:,.2f}"
                    elif col_upper in ['SPENDS', 'SALES']:
                        format_dict[col] = "₹{:,.2f}"
                styled_df = final_display_df.style.format(format_dict, na_rep="-")
                st.dataframe(styled_df, use_container_width=True)
