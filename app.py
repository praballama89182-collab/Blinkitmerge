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

# Custom CSS for center-aligning dataframe headers & cells
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

# Helper function to style downloadable Excel Pivot tables with custom styling, borders, gap row, and Grand Total
def style_and_export_pivot(pivot_df, sheet_name="Comparison"):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    
    # Styles definition
    top_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")  # Dark Steel Blue
    top_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    sec_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")  # Soft Ice Blue
    sec_header_font = Font(name="Calibri", size=10, bold=True, color="1F4E78")
    
    index_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")       # Light Slate Tint
    index_font = Font(name="Calibri", size=10, bold=True, color="000000")
    
    total_fill = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")       # Grand Total Highlight
    total_font = Font(name="Calibri", size=10, bold=True, color="000000")
    
    data_font = Font(name="Calibri", size=10, color="000000")
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )

    if isinstance(pivot_df.columns, pd.MultiIndex):
        months = pivot_df.columns.get_level_values(0)
        metrics = pivot_df.columns.get_level_values(1)
        entity_title = pivot_df.index.name if pivot_df.index.name else ""

        # Row 1: Months (Top Level Header)
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

        # Row 2: Metrics (Second Level Header)
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

        # Row 3: Gap row from Column B onwards (B3 onwards)
        cell_a3 = ws.cell(row=3, column=1, value="")
        cell_a3.border = thin_border
        for col_idx in range(2, len(metrics) + 2):
            gap_cell = ws.cell(row=3, column=col_idx, value="")
            gap_cell.border = thin_border

        # Row 4 onwards: Data Rows
        start_data_row = 4
        for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=start_data_row):
            is_grand_total = (str(idx_val).strip().lower() == 'grand total')
            
            # Row Index Label
            idx_cell = ws.cell(row=r_idx, column=1, value=str(idx_val))
            idx_cell.fill = total_fill if is_grand_total else index_fill
            idx_cell.font = total_font if is_grand_total else index_font
            idx_cell.alignment = align_left
            idx_cell.border = thin_border

            # Row Data Values
            for c_idx, val in enumerate(row_data, start=2):
                val_cell = ws.cell(row=r_idx, column=c_idx, value=val)
                val_cell.font = total_font if is_grand_total else data_font
                if is_grand_total:
                    val_cell.fill = total_fill
                val_cell.alignment = align_center
                val_cell.border = thin_border
                if isinstance(val, (int, float, np.number)):
                    val_cell.number_format = '#,##0.00' if isinstance(val, float) else '#,##0'
    else:
        for col_idx, col_name in enumerate(pivot_df.columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=str(col_name))
            cell.fill = top_header_fill
            cell.font = top_header_font
            cell.alignment = align_center
            cell.border = thin_border

        for r_idx, row_data in enumerate(pivot_df.values, start=2):
            for c_idx, val in enumerate(row_data, start=1):
                val_cell = ws.cell(row=r_idx, column=c_idx, value=val)
                val_cell.font = data_font
                val_cell.alignment = align_center
                val_cell.border = thin_border

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# Helper function to export ALL 4 MoM Pivot Tables into a single combined Excel Workbook
def convert_all_pivots_to_excel(pivot_dict):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    top_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    top_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    sec_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    sec_header_font = Font(name="Calibri", size=10, bold=True, color="1F4E78")
    index_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    index_font = Font(name="Calibri", size=10, bold=True, color="000000")
    total_fill = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
    total_font = Font(name="Calibri", size=10, bold=True, color="000000")
    data_font = Font(name="Calibri", size=10, color="000000")
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )

    for sheet_name, pivot_df in pivot_dict.items():
        if pivot_df is not None and not pivot_df.empty:
            ws = wb.create_sheet(title=sheet_name[:31])
            months = pivot_df.columns.get_level_values(0)
            metrics = pivot_df.columns.get_level_values(1)
            entity_title = pivot_df.index.name if pivot_df.index.name else ""

            # Row 1: Months
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

            # Row 2: Metrics
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

            # Row 3: Gap row from B3 onwards
            ws.cell(row=3, column=1, value="").border = thin_border
            for col_idx in range(2, len(metrics) + 2):
                ws.cell(row=3, column=col_idx, value="").border = thin_border

            # Data Rows
            for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=4):
                is_grand_total = (str(idx_val).strip().lower() == 'grand total')
                
                idx_cell = ws.cell(row=r_idx, column=1, value=str(idx_val))
                idx_cell.fill = total_fill if is_grand_total else index_fill
                idx_cell.font = total_font if is_grand_total else index_font
                idx_cell.alignment = align_left
                idx_cell.border = thin_border

                for c_idx, val in enumerate(row_data, start=2):
                    val_cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    val_cell.font = total_font if is_grand_total else data_font
                    if is_grand_total:
                        val_cell.fill = total_fill
                    val_cell.alignment = align_center
                    val_cell.border = thin_border
                    if isinstance(val, (int, float, np.number)):
                        val_cell.number_format = '#,##0.00' if isinstance(val, float) else '#,##0'

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# 5 Dedicated Upload Boxes
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
                
                # --- Dynamic Month Derivation ---
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

        # --- FALLBACK LOGIC ---
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

        # Helper numeric extractor
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
            final_df['Ad Type Combined'] = final_df[match_col].fillna("Other")
        else:
            final_df['Ad Type Combined'] = "Other"

        # --- WEEK BUCKET LOGIC ---
        date_col = None
        for col_candidate in ['Date', 'date', 'Day', 'DATE']:
            if col_candidate in final_df.columns:
                date_col = col_candidate
                break

        if date_col:
            final_df['_date_dt'] = pd.to_datetime(final_df[date_col], dayfirst=True, errors='coerce')
            def assign_week(row):
                dt = row['_date_dt']
                if pd.isna(dt): return np.nan
                day = dt.day
                if 1 <= day <= 7: return "Week 1"
                elif 8 <= day <= 14: return "Week 2"
                elif 15 <= day <= 21: return "Week 3"
                elif 22 <= day <= 28: return "Week 4"
                elif day >= 29: return "Week 5"
                return np.nan
            final_df['Week'] = final_df.apply(assign_week, axis=1)
        else:
            final_df['Week'] = np.nan

        # Helper function for calculating aggregate statistics
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

            grouped['CPM'] = grouped.apply(lambda r: round((r['SPENDS'] / r['IMPRESSIONS']) * 1000, 2) if r['IMPRESSIONS'] > 0 else 0.0, axis=1)
            grouped['ROAS'] = grouped.apply(lambda r: round(r['SALES'] / r['SPENDS'], 2) if r['SPENDS'] > 0 else 0.0, axis=1)
            grouped['ACOS'] = grouped.apply(lambda r: round((r['SPENDS'] / r['SALES']) * 100, 2) if r['SALES'] > 0 else 0.0, axis=1)

            display_name = group_col.upper()
            if group_col == 'Campaign Name': display_name = 'CAMPAIGN NAME'
            elif group_col == 'Ad Type Combined': display_name = 'MATCH / AD TYPE'

            grouped = grouped.rename(columns={group_col: display_name})
            col_order = [display_name, 'IMPRESSIONS', 'CPM', 'ATC', 'ORDERS', 'SPENDS', 'SALES', 'ROAS', 'ACOS']
            return grouped.reindex(columns=col_order)

        # Helper function for Month-on-Month Comparison matrix with GRAND TOTAL row
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

            grouped['CPM'] = grouped.apply(lambda r: round((r['Spends'] / r['Impressions']) * 1000, 2) if r['Impressions'] > 0 else 0.0, axis=1)
            grouped['ROAS'] = grouped.apply(lambda r: round(r['Sales'] / r['Spends'], 2) if r['Spends'] > 0 else 0.0, axis=1)
            grouped['ACOS'] = grouped.apply(lambda r: round((r['Spends'] / r['Sales']) * 100, 2) if r['Sales'] > 0 else 0.0, axis=1)

            pivot_df = grouped.pivot(
                index=entity_col,
                columns='Month',
                values=['Impressions', 'CPM', 'ATC', 'Orders', 'Spends', 'Sales', 'ROAS', 'ACOS']
            )

            metrics_order = ['Impressions', 'CPM', 'ATC', 'Orders', 'Spends', 'Sales', 'ROAS', 'ACOS']
            all_months = df_input['Month'].unique()
            
            pivot_df = pivot_df.reorder_levels([1, 0], axis=1)
            
            sorted_cols = pd.MultiIndex.from_product(
                [all_months, metrics_order],
                names=['Month', 'Metric']
            )
            
            pivot_df = pivot_df.reindex(columns=sorted_cols).fillna(0)

            # --- CALCULATE GRAND TOTAL ROW ---
            grand_total_series = {}
            for month in all_months:
                month_df = working_df[working_df['Month'] == month]
                total_imp = month_df['_impressions'].sum()
                total_atc = month_df['_atc'].sum()
                total_orders = month_df['_orders'].sum()
                total_spends = month_df['_budget_consumed'].sum()
                total_sales = month_df['_sales'].sum()

                total_cpm = round((total_spends / total_imp) * 1000, 2) if total_imp > 0 else 0.0
                total_roas = round(total_sales / total_spends, 2) if total_spends > 0 else 0.0
                total_acos = round((total_spends / total_sales) * 100, 2) if total_sales > 0 else 0.0

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

        # --- GLOBAL MONTH FILTER FOR DASHBOARD ---
        st.markdown("### 🔍 Global Dashboard Filters")
        available_months = ["All Months"] + sorted(list(final_df['Month'].dropna().unique()))
        selected_month = st.selectbox("Select Month Across Dashboard (Excluding Comparison Tables)", available_months)

        filtered_df = final_df.copy()
        if selected_month != "All Months":
            filtered_df = filtered_df[filtered_df['Month'] == selected_month]

        # Top level KPI cards
        total_impressions = filtered_df['_impressions'].sum()
        total_sales = filtered_df['_sales'].sum()
        total_orders = filtered_df['_orders'].sum()
        total_atc = filtered_df['_atc'].sum()
        total_budget = filtered_df['_budget_consumed'].sum()
        overall_roas = round((total_sales / total_budget), 2) if total_budget > 0 else 0.0

        st.markdown("### 📈 Overall Campaign Performance Dashboard")
        
        row1_col1, row1_col2, row1_col3 = st.columns(3)
        with row1_col1: st.metric("Total Impressions", f"{int(total_impressions):,}")
        with row1_col2: st.metric("Total Sales", f"₹{total_sales:,.2f}")
        with row1_col3: st.metric("Total Budget Consumed", f"₹{total_budget:,.2f}")

        row2_col1, row2_col2, row2_col3 = st.columns(3)
        with row2_col1: st.metric("Overall RoAS", f"{overall_roas:.2f}x")
        with row2_col2: st.metric("Total Orders", f"{int(total_orders):,}")
        with row2_col3: st.metric("Total Add To Cart", f"{int(total_atc):,}")

        st.divider()

        # --- MAIN NAVIGATION TABS ---
        st.markdown("### 📑 Navigation & Performance Breakdown")
        main_tab1, main_tab2, main_tab3, main_tab4, main_tab5, main_tab6, main_tab7, main_tab8 = st.tabs([
            "📄 Raw Files Preview",
            "📌 Consolidated Master",
            "📊 Comparison Tables (MoM)",
            "📈 Interactive Trend Analytics",
            "🎯 Campaign Performance", 
            "📢 Ad Type Performance",
            "🔎 Keyword Performance",
            "📅 Weekly Trend"
        ])

        # TAB 1: Raw Files Preview
        with main_tab1:
            st.caption("Inspect individual sheets tab-by-tab for each uploaded file.")
            selected_file_name = st.selectbox("Select Uploaded File to Preview:", list(raw_files_dict.keys()))
            if selected_file_name:
                sheets = raw_files_dict[selected_file_name]
                selected_sheet = st.selectbox("Select Sheet Tab:", list(sheets.keys()))
                if selected_sheet:
                    st.dataframe(sheets[selected_sheet].head(100), use_container_width=True)

        # TAB 2: Consolidated Master Dataset
        with main_tab2:
            st.caption("Preview the combined dataset across all uploaded files.")
            preview_clean_df = final_df.drop(columns=['_impressions', '_direct_atc', '_indirect_atc', '_atc', '_direct_orders', '_indirect_orders', '_orders', '_direct_sales', '_indirect_sales', '_sales', '_budget_consumed', '_date_dt', 'Ad Type Combined'], errors='ignore')
            st.dataframe(preview_clean_df.head(100), use_container_width=True)

        # TAB 3: Month-on-Month Comparison Tables
        with main_tab3:
            st.subheader("📊 Month-on-Month Comparison Tables")
            st.caption("View side-by-side MoM metrics for Campaign, Ad Type, Keywords, and Weeks. Download individual tables or all 4 tables in a single workbook.")

            # Compute Pivot Tables
            camp_pivot = create_mom_comparison_table(final_df, 'Campaign Name') if 'Campaign Name' in final_df.columns else None
            ad_pivot = create_mom_comparison_table(final_df, 'Ad Type Combined') if 'Ad Type Combined' in final_df.columns else None
            
            kw_col = None
            for c in ['Search Term', 'Keyword', 'Targeting Value']:
                if c in final_df.columns:
                    kw_col = c
                    break
            kw_pivot = create_mom_comparison_table(final_df, kw_col) if kw_col else None
            week_pivot = create_mom_comparison_table(final_df, 'Week') if ('Week' in final_df.columns and final_df['Week'].notna().any()) else None

            # Consolidated 4-in-1 Download Button at top of MoM tab
            mom_dict = {
                "Campaign_MoM": camp_pivot,
                "AdType_MoM": ad_pivot,
                "Keyword_MoM": kw_pivot,
                "Weekly_MoM": week_pivot
            }
            all_pivots_bytes = convert_all_pivots_to_excel(mom_dict)

            st.download_button(
                label="📥 Download All MoM Comparison Tables (.xlsx)",
                data=all_pivots_bytes,
                file_name="All_MoM_Comparison_Tables_Combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_all_mom_pivots",
                type="primary"
            )

            st.divider()

            comp_sub_tab1, comp_sub_tab2, comp_sub_tab3, comp_sub_tab4 = st.tabs([
                "🎯 Campaign Comparison",
                "📢 Ad Type Comparison",
                "🔎 Keyword Comparison",
                "📅 Weekly Comparison"
            ])

            with comp_sub_tab1:
                st.markdown("#### Campaign Month-on-Month Comparison Table")
                if camp_pivot is not None and not camp_pivot.empty:
                    st.dataframe(camp_pivot, use_container_width=True)
                    
                    excel_camp_pivot = style_and_export_pivot(camp_pivot, sheet_name="Campaign_MoM")
                    st.download_button(
                        label="📥 Download Campaign Comparison Table (.xlsx)",
                        data=excel_camp_pivot,
                        file_name="Campaign_MoM_Comparison_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_camp_pivot"
                    )
                else:
                    st.info("No Campaign Name column found.")

            with comp_sub_tab2:
                st.markdown("#### Ad Type Month-on-Month Comparison Table")
                if ad_pivot is not None and not ad_pivot.empty:
                    st.dataframe(ad_pivot, use_container_width=True)
                    
                    excel_ad_pivot = style_and_export_pivot(ad_pivot, sheet_name="AdType_MoM")
                    st.download_button(
                        label="📥 Download Ad Type Comparison Table (.xlsx)",
                        data=excel_ad_pivot,
                        file_name="Ad_Type_MoM_Comparison_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_ad_pivot"
                    )

            with comp_sub_tab3:
                st.markdown("#### Keyword / Search Term Month-on-Month Comparison Table")
                if kw_pivot is not None and not kw_pivot.empty:
                    st.dataframe(kw_pivot, use_container_width=True)
                    
                    excel_kw_pivot = style_and_export_pivot(kw_pivot, sheet_name="Keyword_MoM")
                    st.download_button(
                        label="📥 Download Keyword Comparison Table (.xlsx)",
                        data=excel_kw_pivot,
                        file_name="Keyword_MoM_Comparison_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_kw_pivot"
                    )
                else:
                    st.info("No Keyword or Search Term column found.")

            with comp_sub_tab4:
                st.markdown("#### Weekly Month-on-Month Comparison Table")
                if week_pivot is not None and not week_pivot.empty:
                    st.dataframe(week_pivot, use_container_width=True)
                    
                    excel_week_pivot = style_and_export_pivot(week_pivot, sheet_name="Weekly_MoM")
                    st.download_button(
                        label="📥 Download Weekly Comparison Table (.xlsx)",
                        data=excel_week_pivot,
                        file_name="Weekly_MoM_Comparison_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_week_pivot"
                    )

        # TAB 4: Interactive Trend Analytics
        with main_tab4:
            st.subheader("📈 Interactive Multi-Metric Trend Analytics")
            st.caption("Select multiple months and metrics to compare performance across time with smooth curved lines overlaying metric pillar columns.")

            all_df_months = sorted(list(final_df['Month'].dropna().unique()))
            selected_trend_months = st.multiselect(
                "Select Months to Include in Trend Analysis:",
                options=all_df_months,
                default=all_df_months
            )

            metric_map = {
                'Sales (₹)': '_sales',
                'Spends (₹)': '_budget_consumed',
                'ROAS': 'ROAS',
                'Orders': '_orders',
                'Add To Cart (ATC)': '_atc',
                'Impressions': '_impressions',
                'ACOS (%)': 'ACOS',
                'CPM (₹)': 'CPM'
            }

            selected_trend_metrics = st.multiselect(
                "Select Metrics to Display on Trend Graph:",
                options=list(metric_map.keys()),
                default=['Sales (₹)', 'Spends (₹)', 'ROAS']
            )

            if selected_trend_months and selected_trend_metrics:
                trend_df = final_df[final_df['Month'].isin(selected_trend_months)].copy()

                monthly_summary = trend_df.groupby('Month').agg(
                    _impressions=('_impressions', 'sum'),
                    _atc=('_atc', 'sum'),
                    _orders=('_orders', 'sum'),
                    _budget_consumed=('_budget_consumed', 'sum'),
                    _sales=('_sales', 'sum')
                ).reset_index()

                monthly_summary['CPM'] = monthly_summary.apply(lambda r: round((r['_budget_consumed'] / r['_impressions']) * 1000, 2) if r['_impressions'] > 0 else 0.0, axis=1)
                monthly_summary['ROAS'] = monthly_summary.apply(lambda r: round(r['_sales'] / r['_budget_consumed'], 2) if r['_budget_consumed'] > 0 else 0.0, axis=1)
                monthly_summary['ACOS'] = monthly_summary.apply(lambda r: round((r['_budget_consumed'] / r['_sales']) * 100, 2) if r['_sales'] > 0 else 0.0, axis=1)

                month_order = {m: i for i, m in enumerate(all_df_months)}
                monthly_summary['month_idx'] = monthly_summary['Month'].map(month_order)
                monthly_summary = monthly_summary.sort_values('month_idx').drop(columns=['month_idx'])

                # Render Data Table
                st.markdown("#### 📊 Selected Months Data Summary")
                disp_summary = monthly_summary.rename(columns={
                    'Month': 'MONTH',
                    '_impressions': 'IMPRESSIONS',
                    '_atc': 'ATC',
                    '_orders': 'ORDERS',
                    '_budget_consumed': 'SPENDS',
                    '_sales': 'SALES'
                })
                st.dataframe(disp_summary, use_container_width=True)

                # Render Plotly Curved Spline Line & Bar Combination Graph
                st.markdown("#### 📉 Curved Trend Line & Pillar Combination Graph")
                
                fig_trend = make_subplots(specs=[[{"secondary_y": True}]])
                palette = ['#0D47A1', '#1B5E20', '#B71C1C', '#E65100', '#4A148C', '#006064', '#F57F17']

                for idx, metric_label in enumerate(selected_trend_metrics):
                    col_key = metric_map[metric_label]
                    use_sec_y = metric_label in ['ROAS', 'ACOS (%)', 'CPM (₹)']
                    color = palette[idx % len(palette)]

                    if idx == 0:
                        fig_trend.add_trace(
                            go.Bar(
                                x=monthly_summary['Month'],
                                y=monthly_summary[col_key],
                                name=f"{metric_label} (Volume)",
                                marker=dict(
                                    color=color,
                                    opacity=0.7,
                                    line=dict(color='#000000', width=1)
                                ),
                                text=monthly_summary[col_key].apply(lambda v: f"{v:,.2f}" if isinstance(v, float) else f"{v:,}"),
                                textposition="auto"
                            ),
                            secondary_y=use_sec_y
                        )

                    fig_trend.add_trace(
                        go.Scatter(
                            x=monthly_summary['Month'],
                            y=monthly_summary[col_key],
                            name=metric_label,
                            mode='lines+markers+text',
                            line=dict(shape='spline', width=4, color=color),
                            marker=dict(size=9, color=color, symbol='circle'),
                            text=monthly_summary[col_key].apply(lambda v: f"{v:,.2f}" if isinstance(v, float) else f"{v:,}"),
                            textposition="top center"
                        ),
                        secondary_y=use_sec_y
                    )

                fig_trend.update_layout(
                    title="<b>Multi-Metric Trend Curve & Volume Pillar Analysis</b>",
                    template="plotly_white",
                    height=580,
                    hovermode="x unified",
                    barmode="group",
                    legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
                    xaxis=dict(title="Time Intervals (Months)", showgrid=True),
                    yaxis=dict(title="Numerical Values (Volume / Spends / Sales)", showgrid=True),
                    yaxis2=dict(title="Ratios (ROAS, ACOS %, CPM)", overlaying="y", side="right", showgrid=False)
                )

                st.plotly_chart(fig_trend, use_container_width=True)
            else:
                st.info("Please select at least one Month and one Metric to display the trend analysis.")

        # TAB 5: Campaign Performance
        with main_tab5:
            if 'Campaign Name' in filtered_df.columns:
                campaign_options = ["All"] + sorted([str(x) for x in filtered_df['Campaign Name'].dropna().unique()])
                selected_campaign = st.selectbox("Select Campaign:", campaign_options, key="camp_single_filt")
                campaign_df = compute_grouped_table(filtered_df, 'Campaign Name', selected_campaign)
                if not campaign_df.empty:
                    st.dataframe(campaign_df, use_container_width=True, hide_index=True)

        # TAB 6: Ad Type Performance
        with main_tab6:
            if 'Ad Type Combined' in filtered_df.columns:
                adtype_options = ["All"] + sorted([str(x) for x in filtered_df['Ad Type Combined'].dropna().unique()])
                selected_adtype = st.selectbox("Select Ad Type:", adtype_options, key="ad_single_filt")
                adtype_df = compute_grouped_table(filtered_df, 'Ad Type Combined', selected_adtype)
                if not adtype_df.empty:
                    st.dataframe(adtype_df, use_container_width=True, hide_index=True)

        # TAB 7: Keyword Performance
        with main_tab7:
            kw_col = None
            for c in ['Search Term', 'Keyword', 'Targeting Value']:
                if c in filtered_df.columns:
                    kw_col = c
                    break
            if kw_col:
                kw_options = ["All"] + sorted([str(x) for x in filtered_df[kw_col].dropna().unique()])
                selected_kw = st.selectbox(f"Select {kw_col}:", kw_options, key="kw_single_filt")
                search_df = compute_grouped_table(filtered_df, kw_col, selected_kw)
                if not search_df.empty:
                    st.dataframe(search_df, use_container_width=True, hide_index=True)

        # TAB 8: Weekly Performance Trend
        with main_tab8:
            if 'Week' in filtered_df.columns and filtered_df['Week'].notna().any():
                weekly_df = compute_grouped_table(filtered_df, 'Week', "All")
                st.dataframe(weekly_df, use_container_width=True, hide_index=True)

        st.divider()

        # Output Workbook Download
        st.subheader("💾 Download Consolidated Master Workbook")
        buffer_multi = io.BytesIO()
        with pd.ExcelWriter(buffer_multi, engine='openpyxl') as writer:
            for raw_tab_name, df_list in consolidated_raw_tabs.items():
                combined_raw_tab_df = pd.concat(df_list, ignore_index=True)
                clean_sheet_name = raw_tab_name[:31]
                combined_raw_tab_df.to_excel(writer, sheet_name=clean_sheet_name, index=False)
            
            master_export_df = final_df.drop(columns=['_impressions', '_direct_atc', '_indirect_atc', '_atc', '_direct_orders', '_indirect_orders', '_orders', '_direct_sales', '_indirect_sales', '_sales', '_budget_consumed', '_date_dt', 'Ad Type Combined'], errors='ignore')
            master_export_df.to_excel(writer, sheet_name='Consolidated_Master', index=False)

        buffer_multi.seek(0)
        st.download_button(
            label="📥 Download Complete Excel Workbook (Consolidated Raw Tabs + Final Master Tab)",
            data=buffer_multi,
            file_name="Blinkit_Consolidated_Master_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
