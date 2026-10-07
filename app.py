import os
import io
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(
    page_title="Blinkit Report Merger & Analytics",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for center-aligning dataframe headers & cells
st.markdown("""
<style>
    /* Center align headers and cells in Streamlit dataframes */
    [data-testid="stDataFrame"] th, [data-testid="stDataFrame"] td {
        text-align: center !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Blinkit Ad Report Merger & Analytics")
st.write("Upload up to 5 monthly Excel ad campaign spreadsheets (.xlsx, .xls, .xlsb, .xlsm). Preview raw & consolidated sheets, inspect campaign performance, and view weekly trend lines.")

# Helper function to convert dataframe to downloadable Excel bytes with custom formatting/highlighting
def convert_df_to_excel(df, sheet_name="Performance", highlight_col=None):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Highlight filled rows in Excel if requested
        if highlight_col and '_filled_fallback' in df.columns:
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            from openpyxl.styles import PatternFill
            yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            
            # Find column index for highlight_col
            col_idx = None
            for idx, col in enumerate(df.columns, start=1):
                if col == highlight_col:
                    col_idx = idx
                    break
            
            if col_idx:
                for row_idx, filled in enumerate(df['_filled_fallback'], start=2):  # start=2 for header
                    if filled:
                        worksheet.cell(row=row_idx, column=col_idx).fill = yellow_fill

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
    raw_files_dict = {}         # {filename: {sheet_name: df}}
    consolidated_raw_tabs = {}  # {sheet_name: [df1, df2, ...]}

    for uploaded_file in uploaded_files:
        fallback_month_name = os.path.splitext(uploaded_file.name)[0].upper()
        raw_files_dict[uploaded_file.name] = {}
        
        try:
            xls = pd.ExcelFile(uploaded_file)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                
                # Store raw sheet preview
                raw_files_dict[uploaded_file.name][sheet_name] = df.copy()
                
                # --- Dynamic Month Derivation from Date Column ---
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

                # Raw tab copy for consolidated output workbook
                df_raw = df.copy()
                
                if sheet_name not in consolidated_raw_tabs:
                    consolidated_raw_tabs[sheet_name] = []
                consolidated_raw_tabs[sheet_name].append(df_raw)
                
                # Consolidated Master copy
                df_consolidated = df.copy()
                
                # Reposition Month and Tab Name at the beginning
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
        
        # Determine column order based on PRODUCT_LISTING if available
        base_cols = ['Month', 'Tab Name']
        product_listing_cols = []
        for df in consolidated_dfs:
            if 'PRODUCT_LISTING' in df['Tab Name'].values:
                pl_df = df[df['Tab Name'] == 'PRODUCT_LISTING']
                product_listing_cols = [c for c in pl_df.columns if c not in base_cols]
                break
        
        remaining_cols = [c for c in final_df.columns if c not in base_cols and c not in product_listing_cols]
        final_df = final_df.reindex(columns=base_cols + product_listing_cols + remaining_cols)

        # --- REVISED FALLBACK LOGIC ---
        # Match Type -> Fallback to Targeting Type -> Fallback to Tab Name (Col D)
        match_col = 'Match Type' if 'Match Type' in final_df.columns else None
        target_col = 'Targeting Type' if 'Targeting Type' in final_df.columns else None
        tab_col = 'Tab Name' if 'Tab Name' in final_df.columns else None

        # Helper to detect NA / Empty strings
        def is_na_series(series):
            if series is None or series.empty:
                return pd.Series(True, index=final_df.index)
            cleaned = series.astype(str).str.strip().str.upper()
            return series.isna() | cleaned.isin(['NA', 'N/A', 'NAN', 'NONE', ''])

        final_df['_filled_fallback'] = False

        if match_col:
            na_match = is_na_series(final_df[match_col])

            # 1. Fill NA in Match Type from Targeting Type if available
            if target_col:
                valid_target = ~is_na_series(final_df[target_col])
                fill_from_target = na_match & valid_target
                final_df.loc[fill_from_target, match_col] = final_df.loc[fill_from_target, target_col]
                final_df.loc[fill_from_target, '_filled_fallback'] = True
                
                # Update NA mask for remaining un-filled rows
                na_match = is_na_series(final_df[match_col])

            # 2. Fill remaining NA in Match Type from Tab Name (Column D)
            if tab_col:
                fill_from_tab = na_match
                final_df.loc[fill_from_tab, match_col] = final_df.loc[fill_from_tab, tab_col]
                final_df.loc[fill_from_tab, '_filled_fallback'] = True
        elif target_col:
            # If Match Type column doesn't exist, create it from Targeting Type / Tab Name
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

        # Unified Ad Type column for analytics grouping
        if match_col and match_col in final_df.columns:
            final_df['Ad Type Combined'] = final_df[match_col].fillna("Other")
        else:
            final_df['Ad Type Combined'] = "Other"

        # --- WEEK BUCKET LOGIC (Parse DD-MM-YYYY format) ---
        date_col = None
        for col_candidate in ['Date', 'date', 'Day', 'DATE']:
            if col_candidate in final_df.columns:
                date_col = col_candidate
                break

        if date_col:
            final_df['_date_dt'] = pd.to_datetime(final_df[date_col], dayfirst=True, errors='coerce')
            
            def assign_week(row):
                dt = row['_date_dt']
                if pd.isna(dt):
                    return np.nan
                day = dt.day
                if 1 <= day <= 7:
                    return "Week 1"
                elif 8 <= day <= 14:
                    return "Week 2"
                elif 15 <= day <= 21:
                    return "Week 3"
                elif 22 <= day <= 28:
                    return "Week 4"
                elif day >= 29:
                    return "Week 5"
                return np.nan

            final_df['Week'] = final_df.apply(assign_week, axis=1)
        else:
            final_df['Week'] = np.nan

        # --- GLOBAL MONTH FILTER ---
        st.markdown("### 🔍 Global Dashboard Filters")
        available_months = ["All Months"] + sorted(list(final_df['Month'].dropna().unique()))
        selected_month = st.selectbox("Select Month Across Dashboard", available_months)

        # Apply Global Month Filter
        filtered_df = final_df.copy()
        if selected_month != "All Months":
            filtered_df = filtered_df[filtered_df['Month'] == selected_month]

        # --- TOP LEVEL DASHBOARD METRICS ---
        total_impressions = filtered_df['_impressions'].sum()
        total_sales = filtered_df['_sales'].sum()
        total_orders = filtered_df['_orders'].sum()
        total_atc = filtered_df['_atc'].sum()
        total_budget = filtered_df['_budget_consumed'].sum()
        
        overall_roas = round((total_sales / total_budget), 2) if total_budget > 0 else 0.0

        st.markdown("### 📈 Overall Campaign Performance Dashboard")
        
        row1_col1, row1_col2, row1_col3 = st.columns(3)
        with row1_col1:
            st.metric("Total Impressions", f"{int(total_impressions):,}")
        with row1_col2:
            st.metric("Total Sales", f"₹{total_sales:,.2f}")
        with row1_col3:
            st.metric("Total Budget Consumed", f"₹{total_budget:,.2f}")

        row2_col1, row2_col2, row2_col3 = st.columns(3)
        with row2_col1:
            st.metric("Overall RoAS", f"{overall_roas:.2f}x")
        with row2_col2:
            st.metric("Total Orders", f"{int(total_orders):,}")
        with row2_col3:
            st.metric("Total Add To Cart", f"{int(total_atc):,}")

        st.divider()

        # Helper function for grouping metrics
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

            # CPM = (Spends / Impressions) * 1000
            grouped['CPM'] = grouped.apply(
                lambda r: round((r['SPENDS'] / r['IMPRESSIONS']) * 1000, 2) if r['IMPRESSIONS'] > 0 else 0.0, axis=1
            )
            
            # ROAS = Sales / Spends
            grouped['ROAS'] = grouped.apply(
                lambda r: round(r['SALES'] / r['SPENDS'], 2) if r['SPENDS'] > 0 else 0.0, axis=1
            )

            # ACOS = (Spends / Sales) * 100
            grouped['ACOS'] = grouped.apply(
                lambda r: round((r['SPENDS'] / r['SALES']) * 100, 2) if r['SALES'] > 0 else 0.0, axis=1
            )

            display_name = group_col.upper()
            if group_col == 'Campaign Name':
                display_name = 'CAMPAIGN NAME'
            elif group_col == 'Ad Type Combined':
                display_name = 'MATCH / AD TYPE'

            grouped = grouped.rename(columns={group_col: display_name})

            # Reorder columns explicitly: [Entity, IMPRESSIONS, CPM, ATC, ORDERS, SPENDS, SALES, ROAS, ACOS]
            col_order = [display_name, 'IMPRESSIONS', 'CPM', 'ATC', 'ORDERS', 'SPENDS', 'SALES', 'ROAS', 'ACOS']
            grouped = grouped.reindex(columns=col_order)

            return grouped

        def style_roas(val):
            try:
                val_float = float(val)
                if val_float < 1.0:
                    return 'background-color: #ffcdd2; color: #b71c1c; font-weight: bold; text-align: center;'
                else:
                    return 'background-color: #c8e6c9; color: #1b5e20; font-weight: bold; text-align: center;'
            except:
                return ''

        def style_dataframe(df):
            styler = df.style
            if 'ROAS' in df.columns:
                if hasattr(styler, 'map'):
                    styler = styler.map(style_roas, subset=['ROAS'])
                else:
                    styler = styler.applymap(style_roas, subset=['ROAS'])
            
            styler = styler.set_properties(**{'text-align': 'center'})
            
            format_dict = {
                'SALES': '₹{:,.2f}', 
                'SPENDS': '₹{:,.2f}', 
                'CPM': '₹{:,.2f}',
                'IMPRESSIONS': '{:,.0f}', 
                'ORDERS': '{:,.0f}', 
                'ATC': '{:,.0f}',
                'ROAS': '{:.2f}x',
                'ACOS': '{:.2f}%'
            }
            active_formats = {k: v for k, v in format_dict.items() if k in df.columns}
            
            return styler.format(active_formats)

        # Highlight function for fallback-filled Match Type cells
        def highlight_filled_cells(df):
            styler = df.style
            if '_filled_fallback' in df.columns and match_col and match_col in df.columns:
                def highlight_match_col(row):
                    styles = [''] * len(row)
                    if row.get('_filled_fallback', False):
                        col_idx = row.index.get_loc(match_col)
                        styles[col_idx] = 'background-color: #FFF2CC; font-weight: bold; color: #856404;'
                    return styles

                styler = styler.apply(highlight_match_col, axis=1)
            return styler

        # --- MAIN NAVIGATION TABS ---
        st.markdown("### 📑 Navigation & Performance Breakdown")
        main_tab1, main_tab2, main_tab3, main_tab4, main_tab5, main_tab6 = st.tabs([
            "📄 Raw Files Preview",
            "📌 Consolidated Master Preview",
            "🎯 Campaign Performance", 
            "📢 Ad Type Performance",
            "🔎 Search Term / Keyword Performance",
            "📅 Weekly Performance Trend"
        ])

        # TAB 1: Raw Files Preview
        with main_tab1:
            st.caption("Inspect individual sheets tab-by-tab for each uploaded file.")
            selected_file_name = st.selectbox("Select Uploaded File to Preview:", list(raw_files_dict.keys()))
            if selected_file_name:
                sheets = raw_files_dict[selected_file_name]
                selected_sheet = st.selectbox("Select Sheet Tab:", list(sheets.keys()))
                if selected_sheet:
                    st.write(f"Showing raw data preview for **{selected_file_name}** ➔ **{selected_sheet}** ({len(sheets[selected_sheet])} rows):")
                    st.dataframe(sheets[selected_sheet].head(100), use_container_width=True)

        # TAB 2: Consolidated Master Dataset Preview
        with main_tab2:
            st.caption("Preview the combined dataset across all uploaded files. Yellow highlighted cells in 'Match Type' indicate values filled via fallback logic (Targeting Type ➔ Tab Name).")
            preview_clean_df = final_df.drop(columns=['_impressions', '_direct_atc', '_indirect_atc', '_atc', '_direct_orders', '_indirect_orders', '_orders', '_direct_sales', '_indirect_sales', '_sales', '_budget_consumed', '_date_dt', 'Ad Type Combined'], errors='ignore')
            st.write(f"Total Rows Consolidated: **{len(preview_clean_df):,}**")
            
            st.dataframe(highlight_filled_cells(preview_clean_df.head(100)), use_container_width=True)

        # TAB 3: Campaign Performance
        with main_tab3:
            st.caption("Aggregated performance metrics per campaign.")
            if 'Campaign Name' in filtered_df.columns:
                campaign_options = ["All"] + sorted([str(x) for x in filtered_df['Campaign Name'].dropna().unique()])
                selected_campaign = st.selectbox("Select or Search Campaign:", campaign_options, key="campaign_filter")
                
                campaign_df = compute_grouped_table(filtered_df, 'Campaign Name', selected_campaign)
                if not campaign_df.empty:
                    st.dataframe(style_dataframe(campaign_df), use_container_width=True, hide_index=True)
                    
                    excel_campaign = convert_df_to_excel(campaign_df, sheet_name="Campaign_Performance")
                    st.download_button(
                        label="📥 Download Campaign Performance Excel (.xlsx)",
                        data=excel_campaign,
                        file_name="Campaign_Performance_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_campaign"
                    )
                else:
                    st.info("No campaign data matching the selected criteria.")
            else:
                st.info("No 'Campaign Name' column found in dataset.")

        # TAB 4: Ad Type Performance
        with main_tab4:
            st.caption("Aggregated performance & share analysis across Ad Types & Match Types.")
            if 'Ad Type Combined' in filtered_df.columns:
                adtype_options = ["All"] + sorted([str(x) for x in filtered_df['Ad Type Combined'].dropna().unique()])
                selected_adtype = st.selectbox("Select or Search Ad Type / Targeting Type:", adtype_options, key="adtype_filter")
                
                adtype_df = compute_grouped_table(filtered_df, 'Ad Type Combined', selected_adtype)
                if not adtype_df.empty:
                    st.dataframe(style_dataframe(adtype_df), use_container_width=True, hide_index=True)
                    
                    st.markdown("#### 🥧 Ad Type Share Breakdown")
                    blue_palette = ['#03045E', '#0077B6', '#0096C7', '#00B4D8', '#48CAE4', '#90E0EF', '#ADE8F4', '#CAF0F8']
                    
                    pie_fig = make_subplots(
                        rows=1, cols=2,
                        specs=[[{"type": "domain"}, {"type": "domain"}]],
                        subplot_titles=["<b>Spends Share by Ad Type</b>", "<b>Sales Share by Ad Type</b>"]
                    )

                    pie_fig.add_trace(
                        go.Pie(
                            labels=adtype_df['MATCH / AD TYPE'],
                            values=adtype_df['SPENDS'],
                            name="Spends Share",
                            hole=0.4,
                            marker=dict(colors=blue_palette, line=dict(color='#FFFFFF', width=2)),
                            textinfo="label+value+percent",
                            texttemplate="%{label}<br>₹%{value:,.2f}<br>(%{percent})",
                            hovertemplate="<b>%{label}</b><br>Spends: ₹%{value:,.2f}<br>Share: %{percent}<extra></extra>"
                        ),
                        row=1, col=1
                    )

                    pie_fig.add_trace(
                        go.Pie(
                            labels=adtype_df['MATCH / AD TYPE'],
                            values=adtype_df['SALES'],
                            name="Sales Share",
                            hole=0.4,
                            marker=dict(colors=blue_palette, line=dict(color='#FFFFFF', width=2)),
                            textinfo="label+value+percent",
                            texttemplate="%{label}<br>₹%{value:,.2f}<br>(%{percent})",
                            hovertemplate="<b>%{label}</b><br>Sales: ₹%{value:,.2f}<br>Share: %{percent}<extra></extra>"
                        ),
                        row=1, col=2
                    )

                    pie_fig.update_layout(
                        height=520,
                        template="plotly_white",
                        showlegend=True,
                        legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5),
                        margin=dict(l=20, r=20, t=50, b=60)
                    )

                    st.plotly_chart(pie_fig, use_container_width=True)

                    excel_adtype = convert_df_to_excel(adtype_df, sheet_name="Ad_Type_Performance")
                    st.download_button(
                        label="📥 Download Ad Type Performance Excel (.xlsx)",
                        data=excel_adtype,
                        file_name="Ad_Type_Performance_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_adtype"
                    )
                else:
                    st.info("No Ad Type data matching the selected criteria.")
            else:
                st.info("No Ad Type data found.")

        # TAB 5: Keyword / Search Term Performance
        with main_tab5:
            st.caption("Performance across search terms / target keywords.")
            kw_col = None
            for col_candidate in ['Search Term', 'Keyword', 'Targeting Value']:
                if col_candidate in filtered_df.columns:
                    kw_col = col_candidate
                    break
            
            if kw_col:
                kw_options = ["All"] + sorted([str(x) for x in filtered_df[kw_col].dropna().unique()])
                selected_kw = st.selectbox(f"Select or Search {kw_col}:", kw_options, key="kw_filter")
                
                search_df = compute_grouped_table(filtered_df, kw_col, selected_kw)
                if not search_df.empty:
                    st.dataframe(style_dataframe(search_df), use_container_width=True, hide_index=True, height=500)
                    
                    excel_search = convert_df_to_excel(search_df, sheet_name="Search_Term_Performance")
                    st.download_button(
                        label="📥 Download Search Term Performance Excel (.xlsx)",
                        data=excel_search,
                        file_name="Search_Term_Performance_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_search"
                    )
                else:
                    st.info("No search term data matching the selected criteria.")
            else:
                st.info("No Search Term or Keyword column found in the dataset.")

        # TAB 6: Weekly Performance Trend
        with main_tab6:
            st.caption("Weekly aggregated performance trend (Week 1 through Week 5).")
            
            if 'Week' in filtered_df.columns and filtered_df['Week'].notna().any():
                weekly_df = compute_grouped_table(filtered_df, 'Week', "All")
                
                week_order = ['Week 1', 'Week 2', 'Week 3', 'Week 4', 'Week 5']
                weekly_df['Week_Cat'] = pd.Categorical(weekly_df['WEEK'], categories=week_order, ordered=True)
                weekly_df = weekly_df.sort_values('Week_Cat').drop(columns=['Week_Cat'])

                fig = make_subplots(specs=[[{"secondary_y": True}]])

                fig.add_trace(
                    go.Bar(
                        x=weekly_df['WEEK'],
                        y=weekly_df['SPENDS'],
                        name='Spends (₹)',
                        marker=dict(color='#4285F4', line=dict(color='#1A73E8', width=1.5)),
                        text=[f"₹{v:,.0f}" for v in weekly_df['SPENDS']],
                        textposition='auto'
                    ),
                    secondary_y=False
                )

                fig.add_trace(
                    go.Bar(
                        x=weekly_df['WEEK'],
                        y=weekly_df['SALES'],
                        name='Sales (₹)',
                        marker=dict(color='#34A853', line=dict(color='#1E8E3E', width=1.5)),
                        text=[f"₹{v:,.0f}" for v in weekly_df['SALES']],
                        textposition='auto'
                    ),
                    secondary_y=False
                )

                fig.add_trace(
                    go.Scatter(
                        x=weekly_df['WEEK'],
                        y=weekly_df['ROAS'],
                        name='ROAS',
                        mode='lines+markers+text',
                        line=dict(color='#EA4335', width=3),
                        marker=dict(size=8, color='#EA4335'),
                        text=[f"{v:.2f}x" for v in weekly_df['ROAS']],
                        textposition='top center'
                    ),
                    secondary_y=True
                )

                fig.update_layout(
                    title=dict(text="📊 Weekly Budget Spent vs Sales & ROAS Trend", font=dict(size=18, color="#202124")),
                    barmode='group',
                    template='plotly_white',
                    height=520,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    xaxis=dict(title="Week Bucket"),
                    yaxis=dict(title="Amount (₹)", showgrid=True),
                    yaxis2=dict(title="ROAS", overlaying="y", side="right", showgrid=False)
                )

                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(style_dataframe(weekly_df), use_container_width=True, hide_index=True)
            else:
                st.info("No valid Date column found or dates could not be parsed to assign week buckets.")

        st.divider()

        # Output Excel Generation with Formatting for Fallback Replacements
        st.subheader("💾 Download Consolidated Excel Workbook")
        
        buffer_multi = io.BytesIO()
        with pd.ExcelWriter(buffer_multi, engine='openpyxl') as writer:
            for raw_tab_name, df_list in consolidated_raw_tabs.items():
                combined_raw_tab_df = pd.concat(df_list, ignore_index=True)
                clean_sheet_name = raw_tab_name[:31]
                combined_raw_tab_df.to_excel(writer, sheet_name=clean_sheet_name, index=False)
            
            master_export_df = final_df.drop(columns=['_impressions', '_direct_atc', '_indirect_atc', '_atc', '_direct_orders', '_indirect_orders', '_orders', '_direct_sales', '_indirect_sales', '_sales', '_budget_consumed', '_date_dt', 'Ad Type Combined'], errors='ignore')
            
            master_sheet_name = 'Consolidated_Master'
            master_export_df.to_excel(writer, sheet_name=master_sheet_name, index=False)
            
            # Apply yellow highlight formatting in Excel for fallback-filled Match Type values
            if '_filled_fallback' in master_export_df.columns and match_col and match_col in master_export_df.columns:
                worksheet = writer.sheets[master_sheet_name]
                from openpyxl.styles import PatternFill
                yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                
                col_idx = master_export_df.columns.get_loc(match_col) + 1  # 1-indexed for openpyxl
                for row_idx, filled in enumerate(master_export_df['_filled_fallback'], start=2):
                    if filled:
                        worksheet.cell(row=row_idx, column=col_idx).fill = yellow_fill

        buffer_multi.seek(0)

        st.download_button(
            label="📥 Download Complete Excel Workbook (Consolidated Raw Tabs + Final Master Tab)",
            data=buffer_multi,
            file_name="Blinkit_Consolidated_Master_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
