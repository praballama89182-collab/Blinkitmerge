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

st.title("📊 Blinkit Ad Report Merger & Analytics")
st.write("Upload up to 5 monthly Excel ad campaign spreadsheets (.xlsx, .xls, .xlsb, .xlsm). Preview raw & consolidated sheets, inspect campaign performance, and view weekly trend lines.")

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
    raw_files_dict = {}  # {filename: {sheet_name: df}}
    consolidated_raw_tabs = {}  # {sheet_name: [df1, df2, ...]}

    for uploaded_file in uploaded_files:
        month_name = os.path.splitext(uploaded_file.name)[0].upper()
        raw_files_dict[uploaded_file.name] = {}
        
        try:
            xls = pd.ExcelFile(uploaded_file)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                
                # Store raw sheet preview
                raw_files_dict[uploaded_file.name][sheet_name] = df.copy()
                
                # Raw tab copy for consolidated output workbook
                df_raw = df.copy()
                if 'Month' not in df_raw.columns:
                    df_raw.insert(0, 'Month', month_name)
                else:
                    df_raw['Month'] = month_name
                
                if sheet_name not in consolidated_raw_tabs:
                    consolidated_raw_tabs[sheet_name] = []
                consolidated_raw_tabs[sheet_name].append(df_raw)
                
                # Consolidated Master copy
                df_consolidated = df.copy()
                if sheet_name.strip().upper() == 'PRODUCT_RECOMMENDATION':
                    if 'Targeting Type' in df_consolidated.columns:
                        df_consolidated['Match Type'] = df_consolidated['Targeting Type']
                
                df_consolidated.insert(0, 'Month', month_name)
                df_consolidated.insert(1, 'Tab Name', sheet_name)
                consolidated_dfs.append(df_consolidated)
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")

    if consolidated_dfs:
        final_df = pd.concat(consolidated_dfs, ignore_index=True)
        
        # Determine column order
        base_cols = ['Month', 'Tab Name']
        product_listing_cols = []
        for df in consolidated_dfs:
            if 'PRODUCT_LISTING' in df['Tab Name'].values:
                pl_df = df[df['Tab Name'] == 'PRODUCT_LISTING']
                product_listing_cols = [c for c in pl_df.columns if c not in base_cols]
                break
        
        remaining_cols = [c for c in final_df.columns if c not in base_cols and c not in product_listing_cols]
        final_df = final_df.reindex(columns=base_cols + product_listing_cols + remaining_cols)

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

        # --- WEEK BUCKET & DATE RANGE MAPPING LOGIC ---
        date_col = None
        for col_candidate in ['Date', 'date', 'Day', 'DATE']:
            if col_candidate in final_df.columns:
                date_col = col_candidate
                break

        if date_col:
            final_df['_date_dt'] = pd.to_datetime(final_df[date_col], errors='coerce')
            
            # Helper to generate date range string per week bucket
            def assign_week_and_range(row):
                dt = row['_date_dt']
                if pd.isna(dt):
                    return np.nan
                day = dt.day
                month_str = dt.strftime('%b')
                
                if 1 <= day <= 7:
                    return f"Week 1 ({month_str} 01 - {month_str} 07)"
                elif 8 <= day <= 14:
                    return f"Week 2 ({month_str} 08 - {month_str} 14)"
                elif 15 <= day <= 21:
                    return f"Week 3 ({month_str} 15 - {month_str} 21)"
                elif 22 <= day <= 28:
                    return f"Week 4 ({month_str} 22 - {month_str} 28)"
                elif day >= 29:
                    return f"Week 5 ({month_str} 29+)"
                return np.nan

            final_df['Week'] = final_df.apply(assign_week_and_range, axis=1)
        else:
            final_df['Week'] = np.nan

        # --- TOP LEVEL DASHBOARD METRICS (INCLUDING ROAS & ACOS) ---
        total_sales = final_df['_sales'].sum()
        total_orders = final_df['_orders'].sum()
        total_atc = final_df['_atc'].sum()
        total_budget = final_df['_budget_consumed'].sum()
        
        overall_roas = (total_sales / total_budget) if total_budget > 0 else 0.0
        overall_acos = ((total_budget / total_sales) * 100) if total_sales > 0 else 0.0

        st.markdown("### 📈 Overall Campaign Performance Dashboard")
        kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
        
        with kpi1:
            st.metric("Total Sales", f"₹{total_sales:,.2f}")
        with kpi2:
            st.metric("Total Budget Consumed", f"₹{total_budget:,.2f}")
        with kpi3:
            st.metric("Overall RoAS", f"{overall_roas:.2f}x")
        with kpi4:
            st.metric("Overall ACoS", f"{overall_acos:.1f}%")
        with kpi5:
            st.metric("Total Orders", f"{int(total_orders):,}")
        with kpi6:
            st.metric("Total Add To Cart", f"{int(total_atc):,}")

        st.divider()

        # --- FILTERS ABOVE MAIN TABS ---
        st.markdown("### 🔍 Dashboard Filters & Search")
        filter_col1, filter_col2 = st.columns([1, 2])
        
        available_months = ["All Months"] + sorted(list(final_df['Month'].dropna().unique()))
        
        with filter_col1:
            selected_month = st.selectbox("Select Month", available_months)
        
        with filter_col2:
            search_query = st.text_input("Search Across Campaigns / Keywords / Match Types", "").strip()

        # Apply Month Filter
        filtered_df = final_df.copy()
        if selected_month != "All Months":
            filtered_df = filtered_df[filtered_df['Month'] == selected_month]

        # Helper function for grouping metrics
        def compute_grouped_table(df_subset, group_col, search_term=None):
            if group_col not in df_subset.columns:
                return pd.DataFrame()
            
            df_working = df_subset.dropna(subset=[group_col]).copy()
            df_working[group_col] = df_working[group_col].astype(str)
            
            if search_term:
                df_working = df_working[df_working[group_col].str.contains(search_term, case=False, na=False)]
            
            if df_working.empty:
                return pd.DataFrame()

            grouped = df_working.groupby(group_col).agg(
                Impressions=('_impressions', 'sum'),
                Orders=('_orders', 'sum'),
                Sales=('_sales', 'sum'),
                ATC=('_atc', 'sum'),
                Budget_Consumed=('_budget_consumed', 'sum')
            ).reset_index()

            grouped['RoAS'] = grouped.apply(
                lambda r: round(r['Sales'] / r['Budget_Consumed'], 2) if r['Budget_Consumed'] > 0 else 0.0, axis=1
            )
            grouped['ACoS (%)'] = grouped.apply(
                lambda r: round((r['Budget_Consumed'] / r['Sales']) * 100, 2) if r['Sales'] > 0 else 0.0, axis=1
            )

            grouped = grouped.rename(columns={
                group_col: group_col.replace('_', ' ').title(),
                'Budget_Consumed': 'Budget Consumed (₹)',
                'Sales': 'Sales (₹)'
            })
            return grouped

        def style_roas(val):
            try:
                val_float = float(val)
                if val_float < 1.0:
                    return 'background-color: #ffcdd2; color: #b71c1c; font-weight: bold;'
                else:
                    return 'background-color: #c8e6c9; color: #1b5e20; font-weight: bold;'
            except:
                return ''

        def style_dataframe(df):
            styler = df.style
            if hasattr(styler, 'map'):
                styler = styler.map(style_roas, subset=['RoAS'])
            else:
                styler = styler.applymap(style_roas, subset=['RoAS'])
            return styler.format({
                'Sales (₹)': '₹{:,.2f}', 
                'Budget Consumed (₹)': '₹{:,.2f}', 
                'Impressions': '{:,.0f}', 
                'Orders': '{:,.0f}', 
                'ATC': '{:,.0f}'
            })

        # --- MAIN TABS INCLUDING RAW PREVIEWS, CONSOLIDATED PREVIEW, AND PERFORMANCE TABS ---
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
            st.caption("Preview the combined dataset across all uploaded files before export.")
            preview_clean_df = final_df.drop(columns=['_impressions', '_direct_atc', '_indirect_atc', '_atc', '_direct_orders', '_indirect_orders', '_orders', '_direct_sales', '_indirect_sales', '_sales', '_budget_consumed', '_date_dt'], errors='ignore')
            st.write(f"Total Rows Consolidated: **{len(preview_clean_df):,}**")
            st.dataframe(preview_clean_df.head(100), use_container_width=True)

        # TAB 3: Campaign Performance
        with main_tab3:
            st.caption("Aggregated performance metrics per campaign. Click any column header to toggle ascending/descending sort.")
            campaign_df = compute_grouped_table(filtered_df, 'Campaign Name', search_query)
            if not campaign_df.empty:
                st.dataframe(style_dataframe(campaign_df), use_container_width=True, hide_index=True)
            else:
                st.info("No campaign data matching the filter criteria.")

        # TAB 4: Ad Type Performance
        with main_tab4:
            st.caption("Aggregated performance across all Ad Types / Sheet Tabs present in the dataset.")
            adtype_df = compute_grouped_table(filtered_df, 'Tab Name', search_query)
            if not adtype_df.empty:
                st.dataframe(style_dataframe(adtype_df), use_container_width=True, hide_index=True)
            else:
                st.info("No Ad Type data matching the filter criteria.")

        # TAB 5: Keyword / Search Term Performance
        with main_tab5:
            st.caption("Performance across search terms / target keywords (downward scrollable).")
            kw_col = None
            for col_candidate in ['Search Term', 'Keyword', 'Targeting Value']:
                if col_candidate in filtered_df.columns:
                    kw_col = col_candidate
                    break
            
            if kw_col:
                search_df = compute_grouped_table(filtered_df, kw_col, search_query)
                if not search_df.empty:
                    st.dataframe(style_dataframe(search_df), use_container_width=True, hide_index=True, height=500)
                else:
                    st.info("No search term data matching the filter criteria.")
            else:
                st.info("No Search Term or Keyword column found in the dataset.")

        # TAB 6: Weekly Performance Trend with Date Ranges
        with main_tab6:
            st.caption("Weekly aggregated metrics broken down by explicit date range.")
            
            if 'Week' in filtered_df.columns and filtered_df['Week'].notna().any():
                weekly_df = compute_grouped_table(filtered_df, 'Week', None)
                weekly_df = weekly_df.sort_values('Week')

                # Plotly Chart
                fig = make_subplots(specs=[[{"secondary_y": True}]])

                # Budget Consumed Bar
                fig.add_trace(
                    go.Bar(
                        x=weekly_df['Week'],
                        y=weekly_df['Budget Consumed (₹)'],
                        name='Budget Consumed (₹)',
                        marker=dict(color='#4285F4', line=dict(color='#1A73E8', width=1.5)),
                        text=[f"₹{v:,.0f}" for v in weekly_df['Budget Consumed (₹)']],
                        textposition='auto'
                    ),
                    secondary_y=False
                )

                # Sales Bar
                fig.add_trace(
                    go.Bar(
                        x=weekly_df['Week'],
                        y=weekly_df['Sales (₹)'],
                        name='Sales (₹)',
                        marker=dict(color='#34A853', line=dict(color='#1E8E3E', width=1.5)),
                        text=[f"₹{v:,.0f}" for v in weekly_df['Sales (₹)']],
                        textposition='auto'
                    ),
                    secondary_y=False
                )

                # RoAS Trend Line
                fig.add_trace(
                    go.Scatter(
                        x=weekly_df['Week'],
                        y=weekly_df['RoAS'],
                        name='RoAS',
                        mode='lines+markers+text',
                        line=dict(color='#EA4335', width=3),
                        marker=dict(size=8, color='#EA4335'),
                        text=[f"{v:.2f}x" for v in weekly_df['RoAS']],
                        textposition='top center'
                    ),
                    secondary_y=True
                )

                fig.update_layout(
                    title=dict(text="📊 Weekly Budget Spent vs Sales & RoAS Trend", font=dict(size=18, color="#202124")),
                    barmode='group',
                    template='plotly_white',
                    height=520,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    xaxis=dict(title="Week & Date Range"),
                    yaxis=dict(title="Amount (₹)", showgrid=True),
                    yaxis2=dict(title="RoAS", overlaying="y", side="right", showgrid=False)
                )

                st.plotly_chart(fig, use_container_width=True)

                # Weekly Data Table Display
                st.dataframe(style_dataframe(weekly_df), use_container_width=True, hide_index=True)
            else:
                st.info("No valid Date column found or dates could not be parsed to assign week buckets.")

        st.divider()

        # Output Excel Generation
        st.subheader("💾 Download Consolidated Excel Workbook")
        
        buffer_multi = io.BytesIO()
        with pd.ExcelWriter(buffer_multi, engine='openpyxl') as writer:
            for raw_tab_name, df_list in consolidated_raw_tabs.items():
                combined_raw_tab_df = pd.concat(df_list, ignore_index=True)
                clean_sheet_name = raw_tab_name[:31]
                combined_raw_tab_df.to_excel(writer, sheet_name=clean_sheet_name, index=False)
            
            master_export_df = final_df.drop(columns=['_impressions', '_direct_atc', '_indirect_atc', '_atc', '_direct_orders', '_indirect_orders', '_orders', '_direct_sales', '_indirect_sales', '_sales', '_budget_consumed', '_date_dt'], errors='ignore')
            master_export_df.to_excel(writer, sheet_name='Consolidated_Master', index=False)
            
        buffer_multi.seek(0)

        st.download_button(
            label="📥 Download Complete Excel Workbook (Consolidated Raw Tabs + Final Master Tab)",
            data=buffer_multi,
            file_name="Blinkit_Consolidated_Master_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
