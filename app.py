import os
import io
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(
    page_title="Blinkit Report Merger",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Blinkit Report Merger")
st.write("Upload up to 5 monthly Excel ad campaign spreadsheets (.xlsx, .xls, .xlsb, .xlsm). View performance breakdowns, weekly trends, automatically consolidate sheets, and download the unified workbook.")

# 5 Dedicated Upload Boxes
col1, col2, col3, col4, col5 = st.columns(5)

uploaded_files = []

with col1:
    file1 = st.file_uploader("Upload File 1", type=["xlsx", "xls", "xlsb", "xlsm"], key="file1")
    if file1:
        uploaded_files.append(file1)

with col2:
    file2 = st.file_uploader("Upload File 2", type=["xlsx", "xls", "xlsb", "xlsm"], key="file2")
    if file2:
        uploaded_files.append(file2)

with col3:
    file3 = st.file_uploader("Upload File 3", type=["xlsx", "xls", "xlsb", "xlsm"], key="file3")
    if file3:
        uploaded_files.append(file3)

with col4:
    file4 = st.file_uploader("Upload File 4", type=["xlsx", "xls", "xlsb", "xlsm"], key="file4")
    if file4:
        uploaded_files.append(file4)

with col5:
    file5 = st.file_uploader("Upload File 5", type=["xlsx", "xls", "xlsb", "xlsm"], key="file5")
    if file5:
        uploaded_files.append(file5)

if uploaded_files:
    consolidated_dfs = []
    # Dict to collect dataframes grouped by standard tab name across all uploaded files
    consolidated_raw_tabs = {}

    for uploaded_file in uploaded_files:
        month_name = os.path.splitext(uploaded_file.name)[0].upper()
        
        try:
            xls = pd.ExcelFile(uploaded_file)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                
                # Copy for consolidated raw tab (adds Month column while preserving original raw structure)
                df_raw = df.copy()
                if 'Month' not in df_raw.columns:
                    df_raw.insert(0, 'Month', month_name)
                else:
                    df_raw['Month'] = month_name
                
                if sheet_name not in consolidated_raw_tabs:
                    consolidated_raw_tabs[sheet_name] = []
                consolidated_raw_tabs[sheet_name].append(df_raw)
                
                # Copy for the final Consolidated Master tab
                df_consolidated = df.copy()
                
                # Match Type mapping rule for PRODUCT_RECOMMENDATION
                if sheet_name.strip().upper() == 'PRODUCT_RECOMMENDATION':
                    if 'Targeting Type' in df_consolidated.columns:
                        df_consolidated['Match Type'] = df_consolidated['Targeting Type']
                
                # Insert Month and Tab Name tracking columns for master output
                df_consolidated.insert(0, 'Month', month_name)
                df_consolidated.insert(1, 'Tab Name', sheet_name)
                consolidated_dfs.append(df_consolidated)
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")

    if consolidated_dfs:
        final_df = pd.concat(consolidated_dfs, ignore_index=True)
        
        # Determine column order: Month, Tab Name, PRODUCT_LISTING base columns, then remaining extra columns
        base_cols = ['Month', 'Tab Name']
        
        product_listing_cols = []
        for df in consolidated_dfs:
            if 'PRODUCT_LISTING' in df['Tab Name'].values:
                pl_df = df[df['Tab Name'] == 'PRODUCT_LISTING']
                product_listing_cols = [c for c in pl_df.columns if c not in base_cols]
                break
        
        remaining_cols = [c for c in final_df.columns if c not in base_cols and c not in product_listing_cols]
        final_df = final_df.reindex(columns=base_cols + product_listing_cols + remaining_cols)

        # Helper function for safe numeric column extraction
        def get_numeric_col(df, possible_cols):
            for col in possible_cols:
                if col in df.columns:
                    return pd.to_numeric(df[col], errors='coerce').fillna(0)
            return pd.Series(0, index=df.index)

        # Precompute standard numeric metrics on final_df for unified calculations
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

        # --- WEEK BUCKET MAPPING LOGIC ---
        if 'Date' in final_df.columns:
            final_df['_date_dt'] = pd.to_datetime(final_df['Date'], errors='coerce')
            final_df['_day'] = final_df['_date_dt'].dt.day

            def assign_week_bucket(day):
                if pd.isna(day):
                    return np.nan
                day = int(day)
                if 1 <= day <= 7:
                    return "Week 1 (1-7)"
                elif 8 <= day <= 14:
                    return "Week 2 (8-14)"
                elif 15 <= day <= 21:
                    return "Week 3 (15-21)"
                elif 22 <= day <= 28:
                    return "Week 4 (22-28)"
                elif day >= 29:
                    return "Week 5 (29+)"
                return np.nan

            final_df['Week'] = final_df['_day'].apply(assign_week_bucket)
        else:
            final_df['Week'] = np.nan

        # --- TOP LEVEL DASHBOARD METRICS ---
        total_sales = final_df['_sales'].sum()
        total_orders = final_df['_orders'].sum()
        total_atc = final_df['_atc'].sum()
        total_budget = final_df['_budget_consumed'].sum()

        st.markdown("### 📈 Overall Campaign Performance Overview")
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        
        with kpi1:
            st.metric("Total Sales (Direct + Indirect)", f"₹{total_sales:,.2f}")
        with kpi2:
            st.metric("Total Orders (Direct + Indirect)", f"{int(total_orders):,}")
        with kpi3:
            st.metric("Total Add To Cart (Direct + Indirect)", f"{int(total_atc):,}")
        with kpi4:
            st.metric("Total Budget Consumed", f"₹{total_budget:,.2f}")

        st.divider()

        # --- FILTERS ABOVE TABS ---
        st.markdown("### 🔍 Dashboard Filters & Search")
        filter_col1, filter_col2 = st.columns([1, 2])
        
        available_months = ["All Months"] + sorted(list(final_df['Month'].dropna().unique()))
        
        with filter_col1:
            selected_month = st.selectbox("Select Month", available_months)
        
        with filter_col2:
            search_query = st.text_input("Search (Campaign Name, Search Term, or Match/Targeting Type)", "").strip()

        # Apply Month Filter
        filtered_df = final_df.copy()
        if selected_month != "All Months":
            filtered_df = filtered_df[filtered_df['Month'] == selected_month]

        # Helper function to compute grouped metrics table with RoAS and ACoS
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

            # Derived Metrics: RoAS & ACoS
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

        # RoAS Highlighting Function (Red < 1, Green >= 1)
        def style_roas(val):
            try:
                val_float = float(val)
                if val_float < 1.0:
                    return 'background-color: #ffcdd2; color: #b71c1c; font-weight: bold;'
                else:
                    return 'background-color: #c8e6c9; color: #1b5e20; font-weight: bold;'
            except:
                return ''

        # Safe styling helper compatible across pandas versions
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

        # --- DASHBOARD BREAKDOWN TABS ---
        st.markdown("### 📊 Performance Analytics Breakdown")
        tab1, tab2, tab3, tab4 = st.tabs([
            "🎯 Campaign Performance", 
            "🔎 Search Term Performance", 
            "📢 Ad Type Performance",
            "📅 Weekly Performance"
        ])

        # TAB 1: Consolidated Performance per Campaign
        with tab1:
            st.caption("Aggregated performance metrics per campaign. Click any column header to toggle ascending/descending sort.")
            campaign_df = compute_grouped_table(filtered_df, 'Campaign Name', search_query)
            if not campaign_df.empty:
                st.dataframe(style_dataframe(campaign_df), use_container_width=True, hide_index=True)
            else:
                st.info("No campaign data matching the filter criteria.")

        # TAB 2: Search Term Performance
        with tab2:
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

        # TAB 3: Ad Type Performance
        with tab3:
            st.caption("Aggregated performance across all Ad Types / Sheet Tabs present in the consolidated file.")
            adtype_df = compute_grouped_table(filtered_df, 'Tab Name', search_query)
            if not adtype_df.empty:
                st.dataframe(style_dataframe(adtype_df), use_container_width=True, hide_index=True)
            else:
                st.info("No Ad Type data matching the filter criteria.")

        # TAB 4: Weekly Performance
        with tab4:
            st.caption("Weekly aggregated metrics (Week 1: Days 1-7, Week 2: Days 8-14, Week 3: Days 15-21, Week 4: Days 22-28, Week 5: Days 29+).")
            
            if 'Week' in filtered_df.columns and filtered_df['Week'].notna().any():
                weekly_raw = compute_grouped_table(filtered_df, 'Week', None)
                
                # Order weeks chronologically
                week_order = ["Week 1 (1-7)", "Week 2 (8-14)", "Week 3 (15-21)", "Week 4 (22-28)", "Week 5 (29+)"]
                weekly_raw['sort_key'] = weekly_raw['Week'].apply(lambda x: week_order.index(x) if x in week_order else 99)
                weekly_df = weekly_raw.sort_values('sort_key').drop(columns=['sort_key'])

                # Display Bar Chart + RoAS Trend Line using Plotly
                fig = make_subplots(specs=[[{"secondary_y": True}]])

                # Budget Consumed Bar
                fig.add_trace(
                    go.Bar(
                        x=weekly_df['Week'],
                        y=weekly_df['Budget Consumed (₹)'],
                        name='Budget Consumed (₹)',
                        marker=dict(
                            color='#4285F4',
                            line=dict(color='#1A73E8', width=1.5)
                        ),
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
                        marker=dict(
                            color='#34A853',
                            line=dict(color='#1E8E3E', width=1.5)
                        ),
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
                        text=[f"{v:.2f}" for v in weekly_df['RoAS']],
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
                    xaxis=dict(title="Week Period"),
                    yaxis=dict(title="Amount (₹)", showgrid=True),
                    yaxis2=dict(title="RoAS", overlaying="y", side="right", showgrid=False)
                )

                st.plotly_chart(fig, use_container_width=True)

                # Weekly Data Table Display
                st.dataframe(style_dataframe(weekly_df), use_container_width=True, hide_index=True)
            else:
                st.info("No valid Date information found in uploaded data to compute weekly metrics.")

        st.divider()

        # Data Preview of Consolidated Master
        st.subheader("📌 Consolidated Master Dataset Preview")
        st.dataframe(final_df.drop(columns=['_impressions', '_direct_atc', '_indirect_atc', '_atc', '_direct_orders', '_indirect_orders', '_orders', '_direct_sales', '_indirect_sales', '_sales', '_budget_consumed', '_date_dt', '_day'], errors='ignore').head(50), use_container_width=True)

        # Output Excel Generation
        st.subheader("💾 Download Workbook")
        
        buffer_multi = io.BytesIO()
        with pd.ExcelWriter(buffer_multi, engine='openpyxl') as writer:
            # 1. Write consolidated raw tabs (combining data across uploaded files for each tab format)
            for raw_tab_name, df_list in consolidated_raw_tabs.items():
                combined_raw_tab_df = pd.concat(df_list, ignore_index=True)
                clean_sheet_name = raw_tab_name[:31]  # Excel 31 character sheet limit
                combined_raw_tab_df.to_excel(writer, sheet_name=clean_sheet_name, index=False)
            
            # 2. Write the final master consolidated summary on the LAST tab
            master_export_df = final_df.drop(columns=['_impressions', '_direct_atc', '_indirect_atc', '_atc', '_direct_orders', '_indirect_orders', '_orders', '_direct_sales', '_indirect_sales', '_sales', '_budget_consumed', '_date_dt', '_day'], errors='ignore')
            master_export_df.to_excel(writer, sheet_name='Consolidated_Master', index=False)
            
        buffer_multi.seek(0)

        st.download_button(
            label="📥 Download Complete Excel Workbook (Consolidated Raw Tabs + Final Master Tab)",
            data=buffer_multi,
            file_name="Blinkit_Consolidated_Master_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
