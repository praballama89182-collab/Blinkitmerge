import os
import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Quick-Commerce Multi-File Ad Campaign Data Consolidation",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Quick-Commerce Multi-File Ad Campaign Consolidation")
st.write("Upload up to 5 monthly Excel ad campaign spreadsheets (.xlsx, .xls, .xlsb, .xlsm). View high-level metrics, automatically consolidate sheets, and download the unified workbook.")

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

        # Helper function for safe numeric summation
        def safe_sum(columns):
            total = 0.0
            for col in columns:
                if col in final_df.columns:
                    total += pd.to_numeric(final_df[col], errors='coerce').fillna(0).sum()
            return total

        # Top Dashboard KPI Calculations (Direct + Indirect combined)
        total_sales = safe_sum(['Direct Sales', 'Indirect Sales'])
        total_orders = safe_sum(['Direct Quantities Sold', 'Indirect Quantities Sold', 'Direct Orders', 'Indirect Orders'])
        total_atc = safe_sum(['Direct ATC', 'Indirect ATC'])
        total_budget = safe_sum(['Estimated Budget Consumed', 'Budget Consumed', 'Spend'])

        # Top Dashboard Metrics Display
        st.markdown("### 📈 Campaign Performance Overview")
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

        # Data Preview
        st.subheader("📌 Consolidated Master Preview")
        st.dataframe(final_df.head(50), use_container_width=True)

        # Output Excel Generation
        st.subheader("💾 Download Workbook")
        
        buffer_multi = io.BytesIO()
        with pd.ExcelWriter(buffer_multi, engine='openpyxl') as writer:
            # 1. Write consolidated raw tabs (combining data across uploaded files for each tab format)
            for raw_tab_name, df_list in consolidated_raw_tabs.items():
                combined_raw_tab_df = pd.concat(df_list, ignore_index=True)
                clean_sheet_name = raw_tab_name[:31]  # Excel 31 character sheet name limit
                combined_raw_tab_df.to_excel(writer, sheet_name=clean_sheet_name, index=False)
            
            # 2. Write the final master consolidated summary on the LAST tab
            final_df.to_excel(writer, sheet_name='Consolidated_Master', index=False)
            
        buffer_multi.seek(0)

        st.download_button(
            label="📥 Download Complete Excel Workbook (Consolidated Raw Tabs + Final Master Tab)",
            data=buffer_multi,
            file_name="All_Months_Consolidated_Master.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
