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
st.write("Upload up to 5 monthly Excel ad campaign spreadsheets (.xlsx, .xls, .xlsb, .xlsm). All raw tabs across all uploaded files are preserved, and the final consolidated output is generated on the last tab.")

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
    raw_tabs_dict = {}

    for uploaded_file in uploaded_files:
        month_name = os.path.splitext(uploaded_file.name)[0].upper()
        
        try:
            xls = pd.ExcelFile(uploaded_file)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                
                # Raw tab preservation with month prefix (Excel max sheet name length is 31 chars)
                unique_tab_label = f"{month_name}_{sheet_name}"[:31]
                raw_tabs_dict[unique_tab_label] = df.copy()
                
                # Copy dataframe for consolidated view
                df_consolidated = df.copy()
                
                # Rule: For PRODUCT_RECOMMENDATION, set Match Type exactly equal to Targeting Type
                if sheet_name.strip().upper() == 'PRODUCT_RECOMMENDATION':
                    if 'Targeting Type' in df_consolidated.columns:
                        df_consolidated['Match Type'] = df_consolidated['Targeting Type']
                
                # Insert tracking columns for consolidated view
                df_consolidated.insert(0, 'Month', month_name)
                df_consolidated.insert(1, 'Tab Name', sheet_name)
                consolidated_dfs.append(df_consolidated)
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")

    if consolidated_dfs:
        final_df = pd.concat(consolidated_dfs, ignore_index=True)
        
        # Order columns: Month, Tab Name, PRODUCT_LISTING base columns, then remaining extra columns
        base_cols = ['Month', 'Tab Name']
        
        product_listing_cols = []
        for df in consolidated_dfs:
            if 'PRODUCT_LISTING' in df['Tab Name'].values:
                pl_df = df[df['Tab Name'] == 'PRODUCT_LISTING']
                product_listing_cols = [c for c in pl_df.columns if c not in base_cols]
                break
        
        remaining_cols = [c for c in final_df.columns if c not in base_cols and c not in product_listing_cols]
        final_df = final_df.reindex(columns=base_cols + product_listing_cols + remaining_cols)

        st.success(f"Successfully processed {len(uploaded_files)} file(s) across {len(consolidated_dfs)} tab(s)! Total Rows: {len(final_df)}")

        # Data Preview & Key KPIs
        st.subheader("📌 Consolidated Data Preview")
        st.dataframe(final_df.head(50), use_container_width=True)

        kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
        with kpi_col1:
            st.metric("Total Rows", len(final_df))
        with kpi_col2:
            st.metric("Total Columns", len(final_df.columns))
        with kpi_col3:
            if 'Estimated Budget Consumed' in final_df.columns:
                total_spend = final_df['Estimated Budget Consumed'].sum()
                st.metric("Total Ad Spend (₹)", f"₹{total_spend:,.2f}")

        # Download Section
        st.subheader("💾 Download Consolidated Multi-Tab Excel Workbook")
        
        buffer_multi = io.BytesIO()
        with pd.ExcelWriter(buffer_multi, engine='openpyxl') as writer:
            # 1. First, write all individual raw tabs
            for sheet_title, sheet_data in raw_tabs_dict.items():
                sheet_data.to_excel(writer, sheet_name=sheet_title, index=False)
            
            # 2. Finally, place the consolidated output on the LAST tab
            final_df.to_excel(writer, sheet_name='Consolidated_Master', index=False)
            
        buffer_multi.seek(0)

        st.download_button(
            label="📥 Download Workbook (Raw Tabs + Final Consolidated Tab)",
            data=buffer_multi,
            file_name="All_Months_Consolidated_Master.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
