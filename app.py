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
st.write("Upload multiple monthly Excel ad campaign spreadsheets (e.g., `.xlsx`, `.xls`, `.xlsb`, `.xlsm`). The application automatically consolidates data across files/tabs, retains all raw tabs with unique month-prefixed tab names, and puts the consolidated master summary on the final tab.")

uploaded_files = st.file_uploader(
    "Upload Excel Spreadsheets (.xlsx, .xls, .xlsb, .xlsm)",
    type=["xlsx", "xls", "xlsb", "xlsm"],
    accept_multiple_files=True
)

if uploaded_files:
    consolidated_dfs = []
    raw_tabs_dict = {}

    for uploaded_file in uploaded_files:
        month_name = os.path.splitext(uploaded_file.name)[0].upper()
        
        try:
            xls = pd.ExcelFile(uploaded_file)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                
                # Dynamic tab naming for multi-file sheet preservation
                unique_tab_label = f"{month_name}_{sheet_name}"
                raw_tabs_dict[unique_tab_label[:31]] = df.copy()  # Excel max sheet name length is 31 chars
                
                # Prepare row data for consolidation master
                df_consolidated = df.copy()
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
        ordered_columns = base_cols + product_listing_cols + remaining_cols
        
        final_df = final_df.reindex(columns=ordered_columns)

        st.success(f"Successfully processed {len(uploaded_files)} file(s) across {len(consolidated_dfs)} tab(s)! Total Rows: {len(final_df)}")

        # Dashboard KPIs & Preview
        st.subheader("📌 Consolidated Data Preview")
        st.dataframe(final_df.head(50), use_container_width=True)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Rows", len(final_df))
        with col2:
            st.metric("Total Columns", len(final_df.columns))
        with col3:
            if 'Estimated Budget Consumed' in final_df.columns:
                total_spend = final_df['Estimated Budget Consumed'].sum()
                st.metric("Total Ad Spend (₹)", f"₹{total_spend:,.2f}")

        # Download Section
        st.subheader("💾 Download Consolidated Multi-Tab Excel Workbook")
        
        buffer_multi = io.BytesIO()
        with pd.ExcelWriter(buffer_multi, engine='openpyxl') as writer:
            # 1. First, write all raw individual tabs
            for sheet_title, sheet_data in raw_tabs_dict.items():
                sheet_data.to_excel(writer, sheet_name=sheet_title, index=False)
            
            # 2. Finally, place the consolidated output on the LAST tab
            final_df.to_excel(writer, sheet_name='Consolidated_Master', index=False)
            
        buffer_multi.seek(0)

        st.download_button(
            label="📥 Download Complete Excel Workbook (Raw Tabs + Final Consolidated Tab)",
            data=buffer_multi,
            file_name="All_Months_Consolidated_Master.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
