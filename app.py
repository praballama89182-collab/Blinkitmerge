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
st.write("Upload 4 to 5 monthly Excel ad campaign spreadsheets (e.g., `SEPTEMBER.xlsx`, `OCTOBER.xlsx`). The application automatically extracts month names from file titles, handles unknown/new tabs, aligns columns around `PRODUCT_LISTING`, and allows downloading the consolidated file.")

uploaded_files = st.file_uploader(
    "Upload Excel Spreadsheets (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    consolidated_dfs = []

    for uploaded_file in uploaded_files:
        month_name = os.path.splitext(uploaded_file.name)[0].upper()
        xls = pd.ExcelFile(uploaded_file)
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df.insert(0, 'Month', month_name)
            df.insert(1, 'Tab Name', sheet_name)
            consolidated_dfs.append(df)

    if consolidated_dfs:
        final_df = pd.concat(consolidated_dfs, ignore_index=True)
        
        # Determine column order: Month, Tab Name, PRODUCT_LISTING base columns, then remaining extra columns
        base_cols = ['Month', 'Tab Name']
        
        # Check if PRODUCT_LISTING tab exists in any uploaded file to establish base columns
        product_listing_cols = []
        for df in consolidated_dfs:
            if 'PRODUCT_LISTING' in df['Tab Name'].values:
                pl_df = df[df['Tab Name'] == 'PRODUCT_LISTING']
                product_listing_cols = [c for c in pl_df.columns if c not in base_cols]
                break
        
        remaining_cols = [c for c in final_df.columns if c not in base_cols and c not in product_listing_cols]
        ordered_columns = base_cols + product_listing_cols + remaining_cols
        
        # Reindex to ensure consistent column ordering
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
        st.subheader("💾 Download Consolidated Files")
        
        # 1. Consolidated Single-Tab Excel
        buffer_single = io.BytesIO()
        with pd.ExcelWriter(buffer_single, engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name='Consolidated_Master', index=False)
        buffer_single.seek(0)

        st.download_button(
            label="📥 Download Single-Tab Master Excel (.xlsx)",
            data=buffer_single,
            file_name="Consolidated_Master_All_Months.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
