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
    page_title="Blinkit Ad Report Merger & Analytics",
    page_icon="📊",
    layout="wide",
)

# Custom CSS for seamless table alignment
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
    "Upload up to 5 monthly Excel ad campaign spreadsheets (.xlsx, .xls,"
    " .xlsb, .xlsm). Preview raw & consolidated sheets, view reverse-chronological"
    " Month-on-Month Comparison tables, and analyze multi-metric 3D pillar & curve"
    " trend graphs."
)

# Month sorting helper to order months serially from latest to oldest (e.g. Sept -> Aug -> Jul -> Jun)
MONTH_MAP = {
    "JANUARY": 1,
    "FEBRUARY": 2,
    "MARCH": 3,
    "APRIL": 4,
    "MAY": 5,
    "JUNE": 6,
    "JULY": 7,
    "AUGUST": 8,
    "SEPTEMBER": 9,
    "OCTOBER": 10,
    "NOVEMBER": 11,
    "DECEMBER": 12,
}


def get_sorted_months(months_list, reverse=True):
  """Sorts months serially in reverse order (e.g., September first, then August, July, June)."""

  def month_key(m):
    m_str = str(m).strip().upper()
    return MONTH_MAP.get(m_str, 0)

  unique_months = list(dict.fromkeys(months_list))
  recognized = [
      m for m in unique_months if str(m).strip().upper() in MONTH_MAP
  ]

  if len(recognized) > 0:
    return sorted(unique_months, key=month_key, reverse=reverse)
  else:
    return unique_months[::-1] if reverse else unique_months


# Helper function to style downloadable Excel Pivot tables cleanly
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
          if str(m_col).upper() == "ACOS":
            val_cell.value = val / 100.0 if val > 1 else val
            val_cell.number_format = "0.00%"
          elif str(m_col).upper() == "CPM":
            val_cell.value = round(val)
            val_cell.number_format = "#,##0"
          else:
            val_cell.value = round(val, 2)
            val_cell.number_format = (
                "#,##0.00" if isinstance(val, float) else "#,##0"
            )
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
          val_cell.value = val / 100.0 if val > 1 else val
          val_cell.number_format = "0.00%"
        elif (
            str(col_name).upper() == "CPM"
            and not is_pct_row
            and isinstance(val, (int, float, np.number))
        ):
          val_cell.value = round(val)
          val_cell.number_format = "#,##0"
        elif isinstance(val, float):
          val_cell.value = round(val, 2)
          val_cell.number_format = "#,##0.00"
        elif isinstance(val, (int, np.integer)):
          val_cell.value = val
          val_cell.number_format = "#,##0"
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
              if str(m_col).upper() == "ACOS":
                val_cell.value = val / 100.0 if val > 1 else val
                val_cell.number_format = "0.00%"
              elif str(m_col).upper() == "CPM":
                val_cell.value = round(val)
                val_cell.number_format = "#,##0"
              else:
                val_cell.value = round(val, 2)
                val_cell.number_format = (
                    "#,##0.00" if isinstance(val, float) else "#,##0"
                )
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
              val_cell.value = val / 100.0 if val > 1 else val
              val_cell.number_format = "0.00%"
            elif (
                str(col_name).upper() == "CPM"
                and not is_pct_row
                and isinstance(val, (int, float, np.number))
            ):
              val_cell.value = round(val)
              val_cell.number_format = "#,##0"
            elif isinstance(val, float):
              val_cell.value = round(val, 2)
              val_cell.number_format = "#,##0.00"
            elif isinstance(val, (int, np.integer)):
              val_cell.value = val
              val_cell.number_format = "#,##0"
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
    target_col = (
        "Targeting Type" if "Targeting Type" in final_df.columns else None
    )
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
        final_df.loc[na_match, "Match Type"] = final_df.loc[
            na_match, tab_col
        ]
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
          lambda r: (
              round((r["Spends"] / r["Impressions"]) * 1000)
              if r["Impressions"] > 0
              else 0
          ),
          axis=1,
      )
      grouped["ROAS"] = grouped.apply(
          lambda r: (
              round(r["Sales"] / r["Spends"], 2) if r["Spends"] > 0 else 0.0
          ),
          axis=1,
      )
      grouped["ACOS"] = grouped.apply(
          lambda r: (
              round((r["Spends"] / r["Sales"]) * 100, 2)
              if r["Sales"] > 0
              else 0.0
          ),
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
      # Arrange months serially in reverse chronological order (e.g. September -> August -> July -> June)
      all_months = get_sorted_months(df_input["Month"].unique(), reverse=True)

      pivot_df = pivot_df.reorder_levels([1, 0], axis=1)
      sorted_cols = pd.MultiIndex.from_product(
          [all_months, metrics_order], names=["Month", "Metric"]
      )
      pivot_df = pivot_df.reindex(columns=sorted_cols).fillna(0)

      grand_total_series = {}
      for month in all_months:
        month_df = working_df[working_df["Month"] == month]
        total_imp = month_df["_impressions"].sum()
        total_atc = month_df["_atc"].sum()
        total_orders = month_df["_orders"].sum()
        total_spends = round(month_df["_budget_consumed"].sum(), 2)
        total_sales = round(month_df["_sales"].sum(), 2)

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
      return pivot_df

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
          lambda r: (
              round((r["Spends"] / r["Impressions"]) * 1000)
              if r["Impressions"] > 0
              else 0
          ),
          axis=1,
      )
      monthly_agg["ROAS"] = monthly_agg.apply(
          lambda r: (
              round(r["Sales"] / r["Spends"], 2) if r["Spends"] > 0 else 0.0
          ),
          axis=1,
      )
      monthly_agg["ACOS"] = monthly_agg.apply(
          lambda r: (
              round((r["Spends"] / r["Sales"]) * 100, 2)
              if r["Sales"] > 0
              else 0.0
          ),
          axis=1,
      )

      monthly_agg["Spends"] = monthly_agg["Spends"].round(2)
      monthly_agg["Sales"] = monthly_agg["Sales"].round(2)

      # Sort months serially from latest to oldest
      all_months = get_sorted_months(df_input["Month"].unique(), reverse=True)
      monthly_agg["month_order"] = monthly_agg["Month"].map(
          lambda x: all_months.index(x) if x in all_months else 99
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
        curr_row = monthly_agg.iloc[0]  # Latest Month (e.g. September)
        prev_row = monthly_agg.iloc[1]  # Prior Month (e.g. August)

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
          curr_val = curr_row[metric]
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

      monthly_agg["ACOS"] = monthly_agg["ACOS"].apply(
          lambda v: f"{v:.2f}%" if isinstance(v, (int, float)) else str(v)
      )
      return monthly_agg

    def render_unified_single_table(
        df_to_show, key_prefix="mom", expandable_col=None
    ):
      if df_to_show is None or df_to_show.empty:
        st.info("No data available for display.")
        return

      working_df = df_to_show.copy()

      bottom_rows = None
      if (
          isinstance(working_df.index, pd.Index)
          and "Grand Total" in working_df.index
      ):
        main_df = working_df.drop("Grand Total")
        bottom_rows = working_df.loc[["Grand Total"]]
      elif (
          "Month" in working_df.columns
          and "Percentage %" in working_df["Month"].values
      ):
        main_df = working_df[working_df["Month"] != "Percentage %"].copy()
        bottom_rows = working_df[working_df["Month"] == "Percentage %"]
      else:
        main_df = working_df

      st.dataframe(main_df, use_container_width=True)

      if bottom_rows is not None and not bottom_rows.empty:
        st.markdown("**Totals / Growth Summary:**")
        st.dataframe(bottom_rows, use_container_width=True)

    # Main Tabs UI
    tab1, tab2, tab3 = st.tabs([
        "📈 Month-on-Month Comparison",
        "📊 3D Trend Analytics & Charts",
        "📁 Raw & Consolidated Sheets",
    ])

    # TAB 1: MoM Comparison Tables
    with tab1:
      st.subheader("📅 Month-on-Month Performance Summary")

      monthly_summary_df = create_monthly_summary_table(final_df)
      render_unified_single_table(
          monthly_summary_df, key_prefix="monthly_summary"
      )

      if not monthly_summary_df.empty:
        summary_bytes = style_and_export_pivot(
            monthly_summary_df, sheet_name="Monthly_Summary"
        )
        st.download_button(
            label="📥 Download Monthly Summary Excel",
            data=summary_bytes,
            file_name="Blinkit_Monthly_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

      st.markdown("---")
      st.subheader("📊 Category & Campaign MoM Breakdown")

      mom_pivots = {}

      if "Campaign Name" in final_df.columns:
        st.markdown("#### 🎯 Campaign Name MoM")
        camp_pivot = create_mom_comparison_table(final_df, "Campaign Name")
        render_unified_single_table(camp_pivot, key_prefix="camp_mom")
        mom_pivots["Campaign_Name"] = camp_pivot

      if "Ad Type Combined" in final_df.columns:
        st.markdown("#### 🏷️ Match / Ad Type MoM")
        ad_pivot = create_mom_comparison_table(final_df, "Ad Type Combined")
        render_unified_single_table(ad_pivot, key_prefix="ad_mom")
        mom_pivots["Ad_Type"] = ad_pivot

      if "Tab Name" in final_df.columns:
        st.markdown("#### 📑 Placement / Tab Name MoM")
        tab_pivot = create_mom_comparison_table(final_df, "Tab Name")
        render_unified_single_table(tab_pivot, key_prefix="tab_mom")
        mom_pivots["Placement_Tab"] = tab_pivot

      if mom_pivots:
        mom_pivots["Monthly_Summary"] = monthly_summary_df
        all_pivots_bytes = convert_all_pivots_to_excel(mom_pivots)
        st.markdown("---")
        st.download_button(
            label="📦 Download All Comparison Pivots (Combined Excel)",
            data=all_pivots_bytes,
            file_name="Blinkit_All_MoM_Comparisons.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # TAB 2: 3D Visual Trend Charts with Cool Color Nodes & Matching Trendlines
    with tab2:
      st.subheader(
          "🌐 3D Curved Trend Lines & Multi-Metric Pillar Combination Graph"
      )
      st.write(
          "Interactive 3D visualization featuring cool-colored nodes, matching"
          " curve trendlines, and 3D metric pillars across months."
      )

      # Monthly aggregation
      monthly_agg_df = (
          final_df.groupby("Month")
          .agg(
              Impressions=("_impressions", "sum"),
              ATC=("_atc", "sum"),
              Orders=("_orders", "sum"),
              Spends=("_budget_consumed", "sum"),
              Sales=("_sales", "sum"),
          )
          .reset_index()
      )

      all_m = get_sorted_months(final_df["Month"].unique(), reverse=True)
      monthly_agg_df["m_order"] = monthly_agg_df["Month"].map(
          lambda x: all_m.index(x) if x in all_m else 99
      )
      monthly_agg_df = monthly_agg_df.sort_values("m_order")

      monthly_agg_df["ROAS"] = monthly_agg_df.apply(
          lambda r: (
              round(r["Sales"] / r["Spends"], 2) if r["Spends"] > 0 else 0.0
          ),
          axis=1,
      )
      monthly_agg_df["ACOS"] = monthly_agg_df.apply(
          lambda r: (
              round((r["Spends"] / r["Sales"]) * 100, 2)
              if r["Sales"] > 0
              else 0.0
          ),
          axis=1,
      )

      # Multi-metric selector for 3D Pillars and Trend Curves
      available_metrics = [
          "Spends",
          "Sales",
          "Impressions",
          "Orders",
          "ATC",
          "ROAS",
      ]
      selected_metrics = st.multiselect(
          "Select Metrics to Display in 3D (Multiple metrics create multiple 3D bar pillars & trendlines):",
          options=available_metrics,
          default=["Spends", "Sales", "Orders"],
      )

      # Cool color nodes palette definition
      COOL_COLOR_PALETTE = {
          "Spends": "#00F0FF",  # Electric Cyan
          "Sales": "#FF007F",  # Neon Magenta
          "Impressions": "#00FFAB",  # Cool Mint
          "Orders": "#7000FF",  # Deep Violet
          "ATC": "#FF9F1C",  # Vibrant Coral
          "ROAS": "#3A86FF",  # Electric Blue
          "ACOS": "#E63946",  # Cool Crimson
      }

      if selected_metrics:
        fig_3d = go.Figure()

        for metric in selected_metrics:
          color = COOL_COLOR_PALETTE.get(metric, "#00F0FF")

          # 1. Add 3D Curved Trendline connecting cool-colored nodes
          fig_3d.add_trace(
              go.Scatter3d(
                  x=monthly_agg_df["Month"],
                  y=[metric] * len(monthly_agg_df),
                  z=monthly_agg_df[metric],
                  mode="lines+markers",
                  name=f"{metric} Curve",
                  line=dict(color=color, width=7),
                  marker=dict(
                      size=8,
                      color=color,
                      symbol="circle",
                      opacity=0.9,
                      line=dict(color="#FFFFFF", width=1),
                  ),
                  hovertemplate=(
                      f"<b>Month:</b> %{{x}}<br><b>Metric:</b> {metric}<br><b>Value:</b> %{{z:,.2f}}<extra></extra>"
                  ),
              )
          )

          # 2. Add 3D Vertical Pillar Bars for selected metric
          for _, row in monthly_agg_df.iterrows():
            fig_3d.add_trace(
                go.Scatter3d(
                    x=[row["Month"], row["Month"]],
                    y=[metric, metric],
                    z=[0, row[metric]],
                    mode="lines",
                    line=dict(color=color, width=12),
                    showlegend=False,
                    hoverinfo="skip",
                )
            )

        fig_3d.update_layout(
            title=(
                "🌐 Multi-Metric 3D Curved Trend Lines & Pillar Combination"
                " Graph"
            ),
            scene=dict(
                xaxis_title="Month",
                yaxis_title="Metric Category",
                zaxis_title="Metric Value",
                aspectmode="manual",
                aspectratio=dict(x=1.6, y=1.1, z=0.8),
                bgcolor="#0E1117",
                xaxis=dict(
                    backgroundcolor="#0E1117",
                    gridcolor="#262730",
                    showbackground=True,
                    zerolinecolor="#262730",
                ),
                yaxis=dict(
                    backgroundcolor="#0E1117",
                    gridcolor="#262730",
                    showbackground=True,
                    zerolinecolor="#262730",
                ),
                zaxis=dict(
                    backgroundcolor="#0E1117",
                    gridcolor="#262730",
                    showbackground=True,
                    zerolinecolor="#262730",
                ),
            ),
            paper_bgcolor="#0E1117",
            plot_bgcolor="#0E1117",
            template="plotly_dark",
            height=650,
            margin=dict(l=10, r=10, b=10, t=50),
        )

        st.plotly_chart(fig_3d, use_container_width=True)

      st.markdown("---")
      col_c1, col_c2 = st.columns(2)

      # 3D Spends vs Sales Comparison Chart
      with col_c1:
        fig_sp_sa = go.Figure()

        # Spends Pillars
        for _, r in monthly_agg_df.iterrows():
          fig_sp_sa.add_trace(
              go.Scatter3d(
                  x=[r["Month"], r["Month"]],
                  y=["Spends", "Spends"],
                  z=[0, r["Spends"]],
                  mode="lines",
                  line=dict(color="#00F0FF", width=14),
                  showlegend=False,
                  hoverinfo="skip",
              )
          )

        fig_sp_sa.add_trace(
            go.Scatter3d(
                x=monthly_agg_df["Month"],
                y=["Spends"] * len(monthly_agg_df),
                z=monthly_agg_df["Spends"],
                mode="lines+markers",
                name="Spends (₹)",
                line=dict(color="#00F0FF", width=6),
                marker=dict(size=8, color="#00F0FF"),
            )
        )

        # Sales Pillars
        for _, r in monthly_agg_df.iterrows():
          fig_sp_sa.add_trace(
              go.Scatter3d(
                  x=[r["Month"], r["Month"]],
                  y=["Sales", "Sales"],
                  z=[0, r["Sales"]],
                  mode="lines",
                  line=dict(color="#FF007F", width=14),
                  showlegend=False,
                  hoverinfo="skip",
              )
          )

        fig_sp_sa.add_trace(
            go.Scatter3d(
                x=monthly_agg_df["Month"],
                y=["Sales"] * len(monthly_agg_df),
                z=monthly_agg_df["Sales"],
                mode="lines+markers",
                name="Sales (₹)",
                line=dict(color="#FF007F", width=6),
                marker=dict(size=8, color="#FF007F"),
            )
        )

        fig_sp_sa.update_layout(
            title="3D Spends vs Sales Comparison",
            scene=dict(
                xaxis_title="Month",
                yaxis_title="Type",
                zaxis_title="Amount (₹)",
                bgcolor="#0E1117",
            ),
            paper_bgcolor="#0E1117",
            template="plotly_dark",
            height=450,
        )
        st.plotly_chart(fig_sp_sa, use_container_width=True)

      # 3D ROAS & ACOS Efficiency Trend Curves
      with col_c2:
        fig_roas_acos = go.Figure()

        fig_roas_acos.add_trace(
            go.Scatter3d(
                x=monthly_agg_df["Month"],
                y=["ROAS"] * len(monthly_agg_df),
                z=monthly_agg_df["ROAS"],
                mode="lines+markers",
                name="ROAS Curve",
                line=dict(color="#3A86FF", width=6),
                marker=dict(size=9, color="#3A86FF", symbol="diamond"),
            )
        )

        fig_roas_acos.add_trace(
            go.Scatter3d(
                x=monthly_agg_df["Month"],
                y=["ACOS"] * len(monthly_agg_df),
                z=monthly_agg_df["ACOS"],
                mode="lines+markers",
                name="ACOS (%) Curve",
                line=dict(color="#FF9F1C", width=6),
                marker=dict(size=9, color="#FF9F1C", symbol="circle"),
            )
        )

        fig_roas_acos.update_layout(
            title="3D ROAS & ACOS Trend Curves",
            scene=dict(
                xaxis_title="Month",
                yaxis_title="Metric",
                zaxis_title="Ratio / %",
                bgcolor="#0E1117",
            ),
            paper_bgcolor="#0E1117",
            template="plotly_dark",
            height=450,
        )
        st.plotly_chart(fig_roas_acos, use_container_width=True)

    # TAB 3: Raw & Consolidated Data Preview
    with tab3:
      st.subheader("📋 Consolidated Master Spreadsheet")
      st.write(
          f"Total Merged Records: **{len(final_df):,} rows** across"
          f" **{len(final_df['Month'].unique())} month(s)**."
      )

      st.dataframe(final_df.head(200), use_container_width=True)

      csv_buffer = io.BytesIO()
      final_df.to_csv(csv_buffer, index=False)
      st.download_button(
          label="📥 Download Consolidated CSV Data",
          data=csv_buffer.getvalue(),
          file_name="Blinkit_Consolidated_Master_Data.csv",
          mime="text/csv",
      )

      st.markdown("---")
      st.subheader("🔍 Uploaded File Sheet Inspector")
      selected_file = st.selectbox(
          "Select File to Inspect Sheets:", list(raw_files_dict.keys())
      )
      if selected_file:
        selected_sheet = st.selectbox(
            "Select Sheet:", list(raw_files_dict[selected_file].keys())
        )
        if selected_sheet:
          st.dataframe(
              raw_files_dict[selected_file][selected_sheet].head(100),
              use_container_width=True,
          )
