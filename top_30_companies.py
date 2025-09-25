import pandas as pd
import streamlit as st
from typing import Optional

# Set Streamlit page configuration
st.set_page_config(page_title='TOP 30 COMPANIES', page_icon=':bar_chart:', layout = "wide")

# Display header for the dashboard
st.header('TOP 30 COMPANIES')

# Display the last update date
st.markdown('#### Updated: 25/09/2025')

excel_file = 'TOP30COMPANIES.xlsx'
value = 'VALUE'
quality = 'QUALITY'
price_mom = 'PRICE MOMENTUM'
safety = 'SAFETY'
biz_mom = 'BIZ MOMENTUM'

value_cols = "A:N"
quality_cols = "A:L"
price_cols = "A:L"
safety_cols = "A:J"
biz_cols = "A:G"
                            

# ---------- Helpers ----------
@st.cache_data
def load_excel_data(file_name: str,
                    sheet: str,
                    use_columns: str,
                    header_row: int,
                    nrows: Optional[int] = None) -> pd.DataFrame:
    """Loads Excel range and returns a DataFrame. nrows=None => all rows (dynamic)."""
    return pd.read_excel(
        file_name,
        sheet_name=sheet,
        usecols=use_columns,
        header=header_row,
        nrows=nrows
    )

def coerce_percent(col: pd.Series) -> pd.Series:
    """Accepts values like 0.099 or '9.9%' and returns float (0–1)."""
    if pd.api.types.is_numeric_dtype(col):
        # already fraction like 0.099
        return col
    # strings such as '9.9%' or ' 9.9 %'
    cleaned = (
        col.astype(str)
           .str.strip()
           .str.replace('%', '', regex=False)
           .str.replace(',', '', regex=False)
    )
    return pd.to_numeric(cleaned, errors='coerce') / 100.0


# ---------- Load & Clean ----------
df = load_excel_data(excel_file, value, value_cols, 1, nrows=None)

# Ensure exact column order (rename if your sheet uses spaces/variants)
#expected = ["ETF", "COUNTRY", "CATEGORY", "CURRENT_RETURNS", "MEAN", "STD_DEV", "2_SIGMA", "CURRENT_PRICE", "200_DMA"]
#df.columns = expected  # if your headers already match, this is a no-op

# Coerce percentage columns
				

pct_cols = [ "EDITDA/EV", "GROSS PROFIT/EV", "EPS TRAILING RANK", "EPS FORWARD RANK", "FCF YIELD RANK", "EDITDA/EV RANK", "PROFIT/EV RANK", "VALUE RANK"]
for c in pct_cols:
    if c in df.columns:
        df[c] = coerce_percent(df[c])

# ---------- Display (styled like the screenshot) ----------
# Use pandas Styler for bold ETF and percentage formatting
styler = (
    df.style
      .format({
          **{c: "{:.1%}" for c in pct_cols},
          "EPS TRAILING": "{:.1f}",
          "EPS FORWARD": "{:.1f}",
          "FCF YIELD" : "{:.1f}"
      }, na_rep="-")
      .hide(axis="index")
)

st.subheader("VALUE MODEL")

st.dataframe(styler, hide_index = True)


