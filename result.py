import pandas as pd
import streamlit as st

st.title("Upload Excel File")

uploaded_file = st.file_uploader(
    "Choose an Excel file", type=["xlsx", "xls"]
)

if uploaded_file is not None:
  # Read Excel file
  df = pd.read_excel(uploaded_file)

  st.write("Here is your data:")
  st.dataframe(df)
else:
  st.warning("Please upload an Excel file to get started.")
