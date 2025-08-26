import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="CSV til Excel konverter", page_icon="📊")

st.title("📊 CSV til Excel konverter")

uploaded_file = st.file_uploader("Upload en CSV fil", type=["csv"])

if uploaded_file is not None:
    # Læs CSV til pandas DataFrame
    df = pd.read_csv(uploaded_file)

    st.subheader("Eksempel på data")
    st.dataframe(df.head())

    # Konverter til Excel i hukommelsen
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
    excel_data = output.getvalue()

    # Download knap
    st.download_button(
        label="📥 Download som Excel",
        data=excel_data,
        file_name="converted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
