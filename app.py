import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import statistics
import io

st.set_page_config(page_title="Vandføring fra Hipdata.dk", layout="centered")

st.title("Vandføring fra Hipdata.dk")
st.write("Upload CSV fra hipdata.dk og få konverteret Excel med beregninger.")

uploaded_file = st.file_uploader("Vælg CSV-fil", type="csv")

if uploaded_file:
    try:
        try:
            df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8-sig')
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=';', encoding='latin-1')

        if df.empty:
            st.error("CSV-filen er tom.")
        else:
            # Konverter kolonne B til tal med korrekt dtype (brug kolonnenavn, ikke iloc)
            col_flow = df.columns[1]
            df[col_flow] = pd.to_numeric(df[col_flow].astype(str).str.replace(',', '.'), errors='coerce')

            # Indsæt MONTH-kolonne
            df['Month'] = pd.to_datetime(df.iloc[:, 0], dayfirst=True, errors='coerce').dt.month

            # Beregn middel for maj-sept
            maj_sept_values = df.loc[df['Month'].between(5, 9), col_flow]
            if not maj_sept_values.empty:
                average = float(round(maj_sept_values.mean() * 1000, 6))
            else:
                average = "Ingen data"

            # Beregn median af maksimum for vintermåneder
            vintermåneder = [1, 2, 3, 4, 10, 11, 12]
            max_values = []
            for m in vintermåneder:
                m_values = df.loc[df['Month'] == m, col_flow]
                if not m_values.empty:
                    max_values.append(float(m_values.max()))

            if max_values:
                median_max = float(round(statistics.median(max_values) * 1000, 6))
            else:
                median_max = "Ingen data"

            # Skriv til Excel uden hjælpekolonne (Month)
            df_excel = df.drop(columns=['Month'])
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_excel.to_excel(writer, index=False, sheet_name='Sheet1')
                worksheet = writer.sheets['Sheet1']

                # Skriv middel og median i Excel
                worksheet.cell(row=1, column=4).value = 'Middel (maj–sept)'
                worksheet.cell(row=2, column=4).value = average if isinstance(average, float) else average

                worksheet.cell(row=1, column=5).value = 'Median af max (vintermåneder)'
                worksheet.cell(row=2, column=5).value = median_max if isinstance(median_max, float) else median_max

            output.seek(0)

            # Vis resultater i Streamlit
            st.success(f"Middel (maj–sept): {average} L/s" if isinstance(average, float) else "Ingen data til beregning")
            st.success(f"Median maksimum (vintermåneder): {median_max} L/s" if isinstance(median_max, float) else "Ingen data til median af max")

            st.download_button(
                label="Download Excel-fil",
                data=output,
                file_name="konverteret.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Noget gik galt: {e}")
