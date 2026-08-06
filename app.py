import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Gestão de Entregas", layout="wide")
st.title("🚚 Painel Operacional e Controle de Entregas")
st.markdown("---")

uploaded_file = st.file_uploader("Faça o upload da planilha (.csv ou .xlsx)", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        df = None
        if uploaded_file.name.endswith('.csv'):
            for enc in ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']:
                try:
                    uploaded_file.seek(0)
                    temp_df = pd.read_csv(uploaded_file, sep=';', encoding=enc)
                    if len(temp_df.columns) <= 1:
                        uploaded_file.seek(0)
                        temp_df = pd.read_csv(uploaded_file, sep=',', encoding=enc)
                    df = temp_df
                    break
                except Exception:
                    continue
            if df is None:
                st.error("Não foi possível ler o arquivo CSV. Verifique o formato do arquivo.")
        else:
            df = pd.read_excel(uploaded_file)

        if df is not None:
            # Limpa espaços em branco nos nomes das colunas
            df.columns = df.columns.str.strip()

            # Converte e formata as datas de forma flexível e segura
            df['DATA_HORA_DT'] = pd.to_datetime(df
