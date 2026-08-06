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
        if uploaded_file.name.endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file, sep=';')
                if len(df.columns) <= 1:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, sep=',')
            except Exception:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep=',')
        else:
            df = pd.read_excel(uploaded_file)

        df['DATA_HORA_DT'] = pd.to_datetime(df['DATA_HORA_APROXIMADA'], format='%d/%m/%Y %H:%M', errors='coerce')
        df['DATA'] = df['DATA_HORA_DT'].dt.date
        df['HORA'] = df['DATA_HORA_DT'].dt.strftime('%H:%M')

        st.sidebar.header("🎯 Filtros")

        bases_disponiveis = sorted(df['NOM_BASE_OPERACIONAL'].dropna().unique())
        filtro_base = st.sidebar.multiselect("Base Operacional", options=bases_disponiveis, default=bases_disponiveis)

        municipios_disponiveis = sorted(df['NOM_MUNICIPIO'].dropna().unique())
        filtro_municipio = st.sidebar.multiselect("Município", options=municipios_disponiveis, default=municipios_disponiveis)

        agentes_disponiveis = sorted(df['COD_AGENTE_COMERCIAL'].dropna().unique())
        filtro_agente = st.sidebar.multiselect("Código do Agente", options=agentes_disponiveis, default=agentes_disponiveis)

        datas_disponiveis = sorted(df['DATA'].dropna().unique())
        filtro_data = st.sidebar.multiselect("Data da Entrega", options=datas_disponiveis, default=datas_disponiveis) if datas_disponiveis else []

        df_filtrado = df[
            (df['NOM_BASE_OPERACIONAL'].isin(filtro_base)) &
            (df['NOM_MUNICIPIO'].isin(filtro_municipio)) &
            (df['COD_AGENTE_COMERCIAL'].isin(filtro_agente)) &
            (df['DATA'].isin(filtro_data))
        ]

        if df_filtrado.empty:
            st.warning("Nenhum dado encontrado com os filtros selecionados.")
        else:
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total de Entregas", len(df_filtrado))
            col2.metric("Total de Agentes Ativos", df_filtrado['COD_AGENTE_COMERCIAL'].nunique())
            col3.metric("Bases Atendidas", df_filtrado['NOM_BASE_OPERACIONAL'].nunique())
            col4.metric("Cidades Atendidas", df_filtrado['NOM_MUNICIPIO'].nunique())

            st.markdown("---")

            def min_hora(series):
                valid = series.dropna()
                return valid.min().strftime('%H:%M') if not valid.empty else "N/A"

            def max_hora(series):
                valid = series.dropna()
                return valid.max().strftime('%H:%M') if not valid.empty else "N/A"

            def lista_cidades(series):
                cidades = series.dropna().unique()
                return ", ".join(cidades)

            df_resumo = df_filtrado.groupby(['DATA', 'COD_AGENTE_COMERCIAL', 'NOM_BASE_OPERACIONAL']).agg(
                TOTAL_ENTREGAS=('SEQ_TAREFA', 'count'),
                HORARIO_INICIAL=('DATA_HORA_DT', min_hora),
                HORARIO_FINAL=('DATA_HORA_DT', max_hora),
                CIDADES=('NOM_MUNICIPIO', lista_cidades)
            ).reset_index()

            df_resumo.columns = [
                'Data', 'Código Agente', 'Base Operacional', 
                'Total de Entregas', 'Horário Inicial (1ª)', 
                'Horário Final (Última)', 'Cidades Atendidas'
            ]

            col_graf1, col_graf2 = st.columns(2)

            with col_graf1:
                st.subheader("🏙️ Entregas por Cidade")
                df_cidade = df_filtrado.groupby('NOM_MUNICIPIO').size().reset_index(name='Qtd Entregas')
                fig_cidade = px.bar(df_cidade, x='NOM_MUNICIPIO', y='Qtd Entregas', text_auto=True, color_discrete_sequence=['#1f77b4'])
                st.plotly_chart(fig_cidade, use_container_width=True)

            with col_graf2:
                st.subheader("🏢 Entregas por Base Operacional")
                df_base_graf = df_filtrado.groupby('NOM_BASE_OPERACIONAL').size().reset_index(name='Qtd Entregas')
                fig_base = px.bar(df_base_graf, x='NOM_BASE_OPERACIONAL', y='Qtd Entregas', text_auto=True, color_discrete_sequence=['#2ca02c'])
                st.plotly_chart(fig_base, use_container_width=True)

            st.markdown("---")
            st.subheader("📋 Resumo Diário por Agente")
            st.dataframe(df_resumo, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_resumo.to_excel(writer, index=False, sheet_name='Resumo Entregas')

            st.download_button(
                label="📥 Baixar Planilha Tratada (Excel)",
                data=buffer.getvalue(),
                file_name="resumo_entregas_tratado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
else:
    st.info("👆 Por favor, envie o arquivo de entregas para iniciar.")
