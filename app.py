import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Gestão de Entregas", layout="wide")
st.title("🚚 Painel Operacional e Controle de Entregas por Lote")
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

            # Converte e formata as datas de forma flexível
            df['DATA_HORA_DT'] = pd.to_datetime(df['DATA_HORA_APROXIMADA'], dayfirst=True, errors='coerce')
            df['DATA'] = df['DATA_HORA_DT'].dt.strftime('%d/%m/%Y').fillna('Sem Data')
            df['HORA'] = df['DATA_HORA_DT'].dt.strftime('%H:%M').fillna('N/A')

            # Padroniza texto dos campos de filtro
            df['NOM_BASE_OPERACIONAL_STR'] = df['NOM_BASE_OPERACIONAL'].astype(str).str.strip()
            df['NOM_MUNICIPIO_STR'] = df['NOM_MUNICIPIO'].astype(str).str.strip()
            df['NOM_UNIDADE_LEITURA_STR'] = df['NOM_UNIDADE_LEITURA'].astype(str).str.strip()
            df['COD_AGENTE_COMERCIAL_STR'] = df['COD_AGENTE_COMERCIAL'].astype(str).str.strip()

            st.sidebar.header("🎯 Filtros")

            bases_disponiveis = sorted([x for x in df['NOM_BASE_OPERACIONAL_STR'].unique() if x and x != 'nan'])
            filtro_base = st.sidebar.multiselect("Base Operacional", options=bases_disponiveis, default=bases_disponiveis)

            municipios_disponiveis = sorted([x for x in df['NOM_MUNICIPIO_STR'].unique() if x and x != 'nan'])
            filtro_municipio = st.sidebar.multiselect("Município", options=municipios_disponiveis, default=municipios_disponiveis)

            unidades_disponiveis = sorted([x for x in df['NOM_UNIDADE_LEITURA_STR'].unique() if x and x != 'nan'])
            filtro_unidade = st.sidebar.multiselect("Unidade de Leitura / Lote", options=unidades_disponiveis, default=unidades_disponiveis)

            agentes_disponiveis = sorted([x for x in df['COD_AGENTE_COMERCIAL_STR'].unique() if x and x != 'nan'])
            filtro_agente = st.sidebar.multiselect("Código do Agente", options=agentes_disponiveis, default=agentes_disponiveis)

            datas_disponiveis = sorted([x for x in df['DATA'].unique() if x and x != 'nan'])
            filtro_data = st.sidebar.multiselect("Data da Entrega", options=datas_disponiveis, default=datas_disponiveis)

            # Aplicação dos Filtros
            df_filtrado = df[
                (df['NOM_BASE_OPERACIONAL_STR'].isin(filtro_base)) &
                (df['NOM_MUNICIPIO_STR'].isin(filtro_municipio)) &
                (df['NOM_UNIDADE_LEITURA_STR'].isin(filtro_unidade)) &
                (df['COD_AGENTE_COMERCIAL_STR'].isin(filtro_agente)) &
                (df['DATA'].isin(filtro_data))
            ]

            if df_filtrado.empty:
                st.warning("Nenhum dado encontrado com os filtros selecionados.")
            else:
                # Indicadores
                col1, col2, col3, col4, col5 = st.columns(5)
                col1.metric("Total de Entregas", len(df_filtrado))
                col2.metric("Agentes Ativos", df_filtrado['COD_AGENTE_COMERCIAL_STR'].nunique())
                col3.metric("Bases Atendidas", df_filtrado['NOM_BASE_OPERACIONAL_STR'].nunique())
                col4.metric("Cidades Atendidas", df_filtrado['NOM_MUNICIPIO_STR'].nunique())
                col5.metric("Unid. Leitura / Lotes", df_filtrado['NOM_UNIDADE_LEITURA_STR'].nunique())

                st.markdown("---")

                # Funções para pegar a 1ª e última hora por agrupamento
                def min_hora(series):
                    valid = series.dropna()
                    return valid.min().strftime('%H:%M') if not valid.empty else "N/A"

                def max_hora(series):
                    valid = series.dropna()
                    return valid.max().strftime('%H:%M') if not valid.empty else "N/A"

                # Agrupamento detalhado por DIA, AGENTE, BASE, CIDADE e UNIDADE DE LEITURA
                df_resumo = df_filtrado.groupby([
                    'DATA', 
                    'COD_AGENTE_COMERCIAL_STR', 
                    'NOM_BASE_OPERACIONAL_STR', 
                    'NOM_MUNICIPIO_STR',
                    'NOM_UNIDADE_LEITURA_STR'
                ]).agg(
                    TOTAL_ENTREGAS=('SEQ_TAREFA', 'count'),
                    HORARIO_INICIAL=('DATA_HORA_DT', min_hora),
                    HORARIO_FINAL=('DATA_HORA_DT', max_hora)
                ).reset_index()

                df_resumo.columns = [
                    'Data', 
                    'Código Agente', 
                    'Base Operacional', 
                    'Cidade', 
                    'Unidade de Leitura',
                    'Total de Entregas', 
                    'Horário Inicial (1ª na UL)', 
                    'Horário Final (Última na UL)'
                ]

                # Gráficos
                col_graf1, col_graf2 = st.columns(2)

                with col_graf1:
                    st.subheader("🏙️ Entregas por Cidade")
                    df_cidade = df_filtrado.groupby('NOM_MUNICIPIO_STR').size().reset_index(name='Qtd Entregas')
                    fig_cidade = px.bar(df_cidade, x='NOM_MUNICIPIO_STR', y='Qtd Entregas', text_auto=True, labels={'NOM_MUNICIPIO_STR': 'Município'}, color_discrete_sequence=['#1f77b4'])
                    st.plotly_chart(fig_cidade, use_container_width=True)

                with col_graf2:
                    st.subheader("🏢 Entregas por Base Operacional")
                    df_base_graf = df_filtrado.groupby('NOM_BASE_OPERACIONAL_STR').size().reset_index(name='Qtd Entregas')
                    fig_base = px.bar(df_base_graf, x='NOM_BASE_OPERACIONAL_STR', y='Qtd Entregas', text_auto=True, labels={'NOM_BASE_OPERACIONAL_STR': 'Base Operacional'}, color_discrete_sequence=['#2ca02c'])
                    st.plotly_chart(fig_base, use_container_width=True)

                st.markdown("---")
                st.subheader("📋 Resumo Diário por Agente, Cidade e Unidade de Leitura")
                st.dataframe(df_resumo, use_container_width=True)

                # Download em Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_resumo.to_excel(writer, index=False, sheet_name='Resumo Lotes Entregas')

                st.download_button(
                    label="📥 Baixar Planilha Tratada (Excel)",
                    data=buffer.getvalue(),
                    file_name="resumo_entregas_lotes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
else:
    st.info("👆 Por favor, envie o arquivo de entregas para iniciar.")
