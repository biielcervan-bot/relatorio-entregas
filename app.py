import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Gestão de Entregas", layout="wide")
st.title("🚚 Painel Operacional - Controle de Entregas e Prazos")
st.markdown("---")

# Upload de múltiplos arquivos
uploaded_files = st.file_uploader(
    "Faça o upload das planilhas (.csv ou .xlsx)", 
    type=["csv", "xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    try:
        lista_dfs = []
        
        for file in uploaded_files:
            df_temp = None
            if file.name.endswith('.csv'):
                for enc in ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']:
                    try:
                        file.seek(0)
                        temp = pd.read_csv(file, sep=';', encoding=enc)
                        if len(temp.columns) <= 1:
                            file.seek(0)
                            temp = pd.read_csv(file, sep=',', encoding=enc)
                        df_temp = temp
                        break
                    except Exception:
                        continue
            else:
                df_temp = pd.read_excel(file)

            if df_temp is not None:
                lista_dfs.append(df_temp)

        if lista_dfs:
            df = pd.concat(lista_dfs, ignore_index=True)
            df.columns = df.columns.str.strip()

            # 1. Tratamento da Data/Hora Real da Entrega
            df['DATA_HORA_DT'] = pd.to_datetime(df['DATA_HORA_APROXIMADA'], dayfirst=True, errors='coerce')
            df['DATA_REAL'] = df['DATA_HORA_DT'].dt.strftime('%d/%m/%Y').fillna('Sem Data')
            df['HORA'] = df['DATA_HORA_DT'].dt.strftime('%H:%M').fillna('N/A')

            # 2. Tratamento da Data Prevista da Entrega
            if 'DAT_PREVISTA_ENTREGA' in df.columns:
                df['DAT_PREVISTA_DT'] = pd.to_datetime(df['DAT_PREVISTA_ENTREGA'], dayfirst=True, errors='coerce')
                df['DATA_PREVISTA'] = df['DAT_PREVISTA_DT'].dt.strftime('%d/%m/%Y').fillna('Sem Data Prevista')
            else:
                df['DAT_PREVISTA_DT'] = pd.NaT
                df['DATA_PREVISTA'] = 'Não Informada'

            # 3. Cálculo da Diferença em Dias (Data Real - Data Prevista)
            df['DIFERENCA_DIAS'] = (df['DATA_HORA_DT'].dt.floor('D') - df['DAT_PREVISTA_DT'].dt.floor('D')).dt.days

            # Padronização de texto para os filtros
            df['NOM_BASE_OPERACIONAL_STR'] = df['NOM_BASE_OPERACIONAL'].astype(str).str.strip()
            df['NOM_MUNICIPIO_STR'] = df['NOM_MUNICIPIO'].astype(str).str.strip()
            df['NOM_UNIDADE_LEITURA_STR'] = df['NOM_UNIDADE_LEITURA'].astype(str).str.strip()
            df['COD_AGENTE_COMERCIAL_STR'] = df['COD_AGENTE_COMERCIAL'].astype(str).str.strip()

            # Barra Lateral - Filtros
            st.sidebar.header("🎯 Filtros")

            bases_disponiveis = sorted([x for x in df['NOM_BASE_OPERACIONAL_STR'].unique() if x and x != 'nan'])
            filtro_base = st.sidebar.multiselect("Base Operacional", options=bases_disponiveis, default=bases_disponiveis)

            municipios_disponiveis = sorted([x for x in df['NOM_MUNICIPIO_STR'].unique() if x and x != 'nan'])
            filtro_municipio = st.sidebar.multiselect("Município", options=municipios_disponiveis, default=municipios_disponiveis)

            unidades_disponiveis = sorted([x for x in df['NOM_UNIDADE_LEITURA_STR'].unique() if x and x != 'nan'])
            filtro_unidade = st.sidebar.multiselect("Unidade de Leitura / Lote", options=unidades_disponiveis, default=unidades_disponiveis)

            agentes_disponiveis = sorted([x for x in df['COD_AGENTE_COMERCIAL_STR'].unique() if x and x != 'nan'])
            filtro_agente = st.sidebar.multiselect("Código do Agente", options=agentes_disponiveis, default=agentes_disponiveis)

            datas_disponiveis = sorted([x for x in df['DATA_REAL'].unique() if x and x != 'nan'])
            filtro_data_real = st.sidebar.multiselect("Data Real da Entrega", options=datas_disponiveis, default=datas_disponiveis)

            datas_previstas_disp = sorted([x for x in df['DATA_PREVISTA'].unique() if x and x != 'nan'])
            filtro_data_prevista = st.sidebar.multiselect("Data Prevista", options=datas_previstas_disp, default=datas_previstas_disp)

            # Aplicação dos Filtros
            df_filtrado = df[
                (df['NOM_BASE_OPERACIONAL_STR'].isin(filtro_base)) &
                (df['NOM_MUNICIPIO_STR'].isin(filtro_municipio)) &
                (df['NOM_UNIDADE_LEITURA_STR'].isin(filtro_unidade)) &
                (df['COD_AGENTE_COMERCIAL_STR'].isin(filtro_agente)) &
                (df['DATA_REAL'].isin(filtro_data_real)) &
                (df['DATA_PREVISTA'].isin(filtro_data_prevista))
            ]

            if df_filtrado.empty:
                st.warning("Nenhum dado encontrado com os filtros selecionados.")
            else:
                # Indicadores Principais
                col1, col2, col3, col4, col5 = st.columns(5)
                col1.metric("Total de Entregas", len(df_filtrado))
                col2.metric("Agentes Ativos", df_filtrado['COD_AGENTE_COMERCIAL_STR'].nunique())
                col3.metric("Bases Atendidas", df_filtrado['NOM_BASE_OPERACIONAL_STR'].nunique())
                col4.metric("Cidades Atendidas", df_filtrado['NOM_MUNICIPIO_STR'].nunique())
                col5.metric("Lotes / ULs", df_filtrado['NOM_UNIDADE_LEITURA_STR'].nunique())

                st.markdown("---")

                # Funções para pegar a 1ª e última hora por lote
                def min_hora(series):
                    valid = series.dropna()
                    return valid.min().strftime('%H:%M') if not valid.empty else "N/A"

                def max_hora(series):
                    valid = series.dropna()
                    return valid.max().strftime('%H:%M') if not valid.empty else "N/A"

                def dif_dias_func(series):
                    valid = series.dropna()
                    if not valid.empty:
                        val = int(valid.iloc[0])
                        return f"+{val} dia(s)" if val > 0 else (f"{val} dia(s)" if val < 0 else "No prazo (0d)")
                    return "N/A"

                # Agrupamento detalhado incluindo DATA PREVISTA e DIFERENÇA EM DIAS
                df_resumo = df_filtrado.groupby([
                    'DATA_REAL',
                    'DATA_PREVISTA',
                    'COD_AGENTE_COMERCIAL_STR', 
                    'NOM_BASE_OPERACIONAL_STR', 
                    'NOM_MUNICIPIO_STR',
                    'NOM_UNIDADE_LEITURA_STR'
                ]).agg(
                    TOTAL_ENTREGAS=('SEQ_TAREFA', 'count'),
                    HORARIO_INICIAL=('DATA_HORA_DT', min_hora),
                    HORARIO_FINAL=('DATA_HORA_DT', max_hora),
                    DIFERENCA_DIAS=('DIFERENCA_DIAS', dif_dias_func)
                ).reset_index()

                df_resumo.columns = [
                    'Data Realização', 
                    'Data Prevista',
                    'Código Agente', 
                    'Base Operacional', 
                    'Cidade', 
                    'Unidade de Leitura',
                    'Total de Entregas', 
                    'Horário Inicial (1ª)', 
                    'Horário Final (Última)',
                    'Diferença (Dias)'
                ]

                # Gráficos Visualizadores
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
                st.subheader("📋 Resumo Diário por Lote e Status de Entrega")
                st.dataframe(df_resumo, use_container_width=True)

                # Download em Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_resumo.to_excel(writer, index=False, sheet_name='Resumo Prazos Lotes')

                st.download_button(
                    label="📥 Baixar Planilha Tratada com Prazos (Excel)",
                    data=buffer.getvalue(),
                    file_name="resumo_entregas_prazos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar os arquivos: {e}")
else:
    st.info("👆 Por favor, envie uma ou mais planilhas de entregas para iniciar.")
