import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
import os

# Ignora avisos do pandas
warnings.filterwarnings('ignore')

st.title("Processador de Ausências")
st.write("Este aplicativo processa dados de ausências de funcionários.")

def converter_data_br_para_datetime(data_str):
    """
    Converte uma string de data no formato brasileiro (DD/MM/YYYY) para um objeto datetime.
    Retorna None se a conversão falhar.
    """
    if pd.isna(data_str) or data_str == '':
        return None
    
    if isinstance(data_str, datetime):
        return data_str
        
    try:
        # Se for uma string, tenta converter
        if isinstance(data_str, str):
            # Tenta converter do formato brasileiro DD/MM/YYYY
            return datetime.strptime(data_str, '%d/%m/%Y')
        else:
            return data_str  # Se já for outro tipo (como timestamp), retorna como está
    except (ValueError, TypeError):
        try:
            # Caso falhe, tenta outros formatos comuns
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y']:
                try:
                    return datetime.strptime(str(data_str), fmt)
                except ValueError:
                    continue
            return None
        except Exception:
            return None

def carregar_arquivo_ausencias(uploaded_file):
    """
    Carrega o arquivo de ausências e converte colunas de data corretamente.
    """
    try:
        # Lê o arquivo Excel
        df = pd.read_excel(uploaded_file)
        
        # Converte as colunas de data
        if 'Dia' in df.columns:
            df['Dia'] = df['Dia'].astype(str)
            df['Dia'] = df['Dia'].apply(converter_data_br_para_datetime)
        
        if 'Data de Demissão' in df.columns:
            df['Data de Demissão'] = df['Data de Demissão'].astype(str)
            df['Data de Demissão'] = df['Data de Demissão'].apply(converter_data_br_para_datetime)
        
        return df
    except Exception as e:
        st.error(f"Erro ao carregar arquivo de ausências: {e}")
        return pd.DataFrame()

def carregar_arquivo_funcionarios(uploaded_file):
    """
    Carrega o arquivo de funcionários e converte colunas de data corretamente.
    """
    try:
        # Lê o arquivo Excel
        df = pd.read_excel(uploaded_file)
        
        # Converte as colunas de data
        if 'Data Término Contrato' in df.columns:
            df['Data Término Contrato'] = df['Data Término Contrato'].astype(str)
            df['Data Término Contrato'] = df['Data Término Contrato'].apply(converter_data_br_para_datetime)
        
        if 'Data Admissão' in df.columns:
            df['Data Admissão'] = df['Data Admissão'].astype(str)
            df['Data Admissão'] = df['Data Admissão'].apply(converter_data_br_para_datetime)
        
        return df
    except Exception as e:
        st.error(f"Erro ao carregar arquivo de funcionários: {e}")
        return pd.DataFrame()

def filtrar_ausencias_por_periodo(df_ausencias, data_inicio, data_fim):
    """
    Filtra ausências para um determinado período.
    """
    # Converte data_inicio e data_fim para datetime se forem strings
    if isinstance(data_inicio, str):
        data_inicio = converter_data_br_para_datetime(data_inicio)
    if isinstance(data_fim, str):
        data_fim = converter_data_br_para_datetime(data_fim)
    
    # Verifica se a conversão ocorreu com sucesso
    if data_inicio is None or data_fim is None:
        st.error("Erro: formato de data inválido.")
        return pd.DataFrame()
    
    # Filtra as ausências dentro do período
    try:
        ausencias_periodo = df_ausencias[
            (df_ausencias['Dia'] >= data_inicio) & 
            (df_ausencias['Dia'] <= data_fim)
        ]
        return ausencias_periodo
    except Exception as e:
        st.error(f"Erro ao filtrar ausências por período: {e}")
        return pd.DataFrame()

def processar_dados(df_ausencias, df_funcionarios, mes=None, ano=None):
    """
    Processa os dados de ambos os arquivos com tratamento correto de datas.
    """
    if df_ausencias.empty or df_funcionarios.empty:
        st.error("Erro: Um ou mais arquivos não puderam ser carregados corretamente.")
        return pd.DataFrame()
    
    # Se mês e ano forem fornecidos, filtra para esse período
    if mes is not None and ano is not None:
        # Determine o primeiro e último dia do mês
        primeiro_dia = datetime(ano, mes, 1)
        if mes == 12:
            ultimo_dia = datetime(ano + 1, 1, 1) - timedelta(days=1)
        else:
            ultimo_dia = datetime(ano, mes + 1, 1) - timedelta(days=1)
        
        # Filtra as ausências para o mês/ano especificado
        df_ausencias = filtrar_ausencias_por_periodo(df_ausencias, primeiro_dia, ultimo_dia)
    
    # Mescla os dados de ausências com os dados de funcionários
    # Usando a "Matricula" como chave de junção
    if 'Matricula' in df_ausencias.columns and 'Matrícula' in df_funcionarios.columns:
        # Converte a coluna Matrícula para o mesmo tipo em ambos os DataFrames
        df_ausencias['Matricula'] = df_ausencias['Matricula'].astype(str)
        df_funcionarios['Matrícula'] = df_funcionarios['Matrícula'].astype(str)
        
        # Mescla os DataFrames
        df_combinado = pd.merge(
            df_ausencias,
            df_funcionarios,
            left_on='Matricula',
            right_on='Matrícula',
            how='left'
        )
        
        return df_combinado
    else:
        st.warning("Aviso: Colunas de matrícula não encontradas em um ou ambos os arquivos.")
        return df_ausencias  # Retorna apenas as ausências se não for possível combinar

# Interface Streamlit
st.sidebar.header("Upload de Arquivos")

# Upload dos arquivos Excel
arquivo_ausencias = st.sidebar.file_uploader("Arquivo de Ausências", type=["xlsx", "xls"])
arquivo_funcionarios = st.sidebar.file_uploader("Arquivo de Funcionários", type=["xlsx", "xls"])

# Seleção de período
st.sidebar.header("Filtrar por Período")
meses = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 
         6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 
         10: "Outubro", 11: "Novembro", 12: "Dezembro"}

mes_selecionado = st.sidebar.selectbox("Mês", options=list(meses.keys()), format_func=lambda x: meses[x])
ano_selecionado = st.sidebar.number_input("Ano", min_value=2020, max_value=2030, value=2025)

# Botão para processar
processar = st.sidebar.button("Processar Dados")

# Processamento
if processar:
    if arquivo_ausencias is not None and arquivo_funcionarios is not None:
        with st.spinner("Processando arquivos..."):
            # Carrega os dados
            df_ausencias = carregar_arquivo_ausencias(arquivo_ausencias)
            df_funcionarios = carregar_arquivo_funcionarios(arquivo_funcionarios)
            
            # Processa os dados
            resultado = processar_dados(df_ausencias, df_funcionarios, mes=mes_selecionado, ano=ano_selecionado)
            
            if not resultado.empty:
                st.success(f"Processamento concluído com sucesso. {len(resultado)} registros encontrados.")
                
                # Exibe os resultados
                st.subheader("Resultados")
                st.dataframe(resultado)
                
                # Botão para download
                csv = resultado.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download CSV",
                    data=csv,
                    file_name=f"resultado_ausencias_{meses[mes_selecionado]}_{ano_selecionado}.csv",
                    mime="text/csv"
                )
                
                # Botão para download do Excel
                excel_buffer = resultado.to_excel(index=False)
                st.download_button(
                    label="Download Excel",
                    data=excel_buffer,
                    file_name=f"resultado_ausencias_{meses[mes_selecionado]}_{ano_selecionado}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Nenhum resultado encontrado para o período selecionado.")
    else:
        st.warning("Por favor, faça o upload dos arquivos necessários.")

# Exibe informações de uso
st.sidebar.markdown("---")
st.sidebar.info("""
**Como usar:**
1. Faça o upload do arquivo de ausências
2. Faça o upload do arquivo de funcionários
3. Selecione o mês e ano desejados
4. Clique em "Processar Dados"
5. Veja os resultados e faça o download se necessário
""")
