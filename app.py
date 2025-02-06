"""
Processador de Prêmio Assiduidade - Versão Robusta
"""

import streamlit as st
import pandas as pd
import unicodedata
import base64
from io import BytesIO
import traceback
import logging

# Configuração inicial do Streamlit
st.set_page_config(page_title="Processador de Prêmio Assiduidade (Debug)", layout="wide")

# Configuração de logging
logging.basicConfig(level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s: %(message)s')

# Constantes
PREMIO_VALOR_INTEGRAL = 300.00
PREMIO_VALOR_PARCIAL = 150.00
SALARIO_LIMITE = 2542.86

# Lista para armazenar logs
log_messages = []

def add_to_log(message, level='info'):
    """Adiciona mensagens ao log com diferentes níveis"""
    global log_messages
    log_entry = f"{level.upper()}: {message}"
    log_messages.append(log_entry)
    
    # Usa o logger do Python para saída adicional
    if level == 'debug':
        logging.debug(message)
    elif level == 'warning':
        logging.warning(message)
    elif level == 'error':
        logging.error(message)
    else:
        logging.info(message)

def generate_log_file():
    """Gera um arquivo de log em memória"""
    log_data = "\n".join(log_messages)
    return BytesIO(log_data.encode('utf-8'))

def debug_dataframe(df, name):
    """Função para depuração detalhada de DataFrames"""
    if df is None:
        add_to_log(f"DataFrame {name} is None", 'error')
        return
    
    add_to_log(f"Debug DataFrame: {name}", 'debug')
    add_to_log(f"Shape: {df.shape}", 'debug')
    add_to_log(f"Columns: {df.columns.tolist()}", 'debug')
    add_to_log(f"Column Types:\n{df.dtypes}", 'debug')
    
    # Mostra as primeiras linhas
    add_to_log(f"First 5 rows:\n{df.head().to_string()}", 'debug')
    
    # Verifica valores nulos
    null_counts = df.isnull().sum()
    if null_counts.any():
        add_to_log(f"Null value counts:\n{null_counts}", 'warning')

def read_excel(file, sheet_name=None):
    """Lê um arquivo Excel com estratégias robustas de leitura"""
    try:
        add_to_log(f"Tentando ler arquivo: {file.name}", 'debug')
        
        # Lê todo o arquivo para inspeção
        xls = pd.ExcelFile(file)
        sheet_options = xls.sheet_names if sheet_name is None else [sheet_name]
        
        # Estratégias de leitura
        header_options = [0, 1, None]
        
        for sheet in sheet_options:
            for header in header_options:
                try:
                    # Lê o DataFrame 
                    df = pd.read_excel(
                        file, 
                        sheet_name=sheet, 
                        header=header, 
                        engine='openpyxl'
                    )
                    
                    # Remove colunas sem nome
                    if header is not None:
                        df.columns = [str(col).strip() for col in df.columns]
                    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                    
                    # Filtra linhas totalmente vazias
                    df = df.dropna(how='all')
                    
                    # Verifica se o DataFrame não está vazio
                    if not df.empty and len(df.columns) > 0:
                        add_to_log(f"Arquivo '{file.name}' lido com sucesso. Planilha: {sheet}, Header: {header}", 'info')
                        
                        # Debug detalhado
                        debug_dataframe(df, file.name)
                        
                        return df
                
                except Exception as inner_e:
                    add_to_log(f"Falha na leitura. Sheet: {sheet}, Header: {header}. Erro: {str(inner_e)}", 'debug')
        
        # Se chegar aqui, nenhuma estratégia funcionou
        add_to_log(f"Falha total ao ler o arquivo {file.name}", 'error')
        return None
    
    except Exception as e:
        add_to_log(f"Erro fatal ao ler arquivo '{file.name}': {str(e)}\n{traceback.format_exc()}", 'error')
        return None

# [Restante do código anterior permanece o mesmo]

def main():
    st.title("Processador de Prêmio Assiduidade (Modo Debug)")
    
    base_file = st.file_uploader("Arquivo Base", type=['xlsx'], key='base')
    absence_file = st.file_uploader("Arquivo de Ausências", type=['xlsx'], key='ausencia')
    model_file = st.file_uploader("Modelo de Exportação", type=['xlsx'], key='modelo')

    # Área de debug para mostrar informações detalhadas
    debug_area = st.expander("Detalhes de Depuração")

    if base_file and absence_file and model_file:
        if st.button("Processar Dados"):
            with st.spinner('Processando dados...'):
                df_resultado = process_data(base_file, absence_file, model_file)
                
                if df_resultado is not None:
                    st.success("Dados processados com sucesso!")
                    
                    # Área de debug para mostrar informações detalhadas
                    with debug_area:
                        st.write("Registros processados:", len(df_resultado))
                        st.dataframe(df_resultado.head())
                    
                    # Exportação do resultado
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_resultado.to_excel(writer, index=False, sheet_name='Resultado')
                    output.seek(0)

                    st.download_button("Baixar Relatório", data=output, file_name="resultado_assiduidade.xlsx")
                else:
                    st.error("Falha no processamento. Verifique os logs.")

            # Sempre mostra o log
            st.download_button("Baixar Log Detalhado", data=generate_log_file(), file_name="log_processamento_debug.txt")
            
            # Mostra os logs na interface
            with st.expander("Logs de Processamento"):
                for log in log_messages:
                    st.text(log)

if __name__ == "__main__":
    main()
