"""
Processador de Prêmio Assiduidade - Versão Corrigida
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
        header_options = [0, None]
        
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

def normalize_strings(df):
    """Remove acentuação de strings"""
    try:
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].apply(
                lambda x: unicodedata.normalize('NFKD', str(x)).encode('ascii', 'ignore').decode('utf-8') 
                if isinstance(x, str) else x
            )
        
        add_to_log("Normalização de strings concluída.", 'debug')
        return df
    except Exception as e:
        add_to_log(f"Erro ao normalizar strings: {str(e)}\n{traceback.format_exc()}", 'error')
        return df

def process_faltas(df):
    """Processa a coluna de faltas com log detalhado"""
    try:
        add_to_log("Processando coluna de faltas", 'debug')
        
        # Verifica as possíveis variações da coluna de faltas
        falta_columns = [
            'Falta', 
            'falta', 
            'Faltas', 
            'FALTA'
        ]
        
        falta_column = None
        for col in falta_columns:
            if col in df.columns:
                falta_column = col
                break
        
        if falta_column:
            df['Falta'] = df[falta_column].fillna(0).astype(str).apply(lambda x: x.count('x'))
            add_to_log(f"Coluna 'Falta' processada usando {falta_column}", 'info')
        else:
            df['Falta'] = 0
            add_to_log("Nenhuma coluna de falta encontrada. Criando coluna com valores zerados.", 'warning')
        
        return df
    except Exception as e:
        add_to_log(f"Erro ao processar faltas: {str(e)}\n{traceback.format_exc()}", 'error')
        return df

def calcular_premio(row):
    """Calcula o prêmio com log de depuração"""
    try:
        add_to_log(f"Calculando prêmio para linha: {row.to_dict()}", 'debug')
        
        # Converte valor do salário para float com tratamento de erros
        try:
            salario = float(str(row.get('Salário Mês Atual', 0)).replace(',', '.') or 0)
        except (ValueError, TypeError):
            salario = 0

        if salario > SALARIO_LIMITE:
            add_to_log(f"Salário {salario} maior que limite. Não pagar.", 'warning')
            return 'NÃO PAGAR - SALÁRIO MAIOR', 0

        # Verifica condições de não pagamento
        condicoes_nao_pagar = [
            ('Falta', row.get('Falta', 0) > 0, 'NÃO PAGAR - FALTA'),
            ('Afastamentos', row.get('Afastamentos', None), 'NÃO PAGAR - AFASTAMENTO'),
            ('Ausência Integral', row.get('Ausência Integral', None), 'NÃO PAGAR - AUSÊNCIA INTEGRAL'),
            ('Ausência Parcial', row.get('Ausência Parcial', None), 'NÃO PAGAR - AUSÊNCIA PARCIAL')
        ]

        for condicao, valor, mensagem in condicoes_nao_pagar:
            if valor:
                add_to_log(f"{condicao} detectada: {valor}. {mensagem}", 'warning')
                return mensagem, 0

        # Calcula horas trabalhadas
        try:
            horas = int(str(row.get('Qtd Horas Mensais', 0)).replace(',', '.') or 0)
        except (ValueError, TypeError):
            horas = 0
        
        if horas == 220:
            add_to_log("Horas completas. Pagamento integral.", 'info')
            return 'PAGAR', PREMIO_VALOR_INTEGRAL
        elif horas in [110, 120]:
            add_to_log("Horas parciais. Pagamento parcial.", 'info')
            return 'PAGAR', PREMIO_VALOR_PARCIAL

        add_to_log(f"Horas inesperadas: {horas}. Verificar.", 'warning')
        return 'PAGAR - SEM OCORRÊNCIAS', PREMIO_VALOR_INTEGRAL

    except Exception as e:
        add_to_log(f"Erro no cálculo do prêmio: {str(e)}\n{traceback.format_exc()}", 'error')
        return 'ERRO - VERIFICAR DADOS', 0

def process_data(base_file, absence_file, model_file):
    """Processa os dados com estratégias robustas"""
    try:
        add_to_log("Iniciando processamento de dados", 'info')
        
        # Leitura dos arquivos com múltiplas tentativas
        df_base = read_excel(base_file)
        df_ausencias = read_excel(absence_file)
        df_model = read_excel(model_file)

        if df_base is None or df_ausencias is None or df_model is None:
            add_to_log("Erro: Um ou mais arquivos não foram lidos corretamente.", 'error')
            return None

        # Mapeamento flexível de nomes de colunas
        column_mapping = {
            'Matrícula': ['Matrícula', 'Matricula', 'Codigo', 'Código Funcionário'],
            'Nome': ['Nome', 'Nome Funcionário', 'nome'],
            'Código Funcionário': ['Código Funcionário', 'Codigo Funcionario', 'Matrícula']
        }

        # Função para encontrar a primeira coluna correspondente
        def find_column(df, possible_names):
            for name in possible_names:
                if name in df.columns:
                    return name
            return None

        # Renomeia colunas de forma flexível
        base_cols = {
            'Código Funcionário': find_column(df_base, column_mapping['Código Funcionário']),
            'Nome Funcionário': find_column(df_base, column_mapping['Nome'])
        }
        
        ausencias_cols = {
            'Código Funcionário': find_column(df_ausencias, column_mapping['Código Funcionário']),
            'Nome Funcionário': find_column(df_ausencias, column_mapping['Nome'])
        }

        # Renomeia as colunas encontradas
        for target, source in base_cols.items():
            if source and source != target:
                df_base.rename(columns={source: target}, inplace=True)
        
        for target, source in ausencias_cols.items():
            if source and source != target:
                df_ausencias.rename(columns={source: target}, inplace=True)

        # Processa faltas com estratégia mais flexível
        df_ausencias = process_faltas(df_ausencias)

        # Merge dos dataframes com estratégia flexível
        merge_columns = ['Código Funcionário', 'Nome Funcionário']
        merge_estrategias = [
            {'on': 'Código Funcionário', 'how': 'left'},
            {'on': 'Nome Funcionário', 'how': 'left'}
        ]

        df_merge = None
        for estrategia in merge_estrategias:
            try:
                df_merge = pd.merge(
                    df_base, 
                    df_ausencias[merge_columns + ['Falta', 'Afastamentos', 'Ausência Integral', 'Ausência Parcial']],
                    **estrategia
                )
                
                if not df_merge.empty:
                    break
            except Exception as e:
                add_to_log(f"Falha no merge com {estrategia}: {str(e)}", 'warning')

        if df_merge is None or df_merge.empty:
            add_to_log("Não foi possível fazer o merge dos dados", 'error')
            return None

        # Calcular prêmio por funcionário
        resultados = []
        for _, row in df_merge.iterrows():
            status, valor = calcular_premio(row)
            row['Status Prêmio'] = status
            row['Valor Prêmio'] = valor
            resultados.append(row)

        df_resultado = pd.DataFrame(resultados)

        # Garante a estrutura do modelo, mesmo com colunas diferentes
        modelo_colunas = df_model.columns.tolist()
        for col in modelo_colunas:
            if col not in df_resultado.columns:
                df_resultado[col] = None

        # Reordena de acordo com o modelo, ignorando colunas não existentes
        df_resultado = df_resultado[[col for col in modelo_colunas if col in df_resultado.columns]]

        add_to_log("Processamento concluído com sucesso.", 'info')
        return df_resultado

    except Exception as e:
        add_to_log(f"Erro fatal no processamento: {str(e)}\n{traceback.format_exc()}", 'error')
        return None

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
