"""
Processador de Prêmio Assiduidade - Versão com Depuração Avançada
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

def read_excel(file, sheet_name=None, header_row=None):
    """Lê um arquivo Excel com depuração avançada"""
    try:
        add_to_log(f"Tentando ler arquivo: {file.name}", 'debug')
        
        # Se nenhum header_row for especificado, tenta múltiplas opções
        if header_row is None:
            header_options = [0, 1, 2, 3]  # Tenta diferentes linhas de cabeçalho
        else:
            header_options = [header_row]
        
        for try_header in header_options:
            try:
                # Tenta ler a primeira planilha se nenhuma for especificada
                if sheet_name is None:
                    xls = pd.ExcelFile(file)
                    sheet_options = xls.sheet_names
                else:
                    sheet_options = [sheet_name]
                
                for try_sheet in sheet_options:
                    add_to_log(f"Tentando ler planilha: {try_sheet}, header na linha: {try_header}", 'debug')
                    
                    df = pd.read_excel(
                        file, 
                        sheet_name=try_sheet, 
                        header=try_header, 
                        engine='openpyxl'
                    )
                    
                    # Remove colunas sem nome
                    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                    
                    # Se tem colunas válidas, para de tentar
                    if not df.empty and len(df.columns) > 0:
                        add_to_log(f"Arquivo '{file.name}' lido com sucesso. Planilha: {try_sheet}, Header linha: {try_header}", 'info')
                        debug_dataframe(df, file.name)
                        return df
            
            except Exception as inner_e:
                add_to_log(f"Erro ao tentar ler arquivo com header={try_header}: {str(inner_e)}", 'warning')
        
        # Se chegar aqui, nenhuma tentativa funcionou
        add_to_log(f"Falha total ao ler o arquivo {file.name}", 'error')
        return None
    
    except Exception as e:
        add_to_log(f"Erro fatal ao ler arquivo '{file.name}': {str(e)}\n{traceback.format_exc()}", 'error')
        return None

def normalize_strings(df):
    """Remove acentuação de strings com log detalhado"""
    try:
        original_columns = df.columns.tolist()
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
        
        salario = float(row.get('Salário Mês Atual', 0) or 0)
        
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
        horas = int(row.get('Qtd Horas Mensais', 0) or 0)
        
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
    """Processa os dados com log detalhado"""
    try:
        add_to_log("Iniciando processamento de dados", 'info')
        
        # Leitura dos arquivos com múltiplas tentativas
        df_base = read_excel(base_file)
        df_ausencias = read_excel(absence_file)
        df_model = read_excel(model_file)

        if df_base is None or df_ausencias is None or df_model is None:
            add_to_log("Erro: Um ou mais arquivos não foram lidos corretamente.", 'error')
            return None

        # Normalização de strings
        df_base = normalize_strings(df_base)
        df_ausencias = normalize_strings(df_ausencias)

        # Mapeamento de nomes de colunas para garantir consistência
        column_mapping = {
            'Matrícula': 'Código Funcionário',
            'Nome': 'Nome Funcionário',
            'nome': 'Nome Funcionário'
        }

        # Renomeia colunas se necessário
        for old_col, new_col in column_mapping.items():
            if old_col in df_base.columns:
                df_base.rename(columns={old_col: new_col}, inplace=True)
            if old_col in df_ausencias.columns:
                df_ausencias.rename(columns={old_col: new_col}, inplace=True)

        # Processa faltas
        df_ausencias = process_faltas(df_ausencias)

        # Merge dos dataframes
        df_merge = pd.merge(
            df_base, 
            df_ausencias[['Código Funcionário', 'Falta', 'Afastamentos', 'Ausência Integral', 'Ausência Parcial']],
            on='Código Funcionário', 
            how='left'
        )

        # Calcular prêmio por funcionário
        resultados = []
        for _, row in df_merge.iterrows():
            status, valor = calcular_premio(row)
            row['Status Prêmio'] = status
            row['Valor Prêmio'] = valor
            resultados.append(row)

        df_resultado = pd.DataFrame(resultados)

        # Garantir que a estrutura do modelo seja mantida
        for col in df_model.columns:
            if col not in df_resultado.columns:
                df_resultado[col] = None

        df_resultado = df_resultado[df_model.columns]

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
