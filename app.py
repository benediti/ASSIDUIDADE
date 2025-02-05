"""
Processador de Prêmio Assiduidade - Correção Final
"""

import streamlit as st
import pandas as pd
import unicodedata
import base64
from io import BytesIO

# Configurações iniciais
st.set_page_config(page_title="Processador de Prêmio Assiduidade", layout="wide")

# Constantes
PREMIO_VALOR_INTEGRAL = 300.00
PREMIO_VALOR_PARCIAL = 150.00
SALARIO_LIMITE = 2542.86

# Variável global para armazenar logs
log_messages = []

def add_to_log(message):
    """Adiciona mensagens ao log"""
    global log_messages
    log_messages.append(message)

def generate_log_file():
    """Gera um arquivo de log em memória"""
    log_data = "\n".join(log_messages)
    log_file = BytesIO(log_data.encode('utf-8'))
    return log_file

def read_excel(file, sheet_name=None):
    """Lê o arquivo Excel e retorna o DataFrame com o cabeçalho especificado"""
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, engine='openpyxl')
        add_to_log(f"Arquivo '{file.name}' lido com sucesso. Colunas detectadas: {df.columns.tolist()}")
        return df
    except Exception as e:
        add_to_log(f"Erro ao ler arquivo '{file.name}': {str(e)}")
        return None

def normalize_strings(df):
    """Remove acentuação de strings nas colunas do DataFrame"""
    try:
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].apply(lambda x: unicodedata.normalize('NFKD', str(x)).encode('ascii', 'ignore').decode('utf-8') if isinstance(x, str) else x)
        return df
    except Exception as e:
        add_to_log(f"Erro ao normalizar strings: {str(e)}")
        return df

def process_faltas(df):
    """Converte 'x' na coluna 'Falta' para o valor 1, e preenche valores em branco com 0"""
    try:
        if 'Falta' in df.columns:
            df['Falta'] = df['Falta'].fillna(0).astype(str).apply(lambda x: x.count('x'))
            add_to_log("Coluna 'Falta' processada: valores 'x' convertidos para 1 e vazios para 0.")
        else:
            df['Falta'] = 0
            add_to_log("Coluna 'Falta' não encontrada no arquivo. Criando a coluna com valores zerados.")
        return df
    except Exception as e:
        add_to_log(f"Erro ao processar faltas: {str(e)}")
        return df

def calcular_premio(row):
    """Calcula o prêmio com base nas regras de assiduidade"""
    try:
        salario = float(row.get('Salário Mês Atual', 0) or 0)
        if salario > SALARIO_LIMITE:
            return 'NÃO PAGAR - SALÁRIO MAIOR', 0

        if row.get('Falta', 0) > 0 or row.get('Afastamentos', None) or row.get('Ausência Integral', None) or row.get('Ausência Parcial', None):
            if row.get('Falta', 0) > 0:
                return 'NÃO PAGAR - FALTA', 0
            if row.get('Afastamentos', None):
                return 'NÃO PAGAR - AFASTAMENTO', 0
            if row.get('Ausência Integral', None) or row.get('Ausência Parcial', None):
                return 'AVALIAR - AUSÊNCIA', 0

        horas = int(row.get('Qtd Horas Mensais', 0) or 0)
        if horas == 220:
            return 'PAGAR', PREMIO_VALOR_INTEGRAL
        elif horas in [110, 120]:
            return 'PAGAR', PREMIO_VALOR_PARCIAL

        return 'PAGAR - SEM OCORRÊNCIAS', PREMIO_VALOR_INTEGRAL

    except Exception as e:
        add_to_log(f"Erro no cálculo do prêmio para linha: {row.to_dict()}. Erro: {str(e)}")
        return 'ERRO - VERIFICAR DADOS', 0

def process_data(base_file, absence_file, model_file):
    """Processa os dados do arquivo base e de ausências e usa o modelo fornecido"""
    try:
        df_base = read_excel(base_file)
        df_ausencias = read_excel(absence_file, sheet_name="data")
        df_model = read_excel(model_file)

        if df_base is None or df_ausencias is None or df_model is None:
            add_to_log("Erro: Um ou mais arquivos não foram lidos corretamente.")
            return None

        df_base = normalize_strings(df_base)
        df_ausencias = normalize_strings(df_ausencias)

        # Processar faltas e consolidar ausências
        df_ausencias = process_faltas(df_ausencias)
        df_base['Falta'] = df_ausencias['Falta']
        df_base['Afastamentos'] = df_ausencias.get('Afastamentos', None)
        df_base['Ausência Integral'] = df_ausencias.get('Ausência integral', None)
        df_base['Ausência Parcial'] = df_ausencias.get('Ausência parcial', None)

        # Calcular prêmio por funcionário
        resultados = []
        for _, row in df_base.iterrows():
            status, valor = calcular_premio(row)
            row['Status Prêmio'] = status
            row['Valor Prêmio'] = valor
            resultados.append(row)

        df_resultado = pd.DataFrame(resultados)

        # Ajustar para o modelo fornecido
        for col in df_model.columns:
            if col not in df_resultado.columns:
                df_resultado[col] = None

        df_resultado = df_resultado[df_model.columns]

        add_to_log("Processamento concluído com sucesso.")
        return df_resultado
    except Exception as e:
        add_to_log(f"Erro ao processar dados: {str(e)}")
        return None

def main():
    st.title("Processador de Prêmio Assiduidade")
    
    base_file = st.file_uploader("Arquivo Base", type=['xlsx'])
    absence_file = st.file_uploader("Arquivo de Ausências", type=['xlsx'])
    model_file = st.file_uploader("Modelo de Exportação", type=['xlsx'])

    if base_file and absence_file and model_file:
        if st.button("Processar Dados"):
            df_resultado = process_data(base_file, absence_file, model_file)
            if df_resultado is not None:
                st.success("Dados processados com sucesso!")
                st.download_button("Baixar Relatório", data=df_resultado.to_csv(index=False).encode('utf-8'), file_name="resultado.csv")
                
            st.download_button("Baixar Log", data=generate_log_file(), file_name="log.txt")

if __name__ == "__main__":
    main()
