import streamlit as st
import pandas as pd
import unicodedata
import traceback
from io import BytesIO

# Constantes
PREMIO_VALOR_INTEGRAL = 300.00
PREMIO_VALOR_PARCIAL = 150.00
SALARIO_LIMITE = 2542.86

# Lista para armazenar logs
log_messages = []

def add_to_log(message, level='info'):
    """Adiciona mensagens ao log"""
    global log_messages
    log_entry = f"{level.upper()}: {message}"
    log_messages.append(log_entry)

def generate_log_file():
    """Gera um arquivo de log em memória"""
    log_data = "\n".join(log_messages)
    return BytesIO(log_data.encode('utf-8'))

def read_excel(file, sheet_name=None):
    """Lê um arquivo Excel com estratégias robustas"""
    try:
        # Tenta múltiplas estratégias de leitura
        header_options = [0, 1, None]
        
        for header in header_options:
            try:
                df = pd.read_excel(
                    file, 
                    sheet_name=sheet_name, 
                    header=header, 
                    engine='openpyxl'
                )
                
                # Remove colunas sem nome
                if header is not None:
                    df.columns = [str(col).strip() for col in df.columns]
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                
                # Filtra linhas totalmente vazias
                df = df.dropna(how='all')
                
                if not df.empty and len(df.columns) > 0:
                    add_to_log(f"Arquivo '{file.name}' lido com sucesso. Header: {header}", 'info')
                    return df
            
            except Exception as inner_e:
                add_to_log(f"Falha na leitura. Header: {header}. Erro: {str(inner_e)}", 'debug')
        
        add_to_log(f"Falha total ao ler o arquivo {file.name}", 'error')
        return None
    
    except Exception as e:
        add_to_log(f"Erro fatal ao ler arquivo: {str(e)}", 'error')
        return None

def process_faltas(df):
    """Processa a coluna de faltas"""
    try:
        falta_columns = ['Falta', 'falta', 'Faltas', 'FALTA']
        
        falta_column = None
        for col in falta_columns:
            if col in df.columns:
                falta_column = col
                break
        
        if falta_column:
            df['Falta'] = df[falta_column].fillna(0).astype(str).apply(lambda x: x.count('x'))
        else:
            df['Falta'] = 0
        
        return df
    except Exception as e:
        add_to_log(f"Erro ao processar faltas: {str(e)}", 'error')
        return df

def calcular_premio(row):
    """Calcula o prêmio"""
    try:
        # Converte salário
        try:
            salario = float(str(row.get('Salário Mês Atual', 0)).replace(',', '.') or 0)
        except (ValueError, TypeError):
            salario = 0

        # Verifica condições de não pagamento
        if salario > SALARIO_LIMITE:
            return 'NÃO PAGAR - SALÁRIO MAIOR', 0

        condicoes_nao_pagar = [
            ('Falta', row.get('Falta', 0) > 0, 'NÃO PAGAR - FALTA'),
            ('Afastamentos', row.get('Afastamentos', None), 'NÃO PAGAR - AFASTAMENTO'),
            ('Ausência Integral', row.get('Ausência Integral', None), 'NÃO PAGAR - AUSÊNCIA INTEGRAL'),
            ('Ausência Parcial', row.get('Ausência Parcial', None), 'NÃO PAGAR - AUSÊNCIA PARCIAL')
        ]

        for _, valor, mensagem in condicoes_nao_pagar:
            if valor:
                return mensagem, 0

        # Calcula horas
        try:
            horas = int(str(row.get('Qtd Horas Mensais', 0)).replace(',', '.') or 0)
        except (ValueError, TypeError):
            horas = 0
        
        if horas == 220:
            return 'PAGAR', PREMIO_VALOR_INTEGRAL
        elif horas in [110, 120]:
            return 'PAGAR', PREMIO_VALOR_PARCIAL

        return 'PAGAR - SEM OCORRÊNCIAS', PREMIO_VALOR_INTEGRAL

    except Exception as e:
        add_to_log(f"Erro no cálculo do prêmio: {str(e)}", 'error')
        return 'ERRO - VERIFICAR DADOS', 0

def process_data(base_file, absence_file, model_file):
    """Processa os dados"""
    try:
        # Leitura dos arquivos
        df_base = read_excel(base_file)
        df_ausencias = read_excel(absence_file)
        df_model = read_excel(model_file)

        if df_base is None or df_ausencias is None or df_model is None:
            add_to_log("Erro: Um ou mais arquivos não foram lidos corretamente.", 'error')
            return None

        # Processa faltas
        df_ausencias = process_faltas(df_ausencias)

        # Merge dos dataframes
        df_merge = pd.merge(
            df_base, 
            df_ausencias[['Matrícula', 'Falta', 'Afastamentos', 'Ausência Integral', 'Ausência Parcial']], 
            on='Matrícula', 
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

        # Garante a estrutura do modelo
        modelo_colunas = df_model.columns.tolist()
        for col in modelo_colunas:
            if col not in df_resultado.columns:
                df_resultado[col] = None

        # Reordena de acordo com o modelo
        df_resultado = df_resultado[[col for col in modelo_colunas if col in df_resultado.columns]]

        return df_resultado

    except Exception as e:
        add_to_log(f"Erro fatal no processamento: {str(e)}\n{traceback.format_exc()}", 'error')
        return None

def main():
    st.title("Processador de Prêmio Assiduidade")
    
    base_file = st.file_uploader("Arquivo Base", type=['xlsx'], key='base')
    absence_file = st.file_uploader("Arquivo de Ausências", type=['xlsx'], key='ausencia')
    model_file = st.file_uploader("Modelo de Exportação", type=['xlsx'], key='modelo')

    if base_file and absence_file and model_file:
        if st.button("Processar Dados"):
            df_resultado = process_data(base_file, absence_file, model_file)
            
            if df_resultado is not None:
                st.success("Dados processados com sucesso!")
                
                # Exportação do resultado
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Resultado')
                output.seek(0)

                st.download_button("Baixar Relatório", data=output, file_name="resultado_assiduidade.xlsx")
                
                # Mostra log
                st.download_button("Baixar Log", data=generate_log_file(), file_name="log_processamento.txt")
            else:
                st.error("Falha no processamento. Verifique os logs.")

if __name__ == "__main__":
    main()
