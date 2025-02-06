import streamlit as st
import pandas as pd
import unicodedata
import traceback
from io import BytesIO
import logging

logging.basicConfig(level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s: %(message)s')

log_messages = []

def add_to_log(message, level='info'):
    global log_messages
    log_entry = f"{level.upper()}: {message}"
    log_messages.append(log_entry)
    
    if level == 'debug':
        logging.debug(message)
    elif level == 'warning':
        logging.warning(message)
    elif level == 'error':
        logging.error(message)
    else:
        logging.info(message)

def generate_log_file():
    log_data = "\n".join(log_messages)
    return BytesIO(log_data.encode('utf-8'))

PREMIO_VALOR_INTEGRAL = 300.00
PREMIO_VALOR_PARCIAL = 150.00
SALARIO_LIMITE = 2542.86

def read_excel(file, sheet_name=0):
    """Lê arquivo Excel com tratamento robusto"""
    try:
        df = pd.read_excel(
            file,
            sheet_name=sheet_name,
            engine='openpyxl'
        )
        
        if df.empty:
            df = pd.read_excel(
                file,
                sheet_name=sheet_name,
                header=None,
                engine='openpyxl'
            )
        
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        
        if not df.empty:
            df.columns = df.columns.astype(str)
            df.columns = df.columns.str.strip()
            
            add_to_log(f"Arquivo {file.name} lido com sucesso", 'info')
            return df
            
        add_to_log(f"Arquivo {file.name} está vazio", 'error')
        return None
        
    except Exception as e:
        add_to_log(f"Erro ao ler {file.name}: {str(e)}", 'error')
        return None

def process_faltas(df):
    """Processa faltas com tratamento robusto"""
    try:
        falta_cols = ['Falta', 'falta', 'Faltas', 'FALTA', 'FALTAS']
        
        falta_col = None
        for col in df.columns:
            if any(fc.lower() in col.lower() for fc in falta_cols):
                falta_col = col
                break
        
        if falta_col:
            df['Falta'] = df[falta_col].fillna('').astype(str).apply(
                lambda x: x.lower().count('x') + x.count('1')
            )
        else:
            df['Falta'] = 0
            
        return df
        
    except Exception as e:
        add_to_log(f"Erro no processamento de faltas: {str(e)}", 'error')
        df['Falta'] = 0
        return df

def calcular_premio(row):
    try:
        try:
            salario = float(str(row.get('Salário Mês Atual', 0)).replace(',', '.') or 0)
        except (ValueError, TypeError):
            salario = 0

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
    try:
        df_base = read_excel(base_file)
        df_ausencias = read_excel(absence_file)
        df_model = read_excel(model_file)

        if any(df is None for df in [df_base, df_ausencias, df_model]):
            add_to_log("Um ou mais arquivos não foram lidos", 'error')
            return None

        required_cols = {'Matrícula'}
        for df, name in [(df_base, 'base'), (df_ausencias, 'ausencias')]:
            missing = required_cols - set(df.columns)
            if missing:
                add_to_log(f"Colunas faltando em {name}: {missing}", 'error')
                return None

        df_ausencias = process_faltas(df_ausencias)
      
        df_merge = pd.merge(
            df_base,
            df_ausencias[['Matrícula', 'Falta', 'Afastamentos', 'Ausência Integral', 'Ausência Parcial']],
            on='Matrícula',
            how='left'
        )
  
        df_merge = df_merge.fillna({
            'Falta': 0,
            'Afastamentos': False,
            'Ausência Integral': False,
            'Ausência Parcial': False
        })

        resultados = []
        for _, row in df_merge.iterrows():
            status, valor = calcular_premio(row)
            row_dict = row.to_dict()
            row_dict.update({
                'Status Prêmio': status,
                'Valor Prêmio': valor
            })
            resultados.append(row_dict)

        df_resultado = pd.DataFrame(resultados)

        for col in df_model.columns:
            if col not in df_resultado.columns:
                df_resultado[col] = None

        return df_resultado[df_model.columns]

    except Exception as e:
        add_to_log(f"Erro no processamento: {str(e)}", 'error')
        return None

def main():
    st.title("Processador de Prêmio Assiduidade")
    
    base_file = st.file_uploader("Arquivo Base", type=['xlsx'], key='base')
    absence_file = st.file_uploader("Arquivo de Ausências", type=['xlsx'], key='ausencia')
    model_file = st.file_uploader("Modelo de Exportação", type=['xlsx'], key='modelo')

    log_area = st.expander("Logs de Processamento")

    if base_file and absence_file and model_file:
        if st.button("Processar Dados"):
            log_messages.clear()
            
            add_to_log("Iniciando processamento de dados")
            
            df_resultado = process_data(base_file, absence_file, model_file)
            
            if df_resultado is not None:
                st.success("Dados processados com sucesso!")
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Resultado')
                output.seek(0)

                st.download_button("Baixar Relatório", 
                                 data=output, 
                                 file_name="resultado_assiduidade.xlsx")
                
                with log_area:
                    for log in log_messages:
                        st.text(log)
                
                log_file = generate_log_file()
                st.download_button("Baixar Log Detalhado", 
                                 data=log_file, 
                                 file_name="log_processamento.txt", 
                                 key="download_log")
            else:
                st.error("Falha no processamento. Verifique os logs.")
                
                with log_area:
                    for log in log_messages:
                        st.text(log)
                
                log_file = generate_log_file()
                st.download_button("Baixar Log de Erro", 
                                 data=log_file, 
                                 file_name="log_erro.txt", 
                                 key="download_error_log")

if __name__ == "__main__":
    main()
