"""
Processador de Prêmio Assiduidade - Revisão Final
"""

import streamlit as st
import pandas as pd
import unicodedata
import base64  # Importação necessária para gerar o link de download
from io import BytesIO

# Configurações iniciais
st.set_page_config(page_title="Processador de Prêmio Assiduidade", layout="wide")

# Constantes
PREMIO_VALOR_INTEGRAL = 300.00
PREMIO_VALOR_PARCIAL = 150.00
SALARIO_LIMITE = 2542.86

def read_excel(file, header_row=None):
    """Lê o arquivo Excel e retorna o DataFrame com o cabeçalho especificado"""
    try:
        df = pd.read_excel(file, engine='openpyxl', header=header_row)
        st.write("Colunas detectadas no arquivo:", df.columns.tolist())
        return df
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {str(e)}")
        return None

def normalize_strings(df):
    """Remove acentuação de strings nas colunas do DataFrame"""
    try:
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].apply(lambda x: unicodedata.normalize('NFKD', str(x)).encode('ascii', 'ignore').decode('utf-8') if isinstance(x, str) else x)
        return df
    except Exception as e:
        st.error(f"Erro ao normalizar strings: {str(e)}")
        return df

def calcular_premio(row):
    """Calcula o prêmio com base nas regras de assiduidade"""
    try:
        salario = float(row['Salário Mês Atual'])
        if salario > SALARIO_LIMITE:
            return 'NÃO PAGAR - SALÁRIO MAIOR', 0

        if pd.notna(row.get('Falta')) or pd.notna(row.get('Afastamentos')) or pd.notna(row.get('Ausência Integral')) or pd.notna(row.get('Ausência Parcial')):
            if pd.notna(row.get('Falta')):
                return 'NÃO PAGAR - FALTA', 0
            if pd.notna(row.get('Afastamentos')):
                return 'NÃO PAGAR - AFASTAMENTO', 0
            if pd.notna(row.get('Ausência Integral')) or pd.notna(row.get('Ausência Parcial')):
                return 'AVALIAR - AUSÊNCIA', 0

        horas = int(row['Qtd Horas Mensais'])
        if horas == 220:
            return 'PAGAR', PREMIO_VALOR_INTEGRAL
        elif horas in [110, 120]:
            return 'PAGAR', PREMIO_VALOR_PARCIAL

        return 'PAGAR - SEM OCORRÊNCIAS', PREMIO_VALOR_INTEGRAL

    except Exception as e:
        st.error(f"Erro no cálculo do prêmio: {str(e)}")
        return 'ERRO - VERIFICAR DADOS', 0

def process_data(base_file, absence_file, model_file):
    """Processa os dados do arquivo base e de ausências e usa o modelo fornecido"""
    try:
        df_base = read_excel(base_file, header_row=2)
        df_ausencias = read_excel(absence_file)
        df_model = read_excel(model_file)

        if df_base is None or df_ausencias is None or df_model is None:
            return None

        df_base = normalize_strings(df_base)
        df_ausencias = normalize_strings(df_ausencias)

        df_base['Falta'] = df_ausencias.get('Falta', pd.Series([None] * len(df_base)))
        df_base['Afastamentos'] = df_ausencias.get('Afastamentos', pd.Series([''] * len(df_base))).apply(
            lambda x: '; '.join(str(x).split(';')).strip() if pd.notna(x) else ''
        )
        df_base['Ausência Integral'] = df_ausencias.get('Ausência integral', pd.Series([None] * len(df_base)))
        df_base['Ausência Parcial'] = df_ausencias.get('Ausência parcial', pd.Series([None] * len(df_base)))

        resultados = []
        for _, row in df_base.iterrows():
            status, valor = calcular_premio(row)
            row['Status Prêmio'] = status
            row['Valor Prêmio'] = valor
            resultados.append(row)

        df_resultado = pd.DataFrame(resultados)

        for col in df_model.columns:
            if col not in df_resultado.columns:
                df_resultado[col] = None

        df_resultado = df_resultado[df_model.columns]

        return df_resultado
    except Exception as e:
        st.error(f"Erro ao processar dados: {str(e)}")
        return None

def download_xlsx(df, filename):
    """Cria link para download do arquivo XLSX"""
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Resultado')
        xlsx_data = output.getvalue()
        b64 = base64.b64encode(xlsx_data).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download {filename}</a>'
        return href
    except Exception as e:
        st.error(f"Erro ao criar link de download XLSX: {str(e)}")
        return None

def main():
    st.title("Processador de Prêmio Assiduidade")
    
    st.markdown("### Upload dos Arquivos")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        base_file = st.file_uploader("Arquivo Base (Premio Assiduidade)", type=['xlsx', 'xls'])

    with col2:
        absence_file = st.file_uploader("Arquivo de Ausências", type=['xlsx', 'xls'])

    with col3:
        model_file = st.file_uploader("Modelo de Exportação", type=['xlsx', 'xls'])

    if base_file and absence_file and model_file:
        if st.button("Processar Dados"):
            with st.spinner('Processando dados...'):
                df_resultado = process_data(base_file, absence_file, model_file)
                
                if df_resultado is not None:
                    st.success("Dados processados com sucesso!")
                    
                    st.markdown("### Resultado Consolidado")
                    st.dataframe(df_resultado)
                    
                    st.markdown("### Download do Arquivo")
                    st.markdown(download_xlsx(df_resultado, "resultado_assiduidade_consolidado.xlsx"), unsafe_allow_html=True)

if __name__ == "__main__":
    main()


