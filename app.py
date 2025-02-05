"""
Processador de Prêmio Assiduidade - Ajustado para Erros de Leitura e Reindexação
"""

import streamlit as st
import pandas as pd
import base64

# Configurações iniciais
st.set_page_config(page_title="Processador de Prêmio Assiduidade", layout="wide")

# Constantes
PREMIO_VALOR_INTEGRAL = 300.00
PREMIO_VALOR_PARCIAL = 150.00
SALARIO_LIMITE = 2542.86

def read_excel(file):
    """Lê o arquivo Excel e retorna o DataFrame"""
    try:
        df = pd.read_excel(file, engine='openpyxl', header=None)
        st.write("Pré-visualização do arquivo:", df.head(10))
        return df
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {str(e)}")
        return None

def process_base(df):
    """Processa o arquivo base para ajustar o cabeçalho e as colunas"""
    try:
        # Encontrar linha com o cabeçalho correto
        for i in range(len(df)):
            if "Salário" in df.iloc[i].values:
                df.columns = df.iloc[i]
                df = df[i + 1:].reset_index(drop=True)
                break
        else:
            st.error("Não foi possível encontrar o cabeçalho correto no arquivo base.")
            return None

        # Renomear colunas importantes
        column_mapping = {
            'Matrícula': 'Código Funcionário',
            'Nome': 'Nome Funcionário',
            'Centro de Custo Código': 'Código Local',
            'Centro de Custo Nome': 'Nome Local Funcionário',
            'Horas': 'Qtd Horas Mensais',
            'Contrato': 'Tipo Contrato',
            'Data Término': 'Data Term Contrato',
            'Experiência': 'Dias Experiência',
            'Salário': 'Salário Mês Atual'
        }
        df = df.rename(columns=column_mapping)
        st.write("Colunas processadas:", df.columns)
        return df
    except Exception as e:
        st.error(f"Erro ao processar arquivo base: {str(e)}")
        return None

def calcular_premio(row, ausencias):
    """Calcula o prêmio com base no layout base e nas ausências"""
    try:
        salario = row['Salário Mês Atual']
        if salario > SALARIO_LIMITE:
            return 'NÃO PAGAR - SALÁRIO ALTO', 0

        horas = row['Qtd Horas Mensais']
        if horas == 220:
            premio = PREMIO_VALOR_INTEGRAL
        elif horas in [110, 120]:
            premio = PREMIO_VALOR_PARCIAL
        else:
            return 'AVALIAR - HORAS DIFERENTES', 0

        # Verificar se há ausências
        funcionario_ausencias = ausencias[ausencias['Código Funcionário'] == row['Código Funcionário']]
        if not funcionario_ausencias.empty:
            for _, ausencia in funcionario_ausencias.iterrows():
                if ausencia['Falta'] == 'x' or ausencia['Ausência Integral'] == 'Sim':
                    return 'NÃO PAGAR - AUSÊNCIA', 0

        return 'PAGAR', premio
    except Exception as e:
        st.error(f"Erro no cálculo do prêmio: {str(e)}")
        return 'ERRO - VERIFICAR DADOS', 0

def process_data(base_file, absence_file):
    """Processa os dados do arquivo base e de ausências"""
    try:
        # Ler arquivos
        df_base_raw = read_excel(base_file)
        df_ausencias = read_excel(absence_file)

        if df_base_raw is None or df_ausencias is None:
            return None

        # Processar layout do arquivo base
        df_base = process_base(df_base_raw)
        if df_base is None:
            return None

        # Realizar os cálculos
        resultados = []
        for _, row in df_base.iterrows():
            status, valor = calcular_premio(row, df_ausencias)
            row['Status Prêmio'] = status
            row['Valor Prêmio'] = valor
            resultados.append(row)

        df_resultado = pd.DataFrame(resultados)

        return df_resultado
    except Exception as e:
        st.error(f"Erro ao processar dados: {str(e)}")
        return None

def download_link(df, filename):
    """Cria link para download do arquivo"""
    try:
        csv = df.to_csv(index=False, encoding='utf-8-sig')
        b64 = base64.b64encode(csv.encode()).decode()
        href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download {filename}</a>'
        return href
    except Exception as e:
        st.error(f"Erro ao criar link de download: {str(e)}")
        return None

def main():
    st.title("Processador de Prêmio Assiduidade")
    
    st.markdown("### Upload dos Arquivos")
    
    col1, col2 = st.columns(2)
    
    with col1:
        base_file = st.file_uploader("Arquivo Base (Premio Assiduidade)", type=['xlsx', 'xls'])

    with col2:
        absence_file = st.file_uploader("Arquivo de Ausências", type=['xlsx', 'xls'])

    if base_file and absence_file:
        if st.button("Processar Dados"):
            with st.spinner('Processando dados...'):
                df_resultado = process_data(base_file, absence_file)
                
                if df_resultado is not None:
                    st.success("Dados processados com sucesso!")
                    
                    st.markdown("### Resultado Consolidado")
                    st.dataframe(df_resultado)
                    
                    st.markdown("### Download do Arquivo")
                    st.markdown(download_link(df_resultado, "resultado_assiduidade_consolidado.csv"), unsafe_allow_html=True)

if __name__ == "__main__":
    main()

