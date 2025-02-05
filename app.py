"""
Processador de Prêmio Assiduidade - Ajuste Final com Regras e Layout Personalizado
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

def read_excel(file, header_row=0):
    """Lê o arquivo Excel e retorna o DataFrame com o cabeçalho especificado"""
    try:
        df = pd.read_excel(file, engine='openpyxl', header=header_row)
        st.write("Colunas detectadas no arquivo:", df.columns.tolist())
        return df
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {str(e)}")
        return None

def normalize_column_values(df):
    """Normaliza os valores numéricos para remover vírgulas"""
    try:
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].str.replace(",", ".", regex=False).str.strip()
        return df
    except Exception as e:
        st.error(f"Erro ao normalizar valores das colunas: {str(e)}")
        return df

def calcular_premio(row):
    """Calcula o prêmio com base nas regras de assiduidade"""
    try:
        # Regra 1: Salário acima do limite
        salario = float(row['Salário Mês Atual'])
        if salario > SALARIO_LIMITE:
            return 'NÃO PAGAR - SALÁRIO MAIOR', 0, 'red'

        # Regra 2: Ausências, faltas ou afastamentos
        if pd.notna(row['Falta']) or pd.notna(row['Afastamentos']) or pd.notna(row['Ausência Integral']) or pd.notna(row['Ausência Parcial']):
            if pd.notna(row['Falta']):
                return 'NÃO PAGAR - FALTA', 0, 'red'
            if pd.notna(row['Afastamentos']):
                return 'NÃO PAGAR - AFASTAMENTO', 0, 'red'
            if pd.notna(row['Ausência Integral']) or pd.notna(row['Ausência Parcial']):
                return 'AVALIAR - AUSÊNCIA', 0, 'blue'

        # Regra 3: Verificação de horas
        horas = int(row['Qtd Horas Mensais'])
        if horas == 220:
            return 'PAGAR', PREMIO_VALOR_INTEGRAL, 'green'
        elif horas in [110, 120]:
            return 'PAGAR', PREMIO_VALOR_PARCIAL, 'green'

        # Caso todas as colunas estejam em branco (nenhuma ocorrência)
        return 'PAGAR - SEM OCORRÊNCIAS', PREMIO_VALOR_INTEGRAL, 'green'

    except Exception as e:
        st.error(f"Erro no cálculo do prêmio: {str(e)}")
        return 'ERRO - VERIFICAR DADOS', 0, 'red'

def process_data(base_file, absence_file):
    """Processa os dados do arquivo base e de ausências"""
    try:
        # Ler arquivos
        df_base = read_excel(base_file, header_row=2)  # Cabeçalho na linha 3
        df_ausencias = read_excel(absence_file)

        if df_base is None or df_ausencias is None:
            return None

        # Normalizar valores
        df_base = normalize_column_values(df_base)
        df_ausencias = normalize_column_values(df_ausencias)

        # Consolidar informações de ausências no arquivo base
        df_base['Falta'] = df_ausencias['Falta']
        df_base['Afastamentos'] = df_ausencias['Afastamentos']
        df_base['Ausência Integral'] = df_ausencias['Ausência integral']
        df_base['Ausência Parcial'] = df_ausencias['Ausência parcial']

        # Realizar os cálculos
        resultados = []
        for _, row in df_base.iterrows():
            status, valor, cor = calcular_premio(row)
            row['Status Prêmio'] = status
            row['Valor Prêmio'] = valor
            row['Cor'] = cor
            resultados.append(row)

        df_resultado = pd.DataFrame(resultados)

        # Ordenar colunas no formato solicitado
        colunas_ordem = [
            'Código Funcionário', 'Nome Funcionário', 'Cargo Atual', 'Código Local',
            'Nome Local Funcionário', 'Tipo Contrato', 'Data Term Contrato', 
            'Dias Experiência', 'Salário Mês Atual', 'Qtd Horas Mensais',
            'Status Prêmio', 'Valor Prêmio'
        ]
        colunas_adicionais = [col for col in df_resultado.columns if col not in colunas_ordem]
        df_resultado = df_resultado[colunas_ordem + colunas_adicionais]

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
                    st.dataframe(df_resultado.style.apply(
                        lambda x: [f'background-color: {x["Cor"]}' for _ in x], axis=1
                    ))
                    
                    st.markdown("### Download do Arquivo")
                    st.markdown(download_link(df_resultado, "resultado_assiduidade_consolidado.csv"), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
