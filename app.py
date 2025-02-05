"""
Processador de Prêmio Assiduidade - Revisado para Compatibilidade com Planilhas Base
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
    """Lê o arquivo Excel e exibe as colunas detectadas"""
    try:
        df = pd.read_excel(file, engine='openpyxl')
        st.write("Nomes das colunas detectados:", list(df.columns))
        return df
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {str(e)}")
        return None

def normalize_columns(df):
    """Renomeia colunas para garantir compatibilidade com o código."""
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
    missing_columns = [col for col in column_mapping.keys() if col not in df.columns]
    if missing_columns:
        st.error(f"As seguintes colunas estão ausentes no arquivo: {', '.join(missing_columns)}")
        return None
    df = df.rename(columns=column_mapping)
    return df

def calcular_premio(row, horas_mensais):
    try:
        # Verificar salário
        salario = row['Salário']
        if isinstance(salario, str):
            salario = float(salario.replace('R$', '').replace('.', '').replace(',', '.').strip())
        
        if salario > SALARIO_LIMITE:
            return 'NÃO PAGAR - SALÁRIO MAIOR QUE R$ 2.542,86', 0, '#FF0000'

        # Verificar horas mensais
        horas = horas_mensais
        if isinstance(horas, str):
            horas = float(horas.replace(',', '.'))

        valor_base = 0
        if horas == 220:
            valor_base = PREMIO_VALOR_INTEGRAL
        elif horas in [110, 120]:
            valor_base = PREMIO_VALOR_PARCIAL
        else:
            return 'AVALIAR - HORAS INVÁLIDAS', 0, '#0000FF'

        # Verificar ocorrências
        if row['Falta'] == '1' or row['Falta'] == 'x':
            return 'NÃO PAGAR - FALTA', 0, '#FF0000'

        if not pd.isna(row['Afastamentos']) and str(row['Afastamentos']).strip():
            return 'NÃO PAGAR - AFASTAMENTO', 0, '#FF0000'

        if row.get('Ausência Integral', None) == 'Sim' or row.get('Ausência Parcial', None) != '00:00':
            return 'AVALIAR - AUSÊNCIA', 0, '#0000FF'

        return 'PAGAR', valor_base, '#00FF00'

    except Exception as e:
        st.error(f"Erro no cálculo do prêmio: {str(e)}")
        return 'ERRO - VERIFICAR DADOS', 0, '#FF0000'

def process_data(base_file, absence_file):
    """Processa os dados dos arquivos e verifica a compatibilidade das colunas"""
    try:
        df_base = read_excel(base_file)
        df_base = normalize_columns(df_base)
        if df_base is None:
            return None, None
            
        df_absence = read_excel(absence_file)
        df_absence = normalize_columns(df_absence)
        if df_absence is None:
            return None, None

        base_data = []
        for _, employee in df_base.iterrows():
            if pd.isna(employee['Nome Funcionário']):
                continue
            
            horas_mensais = employee.get('Qtd Horas Mensais', 0)
            
            base_record = {
                'Matrícula': employee.get('Código Funcionário', ''),
                'Nome': employee['Nome Funcionário'],
                'Cargo': employee.get('Cargo Atual', ''),
                'Código Local': employee.get('Código Local', ''),
                'Local': employee.get('Nome Local Funcionário', ''),
                'Horas Mensais': horas_mensais,
                'Tipo Contrato': employee.get('Tipo Contrato', ''),
                'Data Term. Contrato': employee.get('Data Term Contrato', ''),
                'Dias Experiência': employee.get('Dias Experiência', ''),
                'Salário': employee.get('Salário Mês Atual', 0),
                'Ocorrências': []
            }
            
            absences = df_absence[df_absence['Nome Funcionário'] == base_record['Nome']]
            for _, absence in absences.iterrows():
                ocorrencia = {
                    'Data': absence['Dia'],
                    'Ausência Integral': absence['Ausência integral'],
                    'Ausência Parcial': absence['Ausência parcial'],
                    'Afastamentos': absence['Afastamentos'],
                    'Falta': '1' if absence.get('Falta', '') == 'x' else '0'
                }
                base_record['Ocorrências'].append(ocorrencia)
            
            base_data.append(base_record)

        final_data = []
        for employee in base_data:
            if not employee['Ocorrências']:
                record = {k: v for k, v in employee.items() if k != 'Ocorrências'}
                record.update({'Data': None, 'Ausência Integral': None, 'Ausência Parcial': None, 'Afastamentos': None, 'Falta': None})
                status, valor, cor = calcular_premio(record, employee['Horas Mensais'])
                record.update({'Status Prêmio': status, 'Valor Prêmio': valor, 'Cor': cor})
                final_data.append(record)
            else:
                for ocorrencia in employee['Ocorrências']:
                    record = {k: v for k, v in employee.items() if k != 'Ocorrências'}
                    record.update(ocorrencia)
                    status, valor, cor = calcular_premio(record, employee['Horas Mensais'])
                    record.update({'Status Prêmio': status, 'Valor Prêmio': valor, 'Cor': cor})
                    final_data.append(record)

        df_final = pd.DataFrame(final_data)
        df_resumo = df_final.groupby('Matrícula').agg({
            'Nome': 'first',
            'Cargo': 'first',
            'Código Local': 'first',
            'Local': 'first',
            'Horas Mensais': 'first',
            'Tipo Contrato': 'first',
            'Data Term. Contrato': 'first',
            'Dias Experiência': 'first',
            'Salário': 'first',
            'Status Prêmio': lambda x: '; '.join(x.unique()),
            'Valor Prêmio': 'max',
            'Cor': 'first'
        }).reset_index()

        return df_resumo, df_final

    except Exception as e:
        st.error(f"Erro ao processar dados: {str(e)}")
        return None, None

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
                df_resumo, df_detalhado = process_data(base_file, absence_file)
                
                if df_resumo is not None and df_detalhado is not None:
                    st.success("Dados processados com sucesso!")
                    
                    st.markdown("### Resumo por Funcionário")
                    st.dataframe(df_resumo)
                    
                    st.markdown("### Detalhes das Ocorrências")
                    st.dataframe(df_detalhado)
                    
                    st.markdown("### Downloads")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown(download_link(df_resumo, "resumo_assiduidade.csv"), unsafe_allow_html=True)
                    with col2:
                        st.markdown(download_link(df_detalhado, "detalhes_assiduidade.csv"), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
