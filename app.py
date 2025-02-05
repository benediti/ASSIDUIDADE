import streamlit as st
import pandas as pd
import numpy as np
import io
import base64

# Configurações iniciais
st.set_page_config(page_title="Processador de Prêmio Assiduidade", layout="wide")

# Constantes
PREMIO_VALOR_INTEGRAL = 300.00
PREMIO_VALOR_PARCIAL = 150.00
SALARIO_LIMITE = 2542.86

def calcular_premio(row, horas_mensais):
    """Calcula se funcionário deve receber prêmio e quanto"""
    # Verificar salário
    try:
        salario = float(str(row['Salário']).replace('R$', '').replace('.', '').replace(',', '.'))
        if salario > SALARIO_LIMITE:
            return ('NÃO PAGAR - SALÁRIO MAIOR', 0, '#FF0000')
    except:
        pass

    # Verificar se tem ocorrências
    if pd.isna(row['Data']) and pd.isna(row['Ausência Integral']) and \
       pd.isna(row['Ausência Parcial']) and pd.isna(row['Afastamentos']) and \
       pd.isna(row['Falta']):
        # Verificar horas mensais
        if horas_mensais == 220:
            return ('PAGAR - INTEGRAL', PREMIO_VALOR_INTEGRAL, '#00FF00')
        elif horas_mensais in [110, 120]:
            return ('PAGAR - MEIO PERÍODO', PREMIO_VALOR_PARCIAL, '#00FF00')
        else:
            return ('AVALIAR - HORAS DIFERENTES', 0, '#0000FF')
    elif row['Falta'] == '1':
        return ('NÃO PAGAR - FALTA', 0, '#FF0000')
    elif not pd.isna(row['Afastamentos']):
        return ('NÃO PAGAR - AFASTAMENTO', 0, '#FF0000')
    elif not pd.isna(row['Ausência Integral']) or not pd.isna(row['Ausência Parcial']):
        return ('AVALIAR - AUSÊNCIA', 0, '#0000FF')
    else:
        return ('AVALIAR', 0, '#0000FF')

def read_excel(file):
    """Lê arquivo Excel e trata cabeçalhos"""
    try:
        df = pd.read_excel(
            file,
            engine='openpyxl',
            header=None
        )
        
        if 'Premio Assiduidade' in str(df.iloc[0:3].values):
            for idx, row in df.iterrows():
                if 'Premio Assiduidade' in str(row.values):
                    df = pd.read_excel(
                        file,
                        engine='openpyxl',
                        header=idx
                    )
                    break
        return df
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {str(e)}")
        return None

def process_data(base_file, absence_file):
    """Processa os dados dos arquivos"""
    try:
        df_base = read_excel(base_file)
        if df_base is None:
            return None, None
            
        df_absence = pd.read_excel(absence_file, engine='openpyxl')
        if df_absence is None:
            return None, None

        # Processar dados base primeiro
        base_data = []
        for _, employee in df_base.iterrows():
            if pd.isna(employee['Premio Assiduidade']):
                continue
                
            horas_mensais = employee['Qtd Horas Mensais'] if not pd.isna(employee['Qtd Horas Mensais']) else 0
            
            base_record = {
                'Matrícula': str(employee['Premio Assiduidade']),
                'Nome': employee['__EMPTY'],
                'Cargo': employee['__EMPTY_1'],
                'Código Local': employee['__EMPTY_2'],
                'Local': employee['__EMPTY_3'],
                'Horas Mensais': horas_mensais,
                'Tipo Contrato': employee['__EMPTY_5'],
                'Data Term. Contrato': employee['__EMPTY_6'],
                'Dias Experiência': employee['__EMPTY_7'],
                'Salário': employee['__EMPTY_8'],
                'Ocorrências': []
            }
            
            # Encontrar todas as ocorrências do funcionário
            absences = df_absence[df_absence['Nome'] == base_record['Nome']]
            for _, absence in absences.iterrows():
                ocorrencia = {
                    'Data': absence['Dia'],
                    'Ausência Integral': absence['Ausência integral'],
                    'Ausência Parcial': absence['Ausência parcial'],
                    'Afastamentos': absence['Afastamentos'],
                    'Falta': '1' if absence['Falta'] == 'x' else '0'
                }
                base_record['Ocorrências'].append(ocorrencia)
            
            base_data.append(base_record)

        # Criar DataFrame final
        final_data = []
        for employee in base_data:
            if not employee['Ocorrências']:
                # Funcionário sem ocorrências
                record = {
                    'Matrícula': employee['Matrícula'],
                    'Nome': employee['Nome'],
                    'Cargo': employee['Cargo'],
                    'Código Local': employee['Código Local'],
                    'Local': employee['Local'],
                    'Horas Mensais': employee['Horas Mensais'],
                    'Tipo Contrato': employee['Tipo Contrato'],
                    'Data Term. Contrato': employee['Data Term. Contrato'],
                    'Dias Experiência': employee['Dias Experiência'],
                    'Salário': employee['Salário'],
                    'Data': None,
                    'Ausência Integral': None,
                    'Ausência Parcial': None,
                    'Afastamentos': None,
                    'Falta': None
                }
                status, valor, cor = calcular_premio(record, employee['Horas Mensais'])
                record.update({
                    'Status Prêmio': status,
                    'Valor Prêmio': valor,
                    'Cor': cor
                })
                final_data.append(record)
            else:
                # Funcionário com ocorrências
                for ocorrencia in employee['Ocorrências']:
                    record = {
                        'Matrícula': employee['Matrícula'],
                        'Nome': employee['Nome'],
                        'Cargo': employee['Cargo'],
                        'Código Local': employee['Código Local'],
                        'Local': employee['Local'],
                        'Horas Mensais': employee['Horas Mensais'],
                        'Tipo Contrato': employee['Tipo Contrato'],
                        'Data Term. Contrato': employee['Data Term. Contrato'],
                        'Dias Experiência': employee['Dias Experiência'],
                        'Salário': employee['Salário'],
                        **ocorrencia
                    }
                    status, valor, cor = calcular_premio(record, employee['Horas Mensais'])
                    record.update({
                        'Status Prêmio': status,
                        'Valor Prêmio': valor,
                        'Cor': cor
                    })
                    final_data.append(record)

        df_final = pd.DataFrame(final_data)
        
        # Agrupar por funcionário para mostrar uma linha por pessoa
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
        st.error("Detalhes do erro para debug:")
        st.write(e)
        return None, None

def generate_local_summary(df):
    """Gera resumo por local de trabalho"""
    local_summary = df.groupby(['Código Local', 'Local']).agg({
        'Matrícula': 'count',  # Total de funcionários
        'Valor Prêmio': lambda x: [
            sum(x == PREMIO_VALOR_INTEGRAL),  # Contagem de 300
            sum(x == PREMIO_VALOR_PARCIAL),   # Contagem de 150
            sum(x == 0)                       # Contagem de não pagamentos
        ]
    }).reset_index()

    # Expandir a coluna de valores
    local_summary[['Recebem 300', 'Recebem 150', 'Não Recebem']] = pd.DataFrame(
        local_summary['Valor Prêmio'].tolist(), 
        index=local_summary.index
    )
    local_summary['Total Funcionários'] = local_summary['Matrícula']
    local_summary['Valor Total'] = (
        local_summary['Recebem 300'] * PREMIO_VALOR_INTEGRAL + 
        local_summary['Recebem 150'] * PREMIO_VALOR_PARCIAL
    )

    # Organizar colunas
    return local_summary[[
        'Código Local', 
        'Local', 
        'Total Funcionários',
        'Recebem 300',
        'Recebem 150', 
        'Não Recebem',
        'Valor Total'
    ]]

def download_link(df, filename):
    """Cria link para download do arquivo"""
    csv = df.to_csv(index=False, encoding='utf-8-sig')
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download {filename}</a>'
    return href

def main():
    st.title("Processador de Prêmio Assiduidade")
    
    st.markdown("### 1. Upload dos Arquivos")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Arquivo Base (Premio Assiduidade)")
        base_file = st.file_uploader("Escolha o arquivo base", type=['xlsx', 'xls'])

    with col2:
        st.markdown("#### Arquivo de Ausências")
        absence_file = st.file_uploader("Escolha o arquivo de ausências", type=['xlsx', 'xls'])

    if base_file and absence_file:
        if st.button("Processar Dados"):
            with st.spinner('Processando dados...'):
                df_resumo, df_detalhado = process_data(base_file, absence_file)
                
                if df_resumo is not None and df_detalhado is not None:
                    st.success("Dados processados com sucesso!")
                    
                    # Resumo por Local
                    st.markdown("### 2. Resumo por Local")
                    df_local = generate_local_summary(df_resumo)
                    st.dataframe(df_local.style.format({
                        'Valor Total': 'R$ {:,.2f}'.format
                    }))
                    
                    # Download do resumo por local
                    st.markdown(download_link(df_local, "resumo_por_local.csv"), unsafe_allow_html=True)
                    
                    # Totais por Local
                    st.markdown("#### Totais por Local:")
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        total_300 = df_local['Recebem 300'].sum()
                        st.metric("Total Prêmio R$ 300", total_300)
                        
                    with col2:
                        total_150 = df_local['Recebem 150'].sum()
                        st.metric("Total Prêmio R$ 150", total_150)
                        
                    with col3:
                        total_nao = df_local['Não Recebem'].sum()
                        st.metric("Total Não Recebem", total_nao)
                    
                    st.metric("Valor Total Geral", f"R$ {df_local['Valor Total'].sum():,.2f}")
                    
                    # Resumo por Funcionário
                    st.markdown("### 3. Resumo por Funcionário")
                    def highlight_rows(row):
                        return ['background-color: ' + row['Cor']] * len(row)
                    styled_df = df_resumo.style.apply(highlight_rows, axis=1)
                    st.dataframe(styled_df)
                    
                    # Detalhes das Ocorrências
                    st.markdown("### 4. Detalhes das Ocorrências")
                    styled_df_detail = df_detalhado.style.apply(highlight_rows, axis=1)
                    st.dataframe(styled_df_detail)
                    
                    # Downloads
                    st.markdown("### 5. Downloads")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.markdown(download_link(df_local, "resumo_por_local.csv"), unsafe_allow_html=True)
                    with col2:
                        st.markdown(download_link(df_resumo, "resumo_assiduidade.csv"), unsafe_allow_html=True)
                    with col3:
                        st.markdown(download_link(df_detalhado, "detalhes_assiduidade.csv"), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
