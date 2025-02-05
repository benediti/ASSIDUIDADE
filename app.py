import streamlit as st
import pandas as pd
import io
import base64

st.set_page_config(page_title="Processador de Prêmio Assiduidade", layout="wide")

PREMIO_VALOR = 300.00

def calcular_premio(row):
    if pd.isna(row['Data']) and pd.isna(row['Ausência Integral']) and \
       pd.isna(row['Ausência Parcial']) and pd.isna(row['Afastamentos']) and \
       pd.isna(row['Falta']):
        return ('PAGAR', PREMIO_VALOR, '#00FF00')  # Verde
    elif row['Falta'] == '1':
        return ('NÃO PAGAR - FALTA', 0, '#FF0000')  # Vermelho
    elif not pd.isna(row['Afastamentos']):
        return ('NÃO PAGAR - AFASTAMENTO', 0, '#FF0000')  # Vermelho
    elif not pd.isna(row['Ausência Integral']) or not pd.isna(row['Ausência Parcial']):
        return ('AVALIAR - AUSÊNCIA', 0, '#0000FF')  # Azul
    else:
        return ('AVALIAR', 0, '#0000FF')  # Azul

def read_excel(file):
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
                        header=idx,
                        names=['Código', 'Nome', 'Cargo', 'Cód. Local', 'Local', 
                               'Horas', 'Tipo Contrato', 'Data Term', 'Dias Exp', 'Salário']
                    )
                    break
        return df
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {str(e)}")
        return None

def process_data(base_file, absence_file):
    try:
        df_base = read_excel(base_file)
        if df_base is None:
            return None
            
        df_absence = pd.read_excel(absence_file, engine='openpyxl')
        if df_absence is None:
            return None

        df_base = df_base[df_base['Código'].notna()]
        processed_data = []
        
        for _, employee in df_base.iterrows():
            matricula = str(employee['Código'])
            nome = employee['Nome']
            
            absences = df_absence[df_absence['Nome'] == nome]
            
            employee_record = {
                'Matrícula': matricula,
                'Nome': nome,
                'Cargo': employee['Cargo'],
                'Local': employee['Local'],
                'Horas': employee['Horas'],
                'Tipo Contrato': employee['Tipo Contrato'],
                'Data Term.': employee['Data Term'],
                'Dias Exp.': employee['Dias Exp'],
                'Salário': employee['Salário']
            }
            
            if not absences.empty:
                for _, absence in absences.iterrows():
                    record = employee_record.copy()
                    record.update({
                        'Data': absence['Dia'],
                        'Ausência Integral': absence['Ausência integral'],
                        'Ausência Parcial': absence['Ausência parcial'],
                        'Afastamentos': absence['Afastamentos'],
                        'Falta': '1' if absence['Falta'] == 'x' else '0'
                    })
                    status, valor, cor = calcular_premio(record)
                    record.update({
                        'Status Prêmio': status,
                        'Valor Prêmio': valor,
                        'Cor': cor
                    })
                    processed_data.append(record)
            else:
                record = employee_record.copy()
                record.update({
                    'Data': None,
                    'Ausência Integral': None,
                    'Ausência Parcial': None,
                    'Afastamentos': None,
                    'Falta': None
                })
                status, valor, cor = calcular_premio(record)
                record.update({
                    'Status Prêmio': status,
                    'Valor Prêmio': valor,
                    'Cor': cor
                })
                processed_data.append(record)

        return pd.DataFrame(processed_data)
    except Exception as e:
        st.error(f"Erro ao processar dados: {str(e)}")
        st.error("Detalhes do erro para debug:")
        st.write(e)
        return None

def download_link(df, filename):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download do arquivo CSV</a>'
    return href

def main():
    st.title("Processador de Prêmio Assiduidade")
    
    st.markdown("### 1. Upload dos Arquivos")
    st.warning("Nota: Para melhor compatibilidade, use arquivos no formato .xlsx")
    
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
                df_processed = process_data(base_file, absence_file)
                
                if df_processed is not None:
                    st.success("Dados processados com sucesso!")
                    
                    # Mostrar dados com cores
                    st.markdown("### 2. Visualização dos Dados")
                    
                    # Converter DataFrame para HTML com cores
                    def highlight_rows(row):
                        return ['background-color: ' + row['Cor']] * len(row)
                    
                    styled_df = df_processed.style.apply(highlight_rows, axis=1)
                    st.dataframe(styled_df)
                    
                    # Link para download
                    st.markdown("### 3. Download dos Resultados")
                    st.markdown(download_link(df_processed, "relatorio_assiduidade.csv"), unsafe_allow_html=True)
                    
                    # Estatísticas
                    st.markdown("### 4. Estatísticas")
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        total_func = len(df_processed['Matrícula'].unique())
                        st.metric("Total de Funcionários", total_func)
                    
                    with col2:
                        total_pagar = len(df_processed[df_processed['Status Prêmio'] == 'PAGAR'])
                        st.metric("Receberão Prêmio", total_pagar)
                    
                    with col3:
                        total_nao_pagar = len(df_processed[df_processed['Status Prêmio'].str.startswith('NÃO PAGAR')])
                        st.metric("Não Receberão", total_nao_pagar)
                    
                    with col4:
                        total_avaliar = len(df_processed[df_processed['Status Prêmio'].str.startswith('AVALIAR')])
                        st.metric("Para Avaliar", total_avaliar)
                    
                    # Valor total dos prêmios
                    st.metric("Valor Total de Prêmios", f"R$ {df_processed['Valor Prêmio'].sum():,.2f}")

if __name__ == "__main__":
    main()
