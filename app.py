import streamlit as st
import pandas as pd
import io
import base64

st.set_page_config(page_title="Processador de Prêmio Assiduidade", layout="wide")

def process_data(base_file, absence_file):
    try:
        # Ler arquivos
        df_base = pd.read_excel(base_file)
        df_absence = pd.read_excel(absence_file)

        # Filtrar linhas válidas do arquivo base
        df_base = df_base[df_base['Premio Assiduidade'].notna()]

        # Processar dados
        processed_data = []
        
        for _, employee in df_base.iterrows():
            matricula = str(employee['Premio Assiduidade'])
            nome = employee['__EMPTY']
            
            # Encontrar ausências do funcionário
            absences = df_absence[df_absence['Nome'] == nome]
            
            # Criar registro do funcionário
            employee_record = {
                'Matrícula': matricula,
                'Nome': nome,
                'Cargo': employee['__EMPTY_1'],
                'Código Local': employee['__EMPTY_2'],
                'Local': employee['__EMPTY_3'],
                'Horas Mensais': employee['__EMPTY_4'],
                'Tipo Contrato': employee['__EMPTY_5'],
                'Data Term. Contrato': employee['__EMPTY_6'],
                'Dias Experiência': employee['__EMPTY_7'],
                'Salário': employee['__EMPTY_8']
            }
            
            # Adicionar ocorrências
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
                    processed_data.append(record)
            else:
                processed_data.append(employee_record)

        return pd.DataFrame(processed_data)
    except Exception as e:
        st.error(f"Erro ao processar dados: {str(e)}")
        return None

def download_link(df, filename):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download do arquivo CSV</a>'
    return href

def main():
    st.title("Processador de Prêmio Assiduidade")
    
    st.markdown("### 1. Upload dos Arquivos")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Arquivo Base (Premio Assiduidade)")
        base_file = st.file_uploader("Escolha o arquivo base", type=['xls', 'xlsx'])

    with col2:
        st.markdown("#### Arquivo de Ausências")
        absence_file = st.file_uploader("Escolha o arquivo de ausências", type=['xls', 'xlsx'])

    if base_file and absence_file:
        if st.button("Processar Dados"):
            with st.spinner('Processando dados...'):
                df_processed = process_data(base_file, absence_file)
                
                if df_processed is not None:
                    st.success("Dados processados com sucesso!")
                    
                    # Mostrar dados
                    st.markdown("### 2. Visualização dos Dados")
                    st.dataframe(df_processed)
                    
                    # Link para download
                    st.markdown("### 3. Download dos Resultados")
                    st.markdown(download_link(df_processed, "relatorio_assiduidade.csv"), unsafe_allow_html=True)
                    
                    # Estatísticas
                    st.markdown("### 4. Estatísticas")
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric("Total de Funcionários", len(df_processed['Matrícula'].unique()))
                    
                    with col2:
                        faltas = df_processed[df_processed['Falta'] == '1']
                        st.metric("Total de Faltas", len(faltas))
                    
                    with col3:
                        afastamentos = df_processed[df_processed['Afastamentos'].notna()]
                        st.metric("Total de Afastamentos", len(afastamentos))

if __name__ == "__main__":
    main()