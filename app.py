import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io
import logging

# Configuração do logging
logging.basicConfig(
    filename='sistema_premios.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def carregar_tipos_afastamento():
    if os.path.exists("tipos_afastamento.pkl"):
        return pd.read_pickle("tipos_afastamento.pkl")
    return pd.DataFrame({"tipo": [], "categoria": []})

def salvar_tipos_afastamento(df):
    df.to_pickle("tipos_afastamento.pkl")

def processar_ausencias(df):
    # Renomear colunas com acentos
    df = df.rename(columns={
        "Matrícula": "Matricula",
        "Centro de Custo": "Centro_de_Custo",
        "Ausência Integral": "Ausencia_Integral",
        "Ausência Parcial": "Ausencia_Parcial",
        "Data de Demissão": "Data_de_Demissao"
    })
    
    # Converter Matricula para inteiro
    df['Matricula'] = pd.to_numeric(df['Matricula'], errors='coerce')
    df = df.dropna(subset=['Matricula'])
    df['Matricula'] = df['Matricula'].astype(int)
    
    # Processar faltas (X = 1, vazio = 0)
    df['Faltas'] = df['Falta'].fillna('')
    df['Faltas'] = df['Faltas'].apply(lambda x: 1 if str(x).upper().strip() == 'X' else 0)
    
    # Processar atrasos
    def converter_para_horas(tempo):
        if pd.isna(tempo) or tempo == '' or tempo == '00:00':
            return 0
        try:
            if ':' in str(tempo):
                horas, minutos = map(int, str(tempo).split(':'))
                return horas + minutos / 60
            return 0
        except:
            return 0
    
    df['Horas_Atraso'] = df['Ausencia_Parcial'].apply(converter_para_horas)
    df['Afastamentos'] = df['Afastamentos'].fillna('').astype(str)
    
    # Agrupar por matrícula
    resultado = df.groupby('Matricula').agg({
        'Faltas': 'sum',
        'Horas_Atraso': 'sum',
        'Afastamentos': lambda x: '; '.join(sorted(set(filter(None, x))))
    }).reset_index()
    
    # Formatar horas de atraso
    resultado['Atrasos'] = resultado['Horas_Atraso'].apply(
        lambda x: f"{int(x)}:{int((x % 1) * 60):02d}" if x > 0 else ""
    )
    resultado = resultado.drop('Horas_Atraso', axis=1)
    
    return resultado

def calcular_premio(df_funcionarios, df_ausencias, data_limite_admissao):
    # Lista de afastamentos que impedem o prêmio
    afastamentos_impeditivos = [
        "Declaração Acompanhante", "Feriado", "Emenda Feriado", 
        "Licença Maternidade", "Declaração INSS (dias)", 
        "Comparecimento Medico INSS", "Aposentado por Invalidez",
        "Atestado Médico", "Atestado de Óbito", "Licença Paternidade",
        "Licença Casamento", "Acidente de Trabalho", "Auxilio Doença",
        "Primeira Suspensão", "Segunda Suspensão", "Férias",
        "Falta não justificada", "Processo",
        "Falta não justificada (dias)", "Atestado Médico (dias)"
    ]
    
    # Filtrar pela data de admissão
    df_funcionarios['Data_Admissao'] = pd.to_datetime(df_funcionarios['Data_Admissao'], format='%d/%m/%Y')
    df_funcionarios = df_funcionarios[df_funcionarios['Data_Admissao'] <= pd.to_datetime(data_limite_admissao)]
    
    resultados = []
    for _, func in df_funcionarios.iterrows():
        ausencias = df_ausencias[df_ausencias['Matricula'] == func['Matricula']]
        
        tem_afastamento_impeditivo = False
        if not ausencias.empty:
            afastamentos = ' '.join(ausencias['Afastamentos'].fillna('').astype(str)).lower()
            for afastamento in afastamentos_impeditivos:
                if afastamento.lower() in afastamentos:
                    tem_afastamento_impeditivo = True
                    break
        
        # Calcular valor do prêmio
        valor_premio = 0
        if func['Qtd_Horas_Mensais'] == 220:
            valor_premio = 300.00
        elif func['Qtd_Horas_Mensais'] <= 110:
            valor_premio = 150.00
        
        # Determinar status
        status = "Não tem direito" if tem_afastamento_impeditivo else "Tem direito"
        
        # Adicionar ao resultado
        resultados.append({
            'Matricula': func['Matricula'],
            'Nome': func['Nome_Funcionario'],
            'Cargo': func['Cargo'],
            'Local': func['Nome_Local'],
            'Horas_Mensais': func['Qtd_Horas_Mensais'],
            'Data_Admissao': func['Data_Admissao'],
            'Valor_Premio': valor_premio if status == "Tem direito" else 0,
            'Status': status
        })
    
    return pd.DataFrame(resultados)

def exportar_excel(df_resultado):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resultado.to_excel(writer, index=False, sheet_name='Resultados')
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Sistema de Verificação de Prêmios", page_icon="🏆", layout="wide")
    st.title("Sistema de Verificação de Prêmios")
    
    with st.sidebar:
        st.header("Configurações")
        data_limite = st.date_input("Data Limite de Admissão", value=datetime.now())
        uploaded_func = st.file_uploader("Carregar base de funcionários", type=['xlsx'])
        uploaded_ausencias = st.file_uploader("Carregar base de ausências", type=['xlsx'])
    
    if uploaded_func and uploaded_ausencias:
        df_funcionarios = pd.read_excel(uploaded_func)
        df_ausencias = pd.read_excel(uploaded_ausencias)
        df_ausencias = processar_ausencias(df_ausencias)
        df_resultado = calcular_premio(df_funcionarios, df_ausencias, data_limite)
        
        st.subheader("Resultados")
        st.dataframe(df_resultado)
        
        if st.button("Exportar Resultados para Excel"):
            output = exportar_excel(df_resultado)
            st.download_button(
                label="📥 Baixar Excel",
                data=output,
                file_name="resultado_premios.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
