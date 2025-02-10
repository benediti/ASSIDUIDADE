import streamlit as st
import pandas as pd
import pdfkit
from datetime import datetime
import os
import logging

# Configuração do logging
logging.basicConfig(
    filename='sistema_premios.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def carregar_tipos_afastamento():
    """Carrega tipos de afastamento do arquivo"""
    if os.path.exists("tipos_afastamento.pkl"):
        return pd.read_pickle("tipos_afastamento.pkl")
    return pd.DataFrame({"tipo": [], "categoria": []})

def salvar_tipos_afastamento(df):
    """Salva tipos de afastamento em arquivo"""
    df.to_pickle("tipos_afastamento.pkl")

def processar_ausencias(df):
    """Processa e agrupa as ausências por matrícula"""
    df['Matricula'] = pd.to_numeric(df['Matricula'], errors='coerce')
    df = df.dropna(subset=['Matricula'])
    df['Matricula'] = df['Matricula'].astype(int)
    
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
    
    resultado = df.groupby('Matricula').agg({
        'Horas_Atraso': 'sum',
        'Afastamentos': lambda x: '; '.join(sorted(set(filter(None, x))))
    }).reset_index()
    
    resultado['Atrasos'] = resultado['Horas_Atraso'].apply(
        lambda x: f"{int(x)}:{int((x % 1) * 60):02d}" if x > 0 else ""
    )
    resultado = resultado.drop('Horas_Atraso', axis=1)
    
    return resultado

def main():
    st.set_page_config(page_title="Sistema de Verificação de Prêmios", page_icon="🏆", layout="wide")
    st.title("Sistema de Verificação de Prêmios")

    with st.sidebar:
        st.header("Configurações")
        
        data_limite = st.date_input(
            "Data Limite de Admissão",
            help="Funcionários admitidos após esta data não terão direito ao prêmio"
        )
        
        st.subheader("Base de Funcionários")
        uploaded_func = st.file_uploader("Carregar base de funcionários", type=['xlsx'])
        
        st.subheader("Base de Ausências")
        uploaded_ausencias = st.file_uploader("Carregar base de ausências", type=['xlsx'])
        
        st.subheader("Tipos de Afastamento")
        uploaded_tipos = st.file_uploader("Atualizar tipos de afastamento", type=['xlsx'])
        
        if uploaded_tipos is not None:
            try:
                df_tipos_novo = pd.read_excel(uploaded_tipos)
                if 'Nome' in df_tipos_novo.columns and 'Categoria' in df_tipos_novo.columns:
                    df_tipos = df_tipos_novo.rename(columns={'Nome': 'tipo', 'Categoria': 'categoria'})
                    salvar_tipos_afastamento(df_tipos)
                    st.success("Tipos de afastamento atualizados!")
                else:
                    st.error("Arquivo deve conter colunas 'Nome' e 'Categoria'")
            except Exception as e:
                st.error(f"Erro ao processar arquivo: {str(e)}")
    
    if uploaded_func and uploaded_ausencias and data_limite:
        try:
            df_funcionarios = pd.read_excel(uploaded_func)
            df_funcionarios.columns = [
                "Matricula", "Nome_Funcionario", "Cargo", 
                "Codigo_Local", "Nome_Local", "Qtd_Horas_Mensais",
                "Tipo_Contrato", "Data_Termino_Contrato", 
                "Dias_Experiencia", "Salario_Mes_Atual", "Data_Admissao"
            ]
            
            df_funcionarios['Data_Admissao'] = pd.to_datetime(df_funcionarios['Data_Admissao'], format='%d/%m/%Y')
            df_funcionarios['Data_Termino_Contrato'] = pd.to_datetime(df_funcionarios['Data_Termino_Contrato'], format='%d/%m/%Y', errors='coerce')
            
            df_ausencias = pd.read_excel(uploaded_ausencias)
            df_ausencias = df_ausencias.rename(columns={
                "Matrícula": "Matricula",
                "Ausência Parcial": "Ausencia_Parcial"
            })
            df_ausencias = processar_ausencias(df_ausencias)
            
            st.subheader("Resultado do Cálculo de Prêmios")

            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total Funcionários", len(df_funcionarios))
            with col2:
                st.metric("Total Ausências Registradas", len(df_ausencias))

        except Exception as e:
            st.error(f"Erro ao processar os dados: {str(e)}")

if __name__ == "__main__":
    main()
