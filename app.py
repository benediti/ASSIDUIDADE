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
    try:
        if os.path.exists("tipos_afastamento.pkl"):
            return pd.read_pickle("tipos_afastamento.pkl")
        return pd.DataFrame({"tipo": [], "categoria": []})
    except Exception as e:
        logging.error(f"Erro ao carregar tipos de afastamento: {e}")
        st.error("Erro ao carregar tipos de afastamento.")
        return pd.DataFrame({"tipo": [], "categoria": []})

def salvar_tipos_afastamento(df):
    try:
        df.to_pickle("tipos_afastamento.pkl")
    except Exception as e:
        logging.error(f"Erro ao salvar tipos de afastamento: {e}")
        st.error("Erro ao salvar tipos de afastamento.")

def processar_ausencias(df):
    try:
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
            except Exception as e:
                logging.error(f"Erro ao converter tempo para horas: {e}")
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
        
        # Processar tipos de afastamento
        df_tipos = carregar_tipos_afastamento()
        for tipo in df_tipos['tipo'].unique():
            resultado[tipo] = resultado['Afastamentos'].str.contains(tipo, case=False).astype(int)
            resultado[tipo] = resultado[tipo].apply(lambda x: x if x > 0 else "")
        
        return resultado
    except Exception as e:
        logging.error(f"Erro ao processar ausências: {e}")
        st.error("Erro ao processar ausências.")
        return pd.DataFrame()

def main():
    try:
        st.set_page_config(page_title="Sistema de Verificação de Prêmios", page_icon="🏆", layout="wide")
        st.title("Sistema de Verificação de Prêmios")
        
        with st.sidebar:
            st.header("Configurações")
            
            data_limite = st.date_input(
                "Data Limite de Admissão",
                help="Funcionários admitidos após esta data não terão direito ao prêmio",
                value=datetime.now()
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
                    logging.error(f"Erro ao processar arquivo de tipos de afastamento: {e}")
                    st.error(f"Erro ao processar arquivo: {str(e)}")
        
        if uploaded_func is not None and uploaded_ausencias is not None and data_limite is not None:
            try:
                # Carregar base de funcionários
                df_funcionarios = pd.read_excel(uploaded_func)
                df_funcionarios.columns = [
                    "Matricula", "Nome_Funcionario", "Cargo", 
                    "Codigo_Local", "Nome_Local", "Qtd_Horas_Mensais",
                    "Tipo_Contrato", "Data_Termino_Contrato", 
                    "Dias_Experiencia", "Salario_Mes_Atual", "Data_Admissao"
                ]
                
                # Carregar e processar base de ausências
                df_ausencias = pd.read_excel(uploaded_ausencias)
                df_ausencias = processar_ausencias(df_ausencias)
                
                # Exibir mensagem de sucesso
                st.success("Dados carregados e processados com sucesso!")
            
            except Exception as e:
                logging.error(f"Erro ao processar dados: {e}")
                st.error(f"Erro ao processar dados: {str(e)}")
    except Exception as e:
        logging.critical(f"Erro crítico na aplicação: {e}")
        st.error("Erro crítico na aplicação. Verifique os logs para mais detalhes.")

if __name__ == "__main__":
    main()
