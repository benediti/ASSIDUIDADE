import streamlit as st
import pandas as pd
import os
from datetime import datetime
import logging

# Configuração do logging
logging.basicConfig(
    filename='sistema_premios.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def verificar_estrutura_dados(df):
    """
    Verifica se todas as colunas necessárias estão presentes e com os tipos corretos
    Retorna (sucesso, mensagem)
    """
    colunas_esperadas = {
        "Matricula": "numeric",
        "Nome_Funcionario": "string",
        "Cargo": "string",
        "Codigo_Local": "numeric",
        "Nome_Local": "string",
        "Qtd_Horas_Mensais": "numeric",
        "Tipo_Contrato": "string",
        "Data_Termino_Contrato": "datetime",
        "Dias_Experiencia": "numeric",
        "Salario_Mes_Atual": "numeric"
    }
    
    # Colunas que podem conter valores nulos
    colunas_permitir_nulos = ["Data_Termino_Contrato", "Dias_Experiencia"]
    
    info = []
    erros = []
    
    # Verificar colunas presentes
    for coluna in colunas_esperadas.keys():
        if coluna not in df.columns:
            erros.append(f"Coluna ausente: {coluna}")
    
    if erros:
        return False, erros
    
    # Verificar tipos de dados e valores nulos
    for coluna, tipo in colunas_esperadas.items():
        try:
            # Verificar valores nulos
            nulos = df[coluna].isnull().sum()
            if nulos > 0:
                if coluna in colunas_permitir_nulos:
                    info.append(f"Informação: Coluna {coluna} contém {nulos} valores em branco (permitido)")
                else:
                    erros.append(f"Coluna {coluna} contém {nulos} valores nulos")
            
            # Verificar tipos de dados
            if tipo == "numeric":
                if coluna in colunas_permitir_nulos:
                    df[coluna] = pd.to_numeric(df[coluna], errors='coerce')
                else:
                    df[coluna] = pd.to_numeric(df[coluna], errors='raise')
            elif tipo == "datetime":
                if coluna in colunas_permitir_nulos:
                    df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
                else:
                    df[coluna] = pd.to_datetime(df[coluna], errors='raise')
            elif tipo == "string":
                df[coluna] = df[coluna].astype(str)
        
        except Exception as e:
            erros.append(f"Erro na coluna {coluna}: {str(e)}")
    
    if erros:
        return False, erros
    
    return True, ["Todas as colunas foram validadas com sucesso!"]

def carregar_base_funcionarios():
    """Carrega a base de funcionários do arquivo salvo"""
    if os.path.exists("base_funcionarios.pkl"):
        return pd.read_pickle("base_funcionarios.pkl")
    return None

def salvar_base_funcionarios(df):
    """Salva a base de funcionários em um arquivo"""
    df.to_pickle("base_funcionarios.pkl")
    logging.info(f"Base de funcionários salva com sucesso. Total de registros: {len(df)}")

def main():
    # Configuração da página
    st.set_page_config(
        page_title="Sistema de Verificação de Prêmios",
        page_icon="🏆",
        layout="wide"
    )
    
    # Título principal
    st.title("Sistema de Verificação de Prêmios")
    
    # Sidebar para gerenciamento da base de funcionários
    with st.sidebar:
        st.header("Base de Funcionários")
        
        # Upload de nova base
        uploaded_file = st.file_uploader(
            "Atualizar base de funcionários",
            type=['xlsx'],
            help="Selecione o arquivo 'EQUIPPE Base Funcionarios.xlsx'"
        )
        
        if uploaded_file is not None:
            try:
                # Lendo o arquivo Excel
                df_funcionarios = pd.read_excel(uploaded_file)
                logging.info(f"Arquivo carregado: {uploaded_file.name}")
                
                # Formatando as colunas para garantir consistência
                df_funcionarios.columns = [
                    "Matricula", "Nome_Funcionario", "Cargo", 
                    "Codigo_Local", "Nome_Local", "Qtd_Horas_Mensais",
                    "Tipo_Contrato", "Data_Termino_Contrato", 
                    "Dias_Experiencia", "Salario_Mes_Atual"
                ]
                
                # Verificar estrutura dos dados
                sucesso, mensagens = verificar_estrutura_dados(df_funcionarios)
                
                # Exibir resultados da verificação
                with st.expander("Log de Validação", expanded=True):
                    for msg in mensagens:
                        if msg.startswith("Informação:"):
                            st.info(msg)
                            logging.info(msg)
                        elif sucesso:
                            st.success(msg)
                            logging.info(msg)
                        else:
                            st.error(msg)
                            logging.error(msg)
                
                if sucesso:
                    # Salvando a base atualizada
                    salvar_base_funcionarios(df_funcionarios)
                    st.success("Base de funcionários atualizada com sucesso!")
                    
                    # Mostrando data da última atualização
                    st.info(f"Última atualização: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
                    
                    # Estatísticas do carregamento
                    st.write("Estatísticas do carregamento:")
                    st.write(f"- Total de registros: {len(df_funcionarios)}")
                    st.write(f"- Total de colunas: {len(df_funcionarios.columns)}")
                    
            except Exception as e:
                erro = f"Erro ao processar o arquivo: {str(e)}"
                st.error(erro)
                logging.error(erro)
    
    # Área principal
    # Carregando a base de funcionários
    df_funcionarios = carregar_base_funcionarios()
    
    if df_funcionarios is not None:
        st.subheader("Base de Funcionários Atual")
        
        # Filtros
        col1, col2 = st.columns(2)
        with col1:
            filtro_local = st.multiselect(
                "Filtrar por Local",
                options=sorted(df_funcionarios['Nome_Local'].unique())
            )
        
        with col2:
            filtro_cargo = st.multiselect(
                "Filtrar por Cargo",
                options=sorted(df_funcionarios['Cargo'].unique())
            )
        
        # Aplicando filtros
        df_filtrado = df_funcionarios.copy()
        if filtro_local:
            df_filtrado = df_filtrado[df_filtrado['Nome_Local'].isin(filtro_local)]
        if filtro_cargo:
            df_filtrado = df_filtrado[df_filtrado['Cargo'].isin(filtro_cargo)]
        
        # Mostrando estatísticas básicas
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total de Funcionários", len(df_filtrado))
        with col2:
            st.metric("Total de Locais", df_filtrado['Nome_Local'].nunique())
        with col3:
            st.metric("Total de Cargos", df_filtrado['Cargo'].nunique())
        
        # Mostrando os dados
        st.dataframe(
            df_filtrado,
            column_config={
                "Salario_Mes_Atual": st.column_config.NumberColumn(
                    "Salário",
                    format="R$ %.2f"
                ),
                "Data_Termino_Contrato": st.column_config.DateColumn(
                    "Data Término Contrato",
                    format="DD/MM/YYYY"
                )
            }
        )
        
    else:
        st.warning("Por favor, faça o upload da base de funcionários para começar.")

if __name__ == "__main__":
    main()
