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

# Tabela de tipos de afastamento
TIPOS_AFASTAMENTO = {
    'Abonado Gerencia Loja': 'Abonado',
    'Atraso': 'Atraso',
    'Falta': 'Falta',
    'Licença Médica': 'Licença',
    'Férias': 'Férias'
    # Outros tipos podem ser adicionados aqui
}

def verificar_estrutura_dados_funcionarios(df):
    """Verifica estrutura da base de funcionários"""
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
    
    return len(erros) == 0, info + erros

def verificar_estrutura_dados_ausencias(df):
    """Verifica estrutura da base de ausências"""
    colunas_esperadas = {
        "Matricula": "numeric",
        "Nome": "string",
        "Centro_de_Custo": "numeric",
        "Dia": "datetime",
        "Ausencia_Integral": "string",
        "Ausencia_Parcial": "string",
        "Afastamentos": "string",
        "Falta": "string",
        "Data_de_Demissao": "datetime"
    }
    
    info = []
    erros = []
    
    # Verificar colunas presentes
    for coluna in colunas_esperadas.keys():
        if coluna not in df.columns:
            erros.append(f"Coluna ausente: {coluna}")
    
    if not erros:
        try:
            # Converter tipos de dados
            df['Matricula'] = df['Matricula'].astype(int)
            df['Centro_de_Custo'] = pd.to_numeric(df['Centro_de_Custo'], errors='coerce')
            df['Dia'] = pd.to_datetime(df['Dia'], errors='coerce')
            df['Data_de_Demissao'] = pd.to_datetime(df['Data_de_Demissao'], errors='coerce')
            
            # Tratamento especial para a coluna Falta
            df['Falta'] = df['Falta'].fillna('')
            df['Falta'] = df['Falta'].apply(lambda x: 1 if x.lower() == 'x' else 0)
            
            info.append("Estrutura de dados validada com sucesso!")
            
        except Exception as e:
            erros.append(f"Erro ao processar dados: {str(e)}")
    
    return len(erros) == 0, info + erros

def processar_ausencias(df):
    """Processa e agrupa as ausências por matrícula"""
    # Converter 'x' para 1 na coluna Falta
    df['Falta'] = df['Falta'].fillna('')
    df['Falta'] = df['Falta'].apply(lambda x: 1 if str(x).lower() == 'x' else 0)
    
    # Agrupar por matrícula
    resultado = df.groupby('Matricula').agg({
        'Falta': 'sum',
        'Dia': 'count',
        'Afastamentos': lambda x: '; '.join(filter(None, x))
    }).reset_index()
    
    # Renomear colunas
    resultado.columns = ['Matricula', 'Total_Faltas', 'Total_Dias', 'Tipos_Afastamento']
    
    return resultado

def main():
    st.set_page_config(page_title="Sistema de Verificação de Prêmios", page_icon="🏆", layout="wide")
    st.title("Sistema de Verificação de Prêmios")
    
    # Sidebar para configurações e uploads
    with st.sidebar:
        st.header("Configurações")
        
        # Upload da base de funcionários
        st.subheader("Base de Funcionários")
        uploaded_func = st.file_uploader("Carregar base de funcionários", type=['xlsx'])
        
        # Upload da base de ausências
        st.subheader("Base de Ausências")
        uploaded_ausencias = st.file_uploader("Carregar base de ausências", type=['xlsx'])
        
        # Editor da tabela de tipos de afastamento
        st.subheader("Tipos de Afastamento")
        if st.checkbox("Editar Tipos de Afastamento"):
            edited_tipos = {}
            for tipo, categoria in TIPOS_AFASTAMENTO.items():
                nova_categoria = st.text_input(f"Categoria para {tipo}", categoria)
                edited_tipos[tipo] = nova_categoria
            if st.button("Salvar Alterações"):
                TIPOS_AFASTAMENTO.update(edited_tipos)
                st.success("Tipos de afastamento atualizados!")
    
    # Processamento da base de funcionários
    if uploaded_func is not None:
        try:
            df_funcionarios = pd.read_excel(uploaded_func)
            df_funcionarios.columns = [
                "Matricula", "Nome_Funcionario", "Cargo", 
                "Codigo_Local", "Nome_Local", "Qtd_Horas_Mensais",
                "Tipo_Contrato", "Data_Termino_Contrato", 
                "Dias_Experiencia", "Salario_Mes_Atual"
            ]
            df_funcionarios['Matricula'] = df_funcionarios['Matricula'].astype(int)
            
            sucesso, mensagens = verificar_estrutura_dados_funcionarios(df_funcionarios)
            
            with st.expander("Log Base Funcionários", expanded=True):
                for msg in mensagens:
                    if msg.startswith("Informação:"):
                        st.info(msg)
                    elif sucesso:
                        st.success(msg)
                    else:
                        st.error(msg)
    
    # Processamento da base de ausências
    if uploaded_ausencias is not None:
        try:
            df_ausencias = pd.read_excel(uploaded_ausencias)
            # Renomear colunas removendo espaços e caracteres especiais
            df_ausencias.columns = [
                "Matricula", "Nome", "Centro_de_Custo", "Dia",
                "Ausencia_Integral", "Ausencia_Parcial", "Afastamentos",
                "Falta", "Data_de_Demissao"
            ]
            
            sucesso, mensagens = verificar_estrutura_dados_ausencias(df_ausencias)
            
            with st.expander("Log Base Ausências", expanded=True):
                for msg in mensagens:
                    if msg.startswith("Informação:"):
                        st.info(msg)
                    elif sucesso:
                        st.success(msg)
                    else:
                        st.error(msg)
            
            if sucesso:
                # Processar ausências
                df_resumo = processar_ausencias(df_ausencias)
                
                st.subheader("Resumo de Ausências")
                # Filtros
                col1, col2 = st.columns(2)
                with col1:
                    filtro_matricula = st.multiselect(
                        "Filtrar por Matrícula",
                        options=sorted(df_resumo['Matricula'].unique())
                    )
                
                # Aplicar filtros
                df_mostrar = df_resumo
                if filtro_matricula:
                    df_mostrar = df_mostrar[df_mostrar['Matricula'].isin(filtro_matricula)]
                
                # Mostrar dados
                st.dataframe(
                    df_mostrar,
                    column_config={
                        "Matricula": st.column_config.NumberColumn("Matrícula", format="%d"),
                        "Total_Faltas": st.column_config.NumberColumn("Total de Faltas", format="%d"),
                        "Total_Dias": st.column_config.NumberColumn("Total de Dias", format="%d")
                    }
                )
                
                # Estatísticas
                st.subheader("Estatísticas")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total de Funcionários", len(df_mostrar))
                with col2:
                    st.metric("Total de Faltas", df_mostrar['Total_Faltas'].sum())
                with col3:
                    st.metric("Média de Faltas", f"{df_mostrar['Total_Faltas'].mean():.2f}")
        
        except Exception as e:
            st.error(f"Erro ao processar arquivo de ausências: {str(e)}")

if __name__ == "__main__":
    main()
