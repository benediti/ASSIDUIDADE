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
            # Remover linhas com matrícula vazia
            df = df[df['Matricula'].notna() & (df['Matricula'].astype(str).str.strip() != '')]
            
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
    # Converter matrícula para inteiro
    df['Matricula'] = pd.to_numeric(df['Matricula'], errors='coerce')
    df = df.dropna(subset=['Matricula'])
    df['Matricula'] = df['Matricula'].astype(int)
    
    # Converter 'x' para 1 na coluna Falta
    df['Falta'] = df['Falta'].fillna('')
    df['Falta'] = df['Falta'].apply(lambda x: 1 if str(x).lower() == 'x' else 0)
    
    # Processar Ausência Parcial (atrasos)
    def converter_para_horas(tempo):
        if pd.isna(tempo) or tempo == '':
            return 0
        try:
            # Verifica se é no formato HH:MM
            if ':' in str(tempo):
                horas, minutos = map(int, str(tempo).split(':'))
                return horas + minutos/60
            return 0
        except:
            return 0
    
    df['Horas_Atraso'] = df['Ausencia_Parcial'].apply(converter_para_horas)
    
    # Garantir que Afastamentos seja string
    df['Afastamentos'] = df['Afastamentos'].fillna('').astype(str)
    
    # Agrupar por matrícula
    resultado = df.groupby('Matricula').agg({
        'Falta': 'sum',
        'Dia': 'count',
        'Horas_Atraso': 'sum',
        'Afastamentos': lambda x: '; '.join(filter(None, [str(i).strip() for i in x if str(i).strip()]))
    }).reset_index()
    
    # Formatar horas de atraso
    resultado['Horas_Atraso'] = resultado['Horas_Atraso'].apply(lambda x: f"{int(x)}:{int((x % 1) * 60):02d}")
    
    # Renomear colunas
    resultado.columns = ['Matricula', 'Total_Faltas', 'Total_Dias', 'Total_Atrasos', 'Tipos_Afastamento']
    
    return resultado

def carregar_tipos_afastamento():
    """Carrega tipos de afastamento do arquivo"""
    if os.path.exists("tipos_afastamento.pkl"):
        return pd.read_pickle("tipos_afastamento.pkl")
    return pd.DataFrame({"tipo": list(TIPOS_AFASTAMENTO.keys()), 
                        "categoria": list(TIPOS_AFASTAMENTO.values())})

def salvar_tipos_afastamento(df):
    """Salva tipos de afastamento em arquivo"""
    df.to_pickle("tipos_afastamento.pkl")

def main():
    st.set_page_config(page_title="Sistema de Verificação de Prêmios", page_icon="🏆", layout="wide")
    st.title("Sistema de Verificação de Prêmios")
    
    # Carregar tipos de afastamento
    df_tipos = carregar_tipos_afastamento()
    
    # Sidebar para configurações e uploads
    with st.sidebar:
        st.header("Configurações")
        
        # Upload da base de funcionários
        st.subheader("Base de Funcionários")
        uploaded_func = st.file_uploader("Carregar base de funcionários", type=['xlsx'])
        
        # Upload da base de ausências
        st.subheader("Base de Ausências")
        uploaded_ausencias = st.file_uploader("Carregar base de ausências", type=['xlsx'])
        
        # Gerenciamento de tipos de afastamento
        st.subheader("Tipos de Afastamento")
        uploaded_tipos = st.file_uploader("Atualizar tipos de afastamento", type=['xlsx'])
        
        if uploaded_tipos is not None:
            try:
                df_tipos_novo = pd.read_excel(uploaded_tipos)
                if 'tipo' in df_tipos_novo.columns and 'categoria' in df_tipos_novo.columns:
                    df_tipos = df_tipos_novo
                    salvar_tipos_afastamento(df_tipos)
                    st.success("Tipos de afastamento atualizados!")
                else:
                    st.error("Arquivo deve conter colunas 'tipo' e 'categoria'")
            except Exception as e:
                st.error(f"Erro ao processar arquivo: {str(e)}")
        
        # Mostrar tipos atuais
        st.write("Tipos de Afastamento Atuais:")
        st.dataframe(df_tipos)
    
    # Processamento da base de funcionários
    df_funcionarios = None
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
        except Exception as e:
            st.error(f"Erro ao processar base de funcionários: {str(e)}")
    
    # Processamento da base de ausências
    if uploaded_ausencias is not None:
        try:
            df_ausencias = pd.read_excel(uploaded_ausencias)
            # Renomear colunas removendo espaços e caracteres especiais
            df_ausencias = pd.read_excel(uploaded_ausencias)
            # Corrigindo a ordem das colunas
            df_ausencias = df_ausencias.rename(columns={
                df_ausencias.columns[0]: "Nome",
                df_ausencias.columns[1]: "Matricula",
                df_ausencias.columns[2]: "Centro_de_Custo",
                df_ausencias.columns[3]: "Dia",
                df_ausencias.columns[4]: "Ausencia_Integral",
                df_ausencias.columns[5]: "Ausencia_Parcial",
                df_ausencias.columns[6]: "Afastamentos",
                df_ausencias.columns[7]: "Falta",
                df_ausencias.columns[8]: "Data_de_Demissao"
            })
            
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
