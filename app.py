import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io
import logging

logging.basicConfig(
    filename='sistema_premios.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def processar_ausencias(df):
    try:
        # Garantir que as colunas estejam presentes
        colunas_necessarias = [
            "Matrícula", "Nome", "Centro de Custo", "Dia",
            "Ausência Integral", "Ausência Parcial", "Afastamentos",
            "Falta", "Data de Demissão"
        ]
        
        # Verificar se todas as colunas necessárias estão presentes
        colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
        if colunas_faltantes:
            raise ValueError(f"Colunas faltantes: {', '.join(colunas_faltantes)}")
        
        # Renomear colunas removendo acentos
        df = df.rename(columns={
            "Matrícula": "Matricula",
            "Centro de Custo": "Centro_de_Custo",
            "Ausência Integral": "Ausencia_Integral",
            "Ausência Parcial": "Ausencia_Parcial",
            "Data de Demissão": "Data_de_Demissao"
        })
        
        # Converter matrícula para inteiro
        df['Matricula'] = pd.to_numeric(df['Matricula'], errors='coerce')
        df = df.dropna(subset=['Matricula'])
        df['Matricula'] = df['Matricula'].astype(int)
        
        # Processar faltas
        df['Faltas'] = df['Falta'].fillna('')
        df['Faltas'] = df['Faltas'].apply(lambda x: 1 if str(x).upper().strip() == 'X' else 0)
        
        # Processar atrasos
        def converter_para_horas(tempo):
            if pd.isna(tempo) or tempo == '' or tempo == '00:00':
                return 0
            try:
                if ':' in str(tempo):
                    horas, minutos = map(int, str(tempo).split(':'))
                    return horas + minutos/60
                return 0
            except:
                return 0
        
        df['Horas_Atraso'] = df['Ausencia_Parcial'].apply(converter_para_horas)
        
        # Processar afastamentos
        df['Afastamentos'] = df['Afastamentos'].fillna('').astype(str)
        
        # Agrupar por matrícula
        resultado = df.groupby('Matricula').agg({
            'Faltas': 'sum',
            'Horas_Atraso': 'sum',
            'Afastamentos': lambda x: '; '.join(filter(None, set(x)))  # Remove duplicatas
        }).reset_index()
        
        # Formatar horas de atraso
        resultado['Atrasos'] = resultado['Horas_Atraso'].apply(
            lambda x: f"{int(x)}:{int((x % 1) * 60):02d}" if x > 0 else ""
        )
        
        resultado = resultado.drop('Horas_Atraso', axis=1)
        
        return resultado
        
    except Exception as e:
        logging.error(f"Erro ao processar ausências: {str(e)}")
        raise

def calcular_premio(df_funcionarios, df_ausencias, data_limite_admissao):
    try:
        # Converter data limite para datetime
        data_limite = pd.to_datetime(data_limite_admissao, format='%d/%m/%Y')
        
        # Filtrar funcionários pela data de admissão
        df_funcionarios = df_funcionarios[
            pd.to_datetime(df_funcionarios['Data_Admissao']) <= data_limite
        ]
        
        resultados = []
        for _, func in df_funcionarios.iterrows():
            try:
                # Buscar ausências do funcionário
                ausencias = df_ausencias[df_ausencias['Matricula'] == func['Matricula']]
                
                # Calcular valor base do prêmio
                valor_premio = 0
                if func['Qtd_Horas_Mensais'] == 220:
                    valor_premio = 300.00
                elif func['Qtd_Horas_Mensais'] <= 110:
                    valor_premio = 150.00
                
                # Verificar afastamentos
                afastamentos = ausencias['Afastamentos'].fillna('').str.cat(sep=' ').lower()
                faltas = ausencias['Faltas'].sum() if 'Faltas' in ausencias else 0
                
                status = "Tem direito"
                if faltas > 0:
                    status = "Não tem direito"
                
                # Adicionar ao resultado
                resultados.append({
                    'Matricula': func['Matricula'],
                    'Nome': func['Nome_Funcionario'],
                    'Cargo': func['Cargo'],
                    'Local': func['Nome_Local'],
                    'Horas_Mensais': func['Qtd_Horas_Mensais'],
                    'Data_Admissao': func['Data_Admissao'],
                    'Faltas': faltas,
                    'Valor_Premio': valor_premio if status == "Tem direito" else 0,
                    'Status': status
                })
                
            except Exception as e:
                logging.error(f"Erro ao processar funcionário {func['Matricula']}: {str(e)}")
                continue
        
        return pd.DataFrame(resultados)
        
    except Exception as e:
        logging.error(f"Erro ao calcular prêmios: {str(e)}")
        raise

def main():
    st.set_page_config(page_title="Sistema de Verificação de Prêmios", 
                       page_icon="🏆", 
                       layout="wide")
    
    st.title("Sistema de Verificação de Prêmios")
    
    with st.sidebar:
        data_limite = st.date_input(
            "Data Limite de Admissão",
            help="Funcionários admitidos após esta data não terão direito ao prêmio",
            format="DD/MM/YYYY"
        )
        
        uploaded_func = st.file_uploader("Base de Funcionários", type=['xlsx'])
        uploaded_ausencias = st.file_uploader("Base de Ausências", type=['xlsx'])
    
    if uploaded_func and uploaded_ausencias and data_limite:
        try:
            # Carregar base de funcionários
            df_funcionarios = pd.read_excel(uploaded_func)
            
            if len(df_funcionarios.columns) != 11:
                st.error("A base de funcionários deve ter 11 colunas")
                return
            
            df_funcionarios.columns = [
                "Matricula", "Nome_Funcionario", "Cargo", 
                "Codigo_Local", "Nome_Local", "Qtd_Horas_Mensais",
                "Tipo_Contrato", "Data_Termino_Contrato", 
                "Dias_Experiencia", "Salario_Mes_Atual", "Data_Admissao"
            ]
            
            # Carregar base de ausências
            df_ausencias = pd.read_excel(uploaded_ausencias)
            
            # Processar ausências
            df_ausencias_proc = processar_ausencias(df_ausencias)
            
            # Calcular prêmios
            df_resultado = calcular_premio(df_funcionarios, df_ausencias_proc, data_limite)
            
            # Exibir resultados
            st.subheader("Resultado do Cálculo")
            
            # Filtros
            col1, col2 = st.columns(2)
            with col1:
                filtro_status = st.multiselect(
                    "Status",
                    options=sorted(df_resultado['Status'].unique())
                )
            
            with col2:
                filtro_local = st.multiselect(
                    "Local",
                    options=sorted(df_resultado['Local'].unique())
                )
            
            # Aplicar filtros
            df_mostrar = df_resultado
            if filtro_status:
                df_mostrar = df_mostrar[df_mostrar['Status'].isin(filtro_status)]
            if filtro_local:
                df_mostrar = df_mostrar[df_mostrar['Local'].isin(filtro_local)]
            
            # Mostrar dados
            st.dataframe(df_mostrar)
            
            # Estatísticas
            st.subheader("Estatísticas")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Funcionários", len(df_mostrar))
            with col2:
                st.metric("Com Direito", len(df_mostrar[df_mostrar['Status'] == "Tem direito"]))
            with col3:
                st.metric("Total Prêmios", f"R$ {df_mostrar['Valor_Premio'].sum():,.2f}")
            
            # Exportar
            if st.button("Exportar Resultados"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_mostrar.to_excel(writer, index=False)
                
                st.download_button(
                    "⬇️ Download Excel",
                    data=output.getvalue(),
                    file_name="premios.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"Erro: {str(e)}")
            logging.error(f"Erro na execução: {str(e)}")

if __name__ == "__main__":
    main()
