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
                return horas + minutos/60
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
    
    # Processar tipos de afastamento
    df_tipos = carregar_tipos_afastamento()
    for tipo in df_tipos['tipo'].unique():
        resultado[tipo] = resultado['Afastamentos'].str.contains(tipo, case=False).astype(int)
        resultado[tipo] = resultado[tipo].apply(lambda x: x if x > 0 else "")
    
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
    
    # Afastamentos que precisam de decisão
    afastamentos_decisao = ["Abono", "Atraso"]
    
    # Afastamentos permitidos
    afastamentos_permitidos = [
        "Folga Gestor", "Abonado Gerencia Loja",
        "Confraternização universal", "Aniversario de São Paulo"
    ]
    
    # Filtrar pela data de admissão
    df_funcionarios['Data_Admissao'] = pd.to_datetime(df_funcionarios['Data_Admissao'], format='%d/%m/%Y')
    df_funcionarios = df_funcionarios[df_funcionarios['Data_Admissao'] <= pd.to_datetime(data_limite_admissao)]
    
    resultados = []
    for _, func in df_funcionarios.iterrows():
        ausencias = df_ausencias[df_ausencias['Matricula'] == func['Matricula']]
        
        tem_afastamento_impeditivo = False
        tem_afastamento_decisao = False
        tem_apenas_permitidos = False
        
        if not ausencias.empty:
            afastamentos = ' '.join(ausencias['Afastamentos'].fillna('').astype(str)).lower()
            
            # Verificar afastamentos impeditivos
            for afastamento in afastamentos_impeditivos:
                if afastamento.lower() in afastamentos:
                    tem_afastamento_impeditivo = True
                    break
            
            # Verificar afastamentos que precisam de decisão
            if not tem_afastamento_impeditivo:
                for afastamento in afastamentos_decisao:
                    if afastamento.lower() in afastamentos:
                        tem_afastamento_decisao = True
                        break
            
            # Verificar se tem apenas afastamentos permitidos
            tem_apenas_permitidos = not tem_afastamento_impeditivo and not tem_afastamento_decisao
        
        # Calcular valor do prêmio
        valor_premio = 0
        if func['Qtd_Horas_Mensais'] == 220:
            valor_premio = 300.00
        elif func['Qtd_Horas_Mensais'] <= 110:
            valor_premio = 150.00
        
        # Determinar status
        status = "Não tem direito"
        if not tem_afastamento_impeditivo:
            if tem_afastamento_decisao:
                status = "Aguardando decisão"
            elif tem_apenas_permitidos or ausencias.empty:
                status = "Tem direito"
        
        # Adicionar ao resultado
        resultados.append({
            'Matricula': func['Matricula'],
            'Nome': func['Nome_Funcionario'],
            'Cargo': func['Cargo'],
            'Local': func['Nome_Local'],
            'Horas_Mensais': func['Qtd_Horas_Mensais'],
            'Data_Admissao': func['Data_Admissao'],
            'Valor_Premio': valor_premio if status == "Tem direito" else 0,
            'Status': status,
            'Detalhes_Afastamentos': ausencias['Afastamentos'].iloc[0] if not ausencias.empty else ''
        })
    
    return pd.DataFrame(resultados)

def main():
    st.set_page_config(page_title="Sistema de Verificação de Prêmios", page_icon="🏆", layout="wide")
    st.title("Sistema de Verificação de Prêmios")
    
    with st.sidebar:
        st.header("Configurações")
        
        data_limite = st.date_input(
            "Data Limite de Admissão",
            help="Funcionários admitidos após esta data não terão direito ao prêmio",
            format="DD/MM/YYYY"
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
            
            # Calcular prêmios
            df_resultado = calcular_premio(df_funcionarios, df_ausencias, data_limite)
            
            st.subheader("Resultado do Cálculo de Prêmios")
            
            # Filtros
            col1, col2 = st.columns(2)
            with col1:
                filtro_status = st.multiselect(
                    "Filtrar por Status",
                    options=sorted(df_resultado['Status'].unique())
                )
            
            with col2:
                filtro_local = st.multiselect(
                    "Filtrar por Local",
                    options=sorted(df_resultado['Local'].unique())
                )
            
            # Aplicar filtros
            df_mostrar = df_resultado
            if filtro_status:
                df_mostrar = df_mostrar[df_mostrar['Status'].isin(filtro_status)]
            if filtro_local:
                df_mostrar = df_mostrar[df_mostrar['Local'].isin(filtro_local)]
            
            # Mostrar dados
            st.dataframe(
                df_mostrar,
                column_config={
                    "Matricula": st.column_config.NumberColumn("Matrícula", format="%d"),
                    "Valor_Premio": st.column_config.NumberColumn("Valor Prêmio", format="R$ %.2f"),
                    "Data_Admissao": st.column_config.DateColumn("Data Admissão", format="DD/MM/YYYY")
                }
            )
            
            # Estatísticas
            st.subheader("Estatísticas")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total de Funcionários", len(df_mostrar))
            with col2:
                st.metric("Total com Direito", len(df_mostrar[df_mostrar['Status'] == "Tem direito"]))
            with col3:
                st.metric("Valor Total Prêmios", f"R$ {df_mostrar['Valor_Premio'].sum():.2f}")
            
            # Exportar resultados
            if st.button("Exportar Resultados"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_mostrar.to_excel(writer, index=False, sheet_name='Resultados')
                
                st.download_button(
                    label="Download Excel",
                    data=output.getvalue(),
                    file_name="resultado_premios.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        except Exception as e:
            st.error(f"Erro ao processar dados: {str(e)}")

if __name__ == "__main__":
    main()
