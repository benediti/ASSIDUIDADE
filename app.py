from datetime import datetime
import pandas as pd
import streamlit as st
import os
import logging
import io
from utils import editar_valores_status, exportar_novo_excel  # Importar funções do utils.py

# Configuração do logging
logging.basicConfig(
    filename='sistema_premios.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def carregar_tipos_afastamento():
    if os.path.exists("data/tipos_afastamento.pkl"):
        return pd.read_pickle("data/tipos_afastamento.pkl")
    return pd.DataFrame({"tipo": [], "categoria": []})

def salvar_tipos_afastamento(df):
    df.to_pickle("data/tipos_afastamento.pkl")
def processar_ausencias(df):
    # Renomear colunas e configurar dados iniciais
    df = df.rename(columns={
        "Matrícula": "Matricula",
        "Centro de Custo": "Centro_de_Custo",
        "Ausência Integral": "Ausencia_Integral",
        "Ausência Parcial": "Ausencia_Parcial",
        "Data de Demissão": "Data_de_Demissao"
    })
    
    df['Matricula'] = pd.to_numeric(df['Matricula'], errors='coerce')
    df = df.dropna(subset=['Matricula'])
    df['Matricula'] = df['Matricula'].astype(int)
    
    df['Faltas'] = df['Falta'].fillna('')
    df['Faltas'] = df['Faltas'].apply(lambda x: 1 if str(x).upper().strip() == 'X' else 0)
    
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
    
    # Carregar tipos de afastamento
    df_tipos = carregar_tipos_afastamento()
    tipos_conhecidos = df_tipos['tipo'].unique() if not df_tipos.empty else []

    # Identificar afastamentos desconhecidos
    df['Afastamentos_Desconhecidos'] = df['Afastamentos'].apply(
        lambda x: '; '.join([a for a in x.split(';') if a.strip() not in tipos_conhecidos])
    )
    
    # Classificar status
    def classificar_status(afastamentos):
        afastamentos_list = afastamentos.split(';')
        if any(a.strip() in afastamentos_impeditivos for a in afastamentos_list):
            return "Não Tem Direito"
        elif any(a.strip() in afastamentos_decisao for a in afastamentos_list):
            return "Aguardando Decisão"
        return "Tem Direito"
    
    afastamentos_impeditivos = [
        "Licença Maternidade", "Atestado Médico", "Férias", "Feriado", ...
    ]
    afastamentos_decisao = ["Abono", "Atraso"]
    
    df['Status'] = df['Afastamentos'].apply(classificar_status)
    
    # Retornar DataFrame atualizado
    return df

def calcular_premio(df_funcionarios, df_ausencias, data_limite_admissao):
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
    
    afastamentos_decisao = ["Abono", "Atraso"]
    
    afastamentos_permitidos = [
        "Folga Gestor", "Abonado Gerencia Loja",
        "Confraternização universal", "Aniversario de São Paulo"
    ]
    
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
            
            for afastamento in afastamentos_impeditivos:
                if afastamento.lower() in afastamentos:
                    tem_afastamento_impeditivo = True
                    break
            
            if not tem_afastamento_impeditivo:
                for afastamento in afastamentos_decisao:
                    if afastamento.lower() in afastamentos:
                        tem_afastamento_decisao = True
                        break
            
            tem_apenas_permitidos = not tem_afastamento_impeditivo and not tem_afastamento_decisao
        
        valor_premio = 0
        if func['Qtd_Horas_Mensais'] == 220:
            valor_premio = 300.00
        elif func['Qtd_Horas_Mensais'] <= 120:
            valor_premio = 150.00
        
        status = "Não tem direito"
        total_atrasos = ""
        
        if not tem_afastamento_impeditivo:
            if tem_afastamento_decisao:
                status = "Aguardando decisão"
                total_atrasos = ausencias['Atrasos'].iloc[0] if not ausencias.empty else ""
            elif tem_apenas_permitidos or ausencias.empty:
                status = "Tem direito"
        
        resultados.append({
            'Matricula': func['Matricula'],
            'Nome': func['Nome_Funcionario'],
            'Cargo': func['Cargo'],
            'Local': func['Nome_Local'],
            'Horas_Mensais': func['Qtd_Horas_Mensais'],
            'Data_Admissao': func['Data_Admissao'],
            'Valor_Premio': valor_premio if status == "Tem direito" else 0,
            'Status': f"{status} (Total Atrasos: {total_atrasos})" if status == "Aguardando decisão" and total_atrasos else status,
            'Detalhes_Afastamentos': ausencias['Afastamentos'].iloc[0] if not ausencias.empty else '',
            'Observações': ''
        })
    
    return pd.DataFrame(resultados)

def exportar_excel(df_mostrar, df_funcionarios):
    output = io.BytesIO()
    df_export = df_mostrar.copy()
    df_export['Salario'] = df_funcionarios.set_index('Matricula').loc[df_export['Matricula'], 'Salario_Mes_Atual'].values
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Resultados Detalhados')
        
        relatorio_diretoria = pd.DataFrame([
            ["RELATÓRIO DE PRÊMIOS - VISÃO EXECUTIVA", ""],
            [f"Data do relatório: {datetime.now().strftime('%d/%m/%Y')}", ""],
            ["", ""],
            ["RESUMO GERAL", ""],
            [f"Total de Funcionários Analisados: {len(df_export)}", ""],
            [f"Funcionários com Direito: {len(df_export[df_export['Status'] == 'Tem direito'])}", ""],
            [f"Funcionários Aguardando Decisão: {len(df_export[df_export['Status'].str.contains('Aguardando decisão', na=False)])}", ""],
            [f"Valor Total dos Prêmios: R$ {df_export['Valor_Premio'].sum():,.2f}", ""],
            ["", ""],
            ["DETALHAMENTO POR STATUS", ""],
        ])
        
        for status in df_export['Status'].unique():
            df_status = df_export[df_export['Status'] == status]
            relatorio_diretoria = pd.concat([relatorio_diretoria, pd.DataFrame([
                [f"\nStatus: {status}", ""],
                [f"Quantidade de Funcionários: {len(df_status)}", ""],
                [f"Valor Total: R$ {df_status['Valor_Premio'].sum():,.2f}", ""],
                ["Locais Afetados:", ""],
                [", ".join(df_status['Local'].unique()), ""],
                ["", ""]
            ])])
        
        relatorio_diretoria.to_excel(writer, index=False, header=False, sheet_name='Relatório Executivo')
    
    return output.getvalue()

def main():
    st.set_page_config(page_title="Sistema de Verificação de Prêmios", page_icon="🏆", layout="wide")
    st.title("Sistema de Verificação de Prêmios")
    
    with st.sidebar:
        st.header("Configurações")
        
        data_limite = st.date_input(
            "Data Limite de Admissão",
             help="Funcionários admitidos após esta data não terão direito ao prêmio",
            value=datetime.now(),
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
            df_funcionarios = pd.read_excel(uploaded_func)
            df_funcionarios.columns = [
                "Matricula", "Nome_Funcionario", "Cargo", 
                "Codigo_Local", "Nome_Local", "Qtd_Horas_Mensais",
                "Tipo_Contrato", "Data_Termino_Contrato", 
                "Dias_Experiencia", "Salario_Mes_Atual", "Data_Admissao"
            ]
            
            df_ausencias = pd.read_excel(uploaded_ausencias)
            df_ausencias = processar_ausencias(df_ausencias)
            
            df_resultado = calcular_premio(df_funcionarios, df_ausencias, data_limite)
            
            st.subheader("Resultado do Cálculo de Prêmios")
            
            df_mostrar = df_resultado
            
            # Editar resultados
            df_mostrar = editar_valores_status(df_mostrar)
            
            # Mostrar métricas
            st.metric("Total de Funcionários com Direito", len(df_mostrar[df_mostrar['Status'] == "Tem direito"]))
            st.metric("Total de Funcionários sem Direito", len(df_mostrar[df_mostrar['Status'] == "Não tem direito"]))
            st.metric("Valor Total dos Prêmios", f"R$ {df_mostrar['Valor_Premio'].sum():,.2f}")
            
            # Filtros
            status_filter = st.selectbox("Filtrar por Status", options=["Todos", "Tem direito", "Não tem direito", "Aguardando decisão"])
            if status_filter != "Todos":
                df_mostrar = df_mostrar[df_mostrar['Status'] == status_filter]
            
            nome_filter = st.text_input("Filtrar por Nome")
            if nome_filter:
                df_mostrar = df_mostrar[df_mostrar['Nome'].str.contains(nome_filter, case=False)]
            
            # Mostrar tabela de resultados na interface
            st.dataframe(df_mostrar)
            
            # Exportar resultados
            if st.button("Exportar Resultados para Excel"):
                df_exportar = df_mostrar[df_mostrar['Status'] == "Tem direito"].copy()
                df_exportar['CPF'] = ""  # Adicione lógica para preencher CPF
                df_exportar['CNPJ'] = "65035552000180"  # Adicione lógica para preencher CNPJ
                df_exportar = df_exportar.rename(columns={'Valor_Premio': 'SomaDeVALOR'})
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_exportar.to_excel(writer, index=False, sheet_name='Funcionarios com Direito')
                st.download_button("Baixar Excel", output.getvalue(), "funcionarios_com_direito.xlsx")
        
        except Exception as e:
            st.error(f"Erro ao processar dados: {str(e)}")

if __name__ == "__main__":
    main()
