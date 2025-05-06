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
    # Verificar se o diretório 'data' existe e criá-lo se não existir
    if not os.path.exists("data"):
        os.makedirs("data")
        
    if os.path.exists("data/tipos_afastamento.pkl"):
        return pd.read_pickle("data/tipos_afastamento.pkl")
    return pd.DataFrame({"tipo": [], "categoria": []})

def salvar_tipos_afastamento(df):
    # Verificar se o diretório 'data' existe e criá-lo se não existir
    if not os.path.exists("data"):
        os.makedirs("data")
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
    
    # Processar faltas marcadas com X na coluna Falta
    df['Faltas'] = df['Falta'].fillna('')
    df['Faltas'] = df['Faltas'].apply(lambda x: 1 if str(x).upper().strip() == 'X' else 0)
    
    # Detectar faltas não justificadas na coluna Ausência Parcial
    df['Tem_Falta_Nao_Justificada'] = df['Ausencia_Parcial'].fillna('').astype(str).str.contains('Falta não justificada', case=False)
    
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
    
    # Processar informações de atraso na coluna Ausência Parcial
    df['Tem_Atraso'] = df['Ausencia_Parcial'].fillna('').astype(str).str.contains('Atraso', case=False)
    
    # Adicionar tipos de afastamento à coluna Afastamentos quando encontrados na coluna Ausência Parcial
    df['Afastamentos'] = df.apply(
        lambda row: row['Afastamentos'] + '; Atraso' if row['Tem_Atraso'] and 'Atraso' not in str(row['Afastamentos']) 
        else row['Afastamentos'],
        axis=1
    )
    
    # Adicionar Falta não justificada aos afastamentos quando encontrado na coluna Ausência Parcial ou Falta é X
    df['Afastamentos'] = df.apply(
        lambda row: row['Afastamentos'] + '; Falta não justificada' 
        if (row['Tem_Falta_Nao_Justificada'] or row['Faltas'] == 1) and 'Falta não justificada' not in str(row['Afastamentos']) 
        else row['Afastamentos'],
        axis=1
    )
    
    df['Afastamentos'] = df['Afastamentos'].fillna('').astype(str)
    
    # Armazenar os valores de atraso para uso posterior
    df['Atrasos'] = df.apply(
        lambda row: row['Ausencia_Parcial'] if row['Tem_Atraso'] else '',
        axis=1
    )
    
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
        "Licença Maternidade", "Atestado Médico", "Férias", "Feriado", "Falta não justificada"
    ]
    afastamentos_decisao = ["Abono", "Atraso"]
    
    df['Status'] = df['Afastamentos'].apply(classificar_status)
    
    # Retornar DataFrame atualizado
    return df

def calcular_premio(df_funcionarios, df_ausencias, data_limite_admissao):
    # Afastamentos que impedem o recebimento do prêmio (Não tem direito)
    afastamentos_impeditivos = [
        "Declaração Acompanhante", "Feriado", "Emenda Feriado", 
        "Licença Maternidade", "Declaração INSS (dias)", 
        "Comparecimento Medico INSS", "Aposentado por Invalidez",
        "Atestado Médico", "Atestado de Óbito", "Licença Paternidade",
        "Licença Casamento", "Acidente de Trabalho", "Auxilio Doença",
        "Primeira Suspensão", "Segunda Suspensão", "Férias",
        "Falta não justificada", "Processo", "Processo Trabalhista",
        "Falta não justificada (dias)", "Atestado Médico (dias)",
        "Declaração Comparecimento Medico", "INSS", "INSS (dias)",
        "Confraternização universal", "Aniversario de São Paulo"
    ]
    
    # Afastamentos que precisam de decisão (Aguardando decisão)
    afastamentos_decisao = ["Atraso"]
    
    # Afastamentos que permitem receber o prêmio (Tem direito)
    afastamentos_permitidos = [
        "Folga Gestor", "Abonado Gerencia Loja", "Abono Administrativo"
    ]
    
    df_funcionarios['Data_Admissao'] = pd.to_datetime(df_funcionarios['Data_Admissao'], format='%d/%m/%Y')
    df_funcionarios = df_funcionarios[df_funcionarios['Data_Admissao'] <= pd.to_datetime(data_limite_admissao)]
    
    resultados = []
    
    # Agrupar ausências por matrícula para considerar todas as ocorrências juntas
    matriculas_funcionarios = df_funcionarios['Matricula'].unique()
    
    for matricula in matriculas_funcionarios:
        # Buscar dados do funcionário
        func = df_funcionarios[df_funcionarios['Matricula'] == matricula].iloc[0]
        
        # Buscar todas as ausências do funcionário
        ausencias = df_ausencias[df_ausencias['Matricula'] == matricula]
        
        # Inicializar flags
        tem_afastamento_impeditivo = False
        tem_afastamento_decisao = False
        tem_apenas_permitidos = False
        todos_afastamentos = []
        todos_atrasos = []
        
        if not ausencias.empty:
            # Verificar se tem falta não justificada em qualquer linha
            tem_falta_nao_justificada = False
            if 'Tem_Falta_Nao_Justificada' in ausencias.columns:
                tem_falta_nao_justificada = ausencias['Tem_Falta_Nao_Justificada'].any()
            
            # Verificar se tem X na coluna Falta em qualquer linha
            tem_falta_marcada = ausencias['Faltas'].sum() > 0 if 'Faltas' in ausencias.columns else False
            
            # Verificar se em qualquer linha tem "Falta não justificada" na Ausência Parcial
            tem_falta_na_ausencia_parcial = False
            if 'Ausencia_Parcial' in ausencias.columns:
                tem_falta_na_ausencia_parcial = ausencias['Ausencia_Parcial'].fillna('').astype(str).str.contains('Falta não justificada', case=False).any()
            
            # Se tem qualquer tipo de falta, é impeditivo
            if tem_falta_nao_justificada or tem_falta_marcada or tem_falta_na_ausencia_parcial:
                tem_afastamento_impeditivo = True
                todos_afastamentos.append("Falta não justificada")
            
            # Verificar todos os outros tipos de afastamento em todas as linhas
            for _, linha in ausencias.iterrows():
                # Juntar todas as fontes possíveis de afastamento
                afastamentos_linha = str(linha.get('Afastamentos', '')).lower() 
                detalhes_linha = str(linha.get('Detalhes_Afastamentos', '')).lower()
                ausencia_parcial_linha = str(linha.get('Ausencia_Parcial', '')).lower()
                ausencia_integral_linha = str(linha.get('Ausencia_Integral', '')).lower()
                
                # Verificar tipos específicos de afastamento nesta linha
                
                # Verificar atrasos
                if ('atraso' in afastamentos_linha or 'atraso' in ausencia_parcial_linha):
                    tem_afastamento_decisao = True
                    if 'atraso' not in [a.lower() for a in todos_afastamentos]:
                        todos_afastamentos.append("Atraso")
                    # Guardar os detalhes do atraso para mostrar depois
                    if 'Ausencia_Parcial' in linha and 'Atraso' in str(linha['Ausencia_Parcial']):
                        todos_atrasos.append(str(linha['Ausencia_Parcial']))
                
                # Verificar impeditivos (um por um para saber qual foi encontrado)
                for afastamento in afastamentos_impeditivos:
                    af_lower = afastamento.lower()
                    if (af_lower in afastamentos_linha or 
                        af_lower in detalhes_linha or 
                        af_lower in ausencia_parcial_linha or
                        af_lower in ausencia_integral_linha):
                        tem_afastamento_impeditivo = True
                        if afastamento not in todos_afastamentos:
                            todos_afastamentos.append(afastamento)
                
                # Verificar permitidos
                for afastamento in afastamentos_permitidos:
                    af_lower = afastamento.lower()
                    if (af_lower in afastamentos_linha or 
                        af_lower in detalhes_linha or 
                        af_lower in ausencia_parcial_linha or
                        af_lower in ausencia_integral_linha):
                        if afastamento not in todos_afastamentos:
                            todos_afastamentos.append(afastamento)
        
        # Determinar se tem apenas afastamentos permitidos
        if todos_afastamentos:
            tem_apenas_permitidos = True
            for afastamento in todos_afastamentos:
                if (afastamento.lower() in [a.lower() for a in afastamentos_impeditivos] or 
                    afastamento.lower() in [a.lower() for a in afastamentos_decisao]):
                    tem_apenas_permitidos = False
                    break
        
        # Definir valor do prêmio com base nas horas mensais
        valor_premio = 0
        if func['Qtd_Horas_Mensais'] == 220:
            valor_premio = 300.00
        elif func['Qtd_Horas_Mensais'] <= 120:
            valor_premio = 150.00
        
        # Definir status padrão
        status = "Não tem direito"
        total_atrasos = ""
        
        # Determinar status com base nos afastamentos
        if tem_afastamento_impeditivo:
            status = "Não tem direito"
        elif tem_afastamento_decisao:
            status = "Aguardando decisão"
            total_atrasos = "; ".join(todos_atrasos) if todos_atrasos else ""
        elif tem_apenas_permitidos or ausencias.empty:
            status = "Tem direito"
        
        # Criar o dicionário de resultado
        resultado = {
            'Matricula': func['Matricula'],
            'Nome': func['Nome_Funcionario'],
            'Cargo': func['Cargo'],
            'Local': func['Nome_Local'],
            'Horas_Mensais': func['Qtd_Horas_Mensais'],
            'Data_Admissao': func['Data_Admissao'],
            'Valor_Premio': valor_premio if status == "Tem direito" else 0,
            'Status': f"{status} (Total Atrasos: {total_atrasos})" if status == "Aguardando decisão" and total_atrasos else status,
            'Detalhes_Afastamentos': "; ".join(todos_afastamentos) if todos_afastamentos else '',
            'Observações': ''
        }
        
        resultados.append(resultado)
    
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
                # Verificar se as colunas do arquivo carregado estão corretas
                if 'tipo de afastamento' in df_tipos_novo.columns and 'Direito Pagamento' in df_tipos_novo.columns:
                    # Renomear as colunas para os nomes esperados pelo sistema
                    df_tipos = df_tipos_novo.rename(columns={'tipo de afastamento': 'tipo', 'Direito Pagamento': 'categoria'})
                    salvar_tipos_afastamento(df_tipos)
                    st.success("Tipos de afastamento atualizados!")
                else:
                    st.error("Arquivo deve conter colunas 'tipo de afastamento' e 'Direito Pagamento'")
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
            
            # Verificar e exibir afastamentos desconhecidos
            if not df_ausencias['Afastamentos_Desconhecidos'].str.strip().eq('').all():
                st.warning("Foram encontrados afastamentos desconhecidos na tabela de ausências:")
                st.dataframe(df_ausencias[['Matricula', 'Afastamentos_Desconhecidos']])
                st.info("Atualize os tipos de afastamento para corrigir essas inconsistências.")
            
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
