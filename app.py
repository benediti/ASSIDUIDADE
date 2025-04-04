import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
from io import BytesIO
import os

# Ignora avisos do pandas
warnings.filterwarnings('ignore')

st.title("Processador de Ausências")
st.write("Este aplicativo processa dados de ausências de funcionários.")

def converter_data_br_para_datetime(data_str):
    """
    Converte uma string de data no formato brasileiro (DD/MM/YYYY) para um objeto datetime.
    Retorna None se a conversão falhar.
    """
    if pd.isna(data_str) or data_str == '' or data_str is None:
        return None
    if isinstance(data_str, datetime):
        return data_str
    try:
        if isinstance(data_str, str):
            return datetime.strptime(data_str, '%d/%m/%Y')
        else:
            return data_str
    except (ValueError, TypeError):
        try:
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y']:
                try:
                    return datetime.strptime(str(data_str), fmt)
                except ValueError:
                    continue
            return None
        except Exception:
            return None

def carregar_arquivo_ausencias(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [col.strip() for col in df.columns]

        # Converter coluna "Dia" para datetime
        if 'Dia' in df.columns:
            df['Dia'] = df['Dia'].astype(str)
            df['Dia'] = df['Dia'].apply(converter_data_br_para_datetime)

        # Converter coluna "Data de Demissão" para datetime
        if 'Data de Demissão' in df.columns:
            df['Data de Demissão'] = df['Data de Demissão'].astype(str)
            df['Data de Demissão'] = df['Data de Demissão'].apply(converter_data_br_para_datetime)

        # Converter "Ausência Integral" e "Ausência Parcial" para horas decimais (ex: 01:30 → 1.5 horas)
        for col in ['Ausência Integral', 'Ausência Parcial']:
            if col in df.columns:
                try:
                    df[col] = pd.to_timedelta(df[col], errors='coerce').dt.total_seconds() / 3600
                except Exception as e:
                    st.warning(f"Erro ao converter a coluna '{col}' para horas decimais: {e}")

        return df

    except Exception as e:
        st.error(f"Erro ao carregar arquivo de ausências: {e}")
        return pd.DataFrame()

def carregar_arquivo_funcionarios(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [col.strip() for col in df.columns]
        if 'Salário Mês Atual' in df.columns:
            st.write("### Informações sobre a coluna de salário")
            st.write(f"Tipo da coluna: {df['Salário Mês Atual'].dtype}")
            st.write("Amostra de valores de salário:")
            amostra = df['Salário Mês Atual'].head(5).tolist()
            for i, valor in enumerate(amostra):
                st.write(f"Valor {i+1}: {valor} (tipo: {type(valor)})")
        if 'Data Término Contrato' in df.columns:
            df['Data Término Contrato'] = df['Data Término Contrato'].astype(str)
            df['Data Término Contrato'] = df['Data Término Contrato'].apply(converter_data_br_para_datetime)
        if 'Data Admissão' in df.columns:
            df['Data Admissão'] = df['Data Admissão'].astype(str)
            df['Data Admissão'] = df['Data Admissão'].apply(converter_data_br_para_datetime)
        return df
    except Exception as e:
        st.error(f"Erro ao carregar arquivo de funcionários: {e}")
        return pd.DataFrame()

def carregar_arquivo_afastamentos(uploaded_file):
    if uploaded_file is None:
        return pd.DataFrame()
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [col.strip() for col in df.columns]
        return df
    except Exception as e:
        st.error(f"Erro ao carregar arquivo de afastamentos: {e}")
        return pd.DataFrame()

def consolidar_dados_funcionario(df_combinado):
    if df_combinado.empty:
        return df_combinado
    dados_consolidados = []
    for matricula, grupo in df_combinado.groupby('Matricula'):
        funcionario = grupo.iloc[0].copy()
        tem_falta = False
        if 'Falta' in df_combinado.columns:
            tem_falta = grupo['Falta'].notna().any()
        tem_afastamento = False
        if 'Afastamentos' in df_combinado.columns:
            tem_afastamento = grupo['Afastamentos'].notna().any()
        tem_ausencia = False
        if ('Ausência Integral' in df_combinado.columns) and ('Ausência Parcial' in df_combinado.columns):
            tem_ausencia = grupo['Ausência Integral'].notna().any() or grupo['Ausência Parcial'].notna().any()
        funcionario['Tem Falta'] = tem_falta
        funcionario['Tem Afastamento'] = tem_afastamento
        funcionario['Tem Ausência'] = tem_ausencia
        dados_consolidados.append(funcionario)
    df_consolidado = pd.DataFrame(dados_consolidados)
    return df_consolidado

def processar_dados(df_ausencias, df_funcionarios, df_afastamentos, data_limite_admissao=None):
    st.write("Chegou aqui: Início do processamento")
    st.write("Data limite de admissão recebida:", data_limite_admissao)
    if df_ausencias.empty or df_funcionarios.empty:
        st.warning("Um ou mais arquivos não puderam ser carregados corretamente.")
        return pd.DataFrame()
    if data_limite_admissao:
        data_limite = converter_data_br_para_datetime(data_limite_admissao)
    else:
        data_limite = None
    if data_limite is not None:
        df_funcionarios = df_funcionarios[df_funcionarios['Data Admissão'] <= data_limite]
    if 'Matricula' in df_ausencias.columns and 'Matricula' in df_funcionarios.columns:
        df_ausencias['Matricula'] = df_ausencias['Matricula'].astype(str)
        df_funcionarios['Matricula'] = df_funcionarios['Matricula'].astype(str)
        df_combinado = pd.merge(
            df_funcionarios,
            df_ausencias,
            left_on='Matricula',
            right_on='Matricula',
            how='left'
        )
        # Remover colunas duplicadas após o merge
        df_combinado = df_combinado.loc[:, ~df_combinado.columns.duplicated()]
        if not df_afastamentos.empty:
            if 'Matricula' in df_afastamentos.columns:
                df_afastamentos['Matricula'] = df_afastamentos['Matricula'].astype(str)
                df_combinado = pd.merge(
                    df_combinado,
                    df_afastamentos,
                    left_on='Matricula',
                    right_on='Matricula',
                    how='left',
                    suffixes=('', '_afastamento')
                )
                df_combinado = df_combinado.loc[:, ~df_combinado.columns.duplicated()]
        df_consolidado = consolidar_dados_funcionario(df_combinado)
        df_final = aplicar_regras_pagamento(df_consolidado)
        st.write("Chegou aqui: Fim do processamento")
        return df_final
    else:
        st.warning("Aviso: Colunas de Matricula não encontradas em um ou ambos os arquivos.")
        return pd.DataFrame()

def aplicar_regras_pagamento(df):
    df['Valor a Pagar'] = 0.0
    df['Status'] = ''
    df['Cor'] = ''
    df['Observacoes'] = ''
    funcoes_sem_direito = [
        'AUX DE SERV GERAIS (INT)', 
        'AUX DE LIMPEZA (INT)',
        'LIMPADOR DE VIDROS INT', 
        'RECEPCIONISTA INTERMITENTE', 
        'PORTEIRO INTERMITENTE'
    ]
    afastamentos_com_direito = [
        'Abonado Gerencia Loja', 
        'Abono Administrativo'
    ]
    afastamentos_aguardar_decisao = [
        'Atraso'
    ]
    afastamentos_sem_direito = [
        'Atestado Médico', 
        'Atestado de Óbito', 
        'Folga Gestor', 
        'Licença Paternidade',
        'Licença Casamento', 
        'Acidente de Trabalho', 
        'Auxilio Doença', 
        'Primeira Suspensão', 
        'Segunda Suspensão', 
        'Férias', 
        'Abono Atraso', 
        'Falta não justificada', 
        'Processo Atraso', 
        'Confraternização universal', 
        'Atestado Médico (dias)', 
        'Declaração Comparecimento Medico',
        'Processo Trabalhista', 
        'Licença Maternidade'
    ]
    for idx, row in df.iterrows():
        cargo = str(row.get('Cargo', '')).strip()
        if cargo in funcoes_sem_direito:
            df.at[idx, 'Valor a Pagar'] = 0.0
            df.at[idx, 'Status'] = 'Não tem direito'
            df.at[idx, 'Cor'] = 'vermelho'
            df.at[idx, 'Observacoes'] = f'Cargo sem direito: {cargo}'
            continue
        salario = 0
        salario_original = row.get('Salário Mês Atual', 0)
        try:
            if isinstance(salario_original, (int, float)) and not pd.isna(salario_original):
                salario = float(salario_original)
            elif salario_original and not pd.isna(salario_original):
                salario_str = str(salario_original)
                salario_limpo = ''.join(c for c in salario_str if c.isdigit() or c in '.,')
                if ',' in salario_limpo and '.' not in salario_limpo:
                    salario_limpo = salario_limpo.replace(',', '.')
                if salario_limpo:
                    salario = float(salario_limpo)
        except Exception:
            salario = 0
        if salario >= 2542.86:
            df.at[idx, 'Valor a Pagar'] = 0.0
            df.at[idx, 'Status'] = 'Não tem direito'
            df.at[idx, 'Cor'] = 'vermelho'
            df.at[idx, 'Observacoes'] = f'Salário acima do limite: R$ {salario:.2f}'
            continue
        if 'Falta' in df.columns and not pd.isna(row.get('Falta')):
            df.at[idx, 'Valor a Pagar'] = 0.0
            df.at[idx, 'Status'] = 'Não tem direito'
            df.at[idx, 'Cor'] = 'vermelho'
            df.at[idx, 'Observacoes'] = 'Tem falta'
            continue
        if 'Afastamentos' in df.columns and not pd.isna(row.get('Afastamentos')):
            tipo_afastamento = str(row.get('Afastamentos')).strip()
            if tipo_afastamento in afastamentos_sem_direito:
                df.at[idx, 'Valor a Pagar'] = 0.0
                df.at[idx, 'Status'] = 'Não tem direito'
                df.at[idx, 'Cor'] = 'vermelho'
                df.at[idx, 'Observacoes'] = f'Afastamento sem direito: {tipo_afastamento}'
                continue
            elif tipo_afastamento in afastamentos_aguardar_decisao:
                df.at[idx, 'Valor a Pagar'] = 0.0
                df.at[idx, 'Status'] = 'Aguardando decisão'
                df.at[idx, 'Cor'] = 'azul'
                df.at[idx, 'Observacoes'] = f'Afastamento para avaliar: {tipo_afastamento}'
                continue
            elif tipo_afastamento in afastamentos_com_direito:
                pass
            else:
                df.at[idx, 'Valor a Pagar'] = 0.0
                df.at[idx, 'Status'] = 'Aguardando decisão'
                df.at[idx, 'Cor'] = 'azul'
                df.at[idx, 'Observacoes'] = f'Afastamento não classificado: {tipo_afastamento}'
                continue
        if (('Ausência Integral' in df.columns and not pd.isna(row.get('Ausência Integral'))) or 
            ('Ausência Parcial' in df.columns and not pd.isna(row.get('Ausência Parcial')))):
            df.at[idx, 'Valor a Pagar'] = 0.0
            df.at[idx, 'Status'] = 'Aguardando decisão'
            df.at[idx, 'Cor'] = 'azul'
            df.at[idx, 'Observacoes'] = 'Tem ausência - Necessita avaliação'
            continue
        horas = 0
        if 'Qtd Horas Mensais' in df.columns:
            try:
                qtd_horas = row.get('Qtd Horas Mensais', 0)
                if isinstance(qtd_horas, (int, float)) and not pd.isna(qtd_horas):
                    horas = float(qtd_horas)
                elif qtd_horas and not pd.isna(qtd_horas):
                    if str(qtd_horas).isdigit():
                        horas = float(qtd_horas)
            except Exception:
                horas = 0
        if horas == 220:
            df.at[idx, 'Valor a Pagar'] = 300.00
            df.at[idx, 'Status'] = 'Tem direito'
            df.at[idx, 'Cor'] = 'verde'
            df.at[idx, 'Observacoes'] = 'Sem ausências - 220 horas'
        elif horas in [110, 120]:
            df.at[idx, 'Valor a Pagar'] = 150.00
            df.at[idx, 'Status'] = 'Tem direito'
            df.at[idx, 'Cor'] = 'verde'
            df.at[idx, 'Observacoes'] = f'Sem ausências - {int(horas)} horas'
        else:
            df.at[idx, 'Valor a Pagar'] = 0.00
            df.at[idx, 'Status'] = 'Aguardando decisão'
            df.at[idx, 'Cor'] = 'azul'
            df.at[idx, 'Observacoes'] = f'Sem ausências - verificar horas: {horas}'
    df = df.rename(columns={
        'Matricula': 'Matricula',
        'Nome Funcionário': 'Nome',
        'Nome Local': 'Local',
        'Valor a Pagar': 'Valor_Premio'
    })
    return df

def exportar_novo_excel(df):
    output = BytesIO()
    colunas_adicionais = []
    for coluna in ['Cargo', 'Salário Mês Atual', 'Qtd Horas Mensais', 'Falta', 'Afastamentos', 'Ausência Integral', 'Ausência Parcial']:
        if coluna in df.columns:
            colunas_adicionais.append(coluna)
    colunas_principais = ['Matricula', 'Nome', 'Local', 'Valor_Premio', 'Observacoes']
    colunas_exportar = colunas_adicionais + colunas_principais
    df_tem_direito = df[df['Status'] == 'Tem direito'][colunas_exportar]
    df_nao_tem_direito = df[df['Status'] == 'Não tem direito'][colunas_exportar]
    df_aguardando = df[df['Status'] == 'Aguardando decisão'][colunas_exportar]
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_tem_direito.to_excel(writer, index=False, sheet_name='Com Direito')
        workbook = writer.book
        worksheet = writer.sheets['Com Direito']
        format_green = workbook.add_format({'bg_color': '#CCFFCC'})
        for i in range(len(df_tem_direito)):
            for j in range(len(colunas_exportar)):
                worksheet.write(i+1, j, str(df_tem_direito.iloc[i][colunas_exportar[j]]), format_green)
        df_nao_tem_direito.to_excel(writer, index=False, sheet_name='Sem Direito')
        worksheet = writer.sheets['Sem Direito']
        format_red = workbook.add_format({'bg_color': '#FFCCCC'})
        for i in range(len(df_nao_tem_direito)):
            for j in range(len(colunas_exportar)):
                worksheet.write(i+1, j, str(df_nao_tem_direito.iloc[i][colunas_exportar[j]]), format_red)
        df_aguardando.to_excel(writer, index=False, sheet_name='Aguardando Decisão')
        worksheet = writer.sheets['Aguardando Decisão']
        format_blue = workbook.add_format({'bg_color': '#CCCCFF'})
        for i in range(len(df_aguardando)):
            for j in range(len(colunas_exportar)):
                worksheet.write(i+1, j, str(df_aguardando.iloc[i][colunas_exportar[j]]), format_blue)
        resumo_data = [
            ['RESUMO DO PROCESSAMENTO'],
            [f'Data de Geração: {pd.Timestamp.now().strftime("%d/%m/%Y %H:%M:%S")}'],
            [''],
            ['Métricas Gerais'],
            [f'Total de Funcionários Processados: {len(df)}'],
            [f'Total de Funcionários com Direito: {len(df_tem_direito)}'],
            [f'Total de Funcionários sem Direito: {len(df_nao_tem_direito)}'],
            [f'Total de Funcionários Aguardando Decisão: {len(df_aguardando)}'],
            [f'Valor Total dos Prêmios: R$ {df_tem_direito["Valor_Premio"].sum():,.2f}'],
        ]
        pd.DataFrame(resumo_data).to_excel(writer, index=False, header=False, sheet_name='Resumo')
    output.seek(0)
    return output.getvalue()

# Sidebar - Upload de Arquivos
st.sidebar.header("Upload de Arquivos")
arquivo_ausencias = st.sidebar.file_uploader("Arquivo de Ausências", type=["xlsx", "xls"])
arquivo_funcionarios = st.sidebar.file_uploader("Arquivo de Funcionários", type=["xlsx", "xls"])
arquivo_afastamentos = st.sidebar.file_uploader("Arquivo de Afastamentos (opcional)", type=["xlsx", "xls"])

# Verificar se os arquivos foram carregados
if arquivo_ausencias is not None:
    st.sidebar.write("Arquivo de ausências carregado com sucesso.")
else:
    st.sidebar.error("Erro ao carregar o arquivo de ausências.")

if arquivo_funcionarios is not None:
    st.sidebar.write("Arquivo de funcionários carregado com sucesso.")
else:
    st.sidebar.error("Erro ao carregar o arquivo de funcionários.")

st.sidebar.header("Data Limite de Admissão")
data_limite = st.sidebar.date_input(
    "Considerar apenas funcionários admitidos até:",
    value=datetime(2025, 3, 1)
)

# Tabs para Processamento e Exportação
tab1, tab2 = st.tabs(["Processamento Inicial", "Edição e Exportação"])

with tab1:
    processar = st.button("Processar Dados")
    if processar:
        st.write("Chegou aqui: Início da etapa de processamento")
        if arquivo_ausencias is not None and arquivo_funcionarios is not None:
            with st.spinner("Processando arquivos..."):
                df_ausencias = carregar_arquivo_ausencias(arquivo_ausencias)
                df_funcionarios = carregar_arquivo_funcionarios(arquivo_funcionarios)
                df_afastamentos = carregar_arquivo_afastamentos(arquivo_afastamentos)
                
                # Verificar as colunas dos DataFrames
                if not df_ausencias.empty:
                    st.write("Colunas do arquivo de ausências:", df_ausencias.columns.tolist())
                if not df_funcionarios.empty:
                    st.write("Colunas do arquivo de funcionários:", df_funcionarios.columns.tolist())
                
                data_limite_str = data_limite.strftime("%d/%m/%Y")
                st.write("Data limite de admissão:", data_limite_str)
                resultado = processar_dados(df_ausencias, df_funcionarios, df_afastamentos, data_limite_admissao=data_limite_str)
                if not resultado.empty:
                    st.success(f"Processamento concluído com sucesso. {len(resultado)} registros encontrados.")
                    df_display = resultado.copy()
                    for col in df_display.columns:
                        df_display[col] = df_display[col].astype(str)
                    st.write("Primeiras linhas do DataFrame:")
                    st.write(df_display.head())
                    st.write("Tipos de dados das colunas:")
                    st.write(df_display.dtypes)
                    st.write("Distribuição dos status:", df_display['Status'].value_counts())
                    st.subheader("Resultados Preliminares")
                    st.dataframe(df_display)
                    total_a_pagar = resultado['Valor_Premio'].sum()
                    contagem_por_status = resultado['Status'].value_counts()
                    st.subheader("Resumo")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Total a Pagar", f"R$ {total_a_pagar:.2f}")
                    with col2:
                        st.write("Contagem por Status:")
                        st.write(contagem_por_status)
                    st.info("Vá para a aba 'Edição e Exportação' para ajustar os valores e exportar o resultado final.")
                else:
                    st.warning("Nenhum resultado encontrado.")
        else:
            st.warning("Por favor, faça o upload dos arquivos necessários.")

with tab2:
    if 'resultado_processado' in st.session_state:
        st.dataframe(st.session_state.resultado_processado)
    else:
        st.info("Por favor, primeiro processe os dados na aba 'Processamento Inicial'.")

st.sidebar.markdown("---")
st.sidebar.info("""
**Como usar:**
1. Faça o upload dos arquivos necessários.
2. Defina a data limite de admissão.
3. Na aba 'Processamento Inicial', clique em 'Processar Dados'.
4. Revise os resultados e vá para 'Edição e Exportação' para ajustar valores e exportar.
5. O relatório final conterá abas para:
   - Funcionários com Direito
   - Funcionários sem Direito
   - Funcionários Aguardando Decisão
   - Resumo do processamento
""")
