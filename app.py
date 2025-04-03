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

def debug_dataframe(df):
    try:
        # Check for duplicate columns first
        duplicate_cols = df.columns[df.columns.duplicated()]
        if not duplicate_cols.empty:
            st.warning(f"Colunas duplicadas detectadas: {list(duplicate_cols)}")
            
            # Create a copy with renamed columns to avoid display issues
            df_clean = df.copy()
            new_cols = []
            seen = set()
            for col in df_clean.columns:
                if col in seen:
                    i = 1
                    new_col = f"{col}_{i}"
                    while new_col in seen:
                        i += 1
                        new_col = f"{col}_{i}"
                    new_cols.append(new_col)
                    seen.add(new_col)
                else:
                    new_cols.append(col)
                    seen.add(col)
            df_clean.columns = new_cols
            df = df_clean
            
        st.write("DataFrame Head:")
        # Convert to string to avoid PyArrow conversion issues
        df_head = df.head().copy()
        for col in df_head.columns:
            df_head[col] = df_head[col].astype(str)
        st.dataframe(df_head)
        
        st.write("DataFrame Info:")
        # Display column types safely
        info_data = []
        for col in df.columns:
            info_data.append({"Column": col, "Type": str(df[col].dtype)})
        st.dataframe(pd.DataFrame(info_data))
        
        st.write("DataFrame Describe:")
        # Try to display statistics, but fallback to basic info if it fails
        try:
            st.write(df.describe(include='all'))
        except Exception as e:
            st.write(f"Unable to generate statistics: {str(e)}")
            st.write(f"DataFrame shape: {df.shape}")
    except Exception as e:
        st.error(f"Error displaying DataFrame debug info: {str(e)}")

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
        if 'Dia' in df.columns:
            df['Dia'] = df['Dia'].astype(str)
            df['Dia'] = df['Dia'].apply(converter_data_br_para_datetime)
        if 'Data de Demissão' in df.columns:
            df['Data de Demissão'] = df['Data de Demissão'].astype(str)
            df['Data de Demissão'] = df['Data de Demissão'].apply(converter_data_br_para_datetime)
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
        # Rename duplicate columns before merging
        df_ausencias_renamed = df_ausencias.copy()
        df_funcionarios_renamed = df_funcionarios.copy()
        
        # Add suffix to columns in df_ausencias to avoid duplicate names after merge
        rename_map = {}
        for col in df_ausencias_renamed.columns:
            if col in df_funcionarios_renamed.columns and col != 'Matricula':
                rename_map[col] = f"{col}_ausencia"
        
        if rename_map:
            df_ausencias_renamed = df_ausencias_renamed.rename(columns=rename_map)
        
        # Convert Matricula columns to string for merging
        df_ausencias_renamed['Matricula'] = df_ausencias_renamed['Matricula'].astype(str)
        df_funcionarios_renamed['Matricula'] = df_funcionarios_renamed['Matricula'].astype(str)
        
        df_combinado = pd.merge(
            df_funcionarios_renamed,
            df_ausencias_renamed,
            left_on='Matricula',
            right_on='Matricula',
            how='left'
        )
        
        if not df_afastamentos.empty:
            if 'Matricula' in df_afastamentos.columns:
                # Rename columns in df_afastamentos to avoid duplicates
                df_afastamentos_renamed = df_afastamentos.copy()
                rename_map = {}
                for col in df_afastamentos_renamed.columns:
                    if col in df_combinado.columns and col != 'Matricula':
                        rename_map[col] = f"{col}_afastamento"
                
                if rename_map:
                    df_afastamentos_renamed = df_afastamentos_renamed.rename(columns=rename_map)
                
                df_afastamentos_renamed['Matricula'] = df_afastamentos_renamed['Matricula'].astype(str)
                
                df_combinado = pd.merge(
                    df_combinado,
                    df_afastamentos_renamed,
                    left_on='Matricula',
                    right_on='Matricula',
                    how='left'
                )
        
        # Check for duplicate columns after all merges
        duplicate_cols = df_combinado.columns[df_combinado.columns.duplicated()]
        if not duplicate_cols.empty:
            st.warning(f"Colunas duplicadas encontradas: {list(duplicate_cols)}")
            # Rename duplicates to make them unique
            new_cols = []
            seen = set()
            for col in df_combinado.columns:
                if col in seen:
                    i = 1
                    new_col = f"{col}_{i}"
                    while new_col in seen:
                        i += 1
                        new_col = f"{col}_{i}"
                    new_cols.append(new_col)
                    seen.add(new_col)
                else:
                    new_cols.append(col)
                    seen.add(col)
            df_combinado.columns = new_cols
        
        df_consolidado = consolidar_dados_funcionario(df_combinado)
        df_final = aplicar_regras_pagamento(df_consolidado)
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
    df_tem_direito = df[df['Status'] == 'Tem direito'][colunas_exportar].copy()
    df_nao_tem_direito = df[df['Status'] == 'Não tem direito'][colunas_exportar].copy()
    df_aguardando = df[df['Status'] == 'Aguardando decisão'][colunas_exportar].copy()
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

st.sidebar.header("Upload de Arquivos")
arquivo_ausencias = st.sidebar.file_uploader("Arquivo de Ausências", type=["xlsx", "xls"])
arquivo_funcionarios = st.sidebar.file_uploader("Arquivo de Funcionários", type=["xlsx", "xls"])
arquivo_afastamentos = st.sidebar.file_uploader("Arquivo de Afastamentos (opcional)", type=["xlsx", "xls"])

st.sidebar.header("Data Limite de Admissão")
data_limite = st.sidebar.date_input(
    "Considerar apenas funcionários admitidos até:",
    value=datetime(2025, 3, 1),
    format="DD/MM/YYYY"
)

tab1, tab2 = st.tabs(["Processamento Inicial", "Edição e Exportação"])

with tab1:
    processar = st.button("Processar Dados")
    if processar:
        if arquivo_ausencias is not None and arquivo_funcionarios is not None:
            with st.spinner("Processando arquivos..."):
                df_ausencias = carregar_arquivo_ausencias(arquivo_ausencias)
                df_funcionarios = carregar_arquivo_funcionarios(arquivo_funcionarios)
                df_afastamentos = carregar_arquivo_afastamentos(arquivo_afastamentos)
                data_limite_str = data_limite.strftime("%d/%m/%Y")
                resultado = processar_dados(df_ausencias, df_funcionarios, df_afastamentos, data_limite_admissao=data_limite_str)
                if not resultado.empty:
                    st.success(f"Processamento concluído com sucesso. {len(resultado)} registros encontrados.")
                    
                    # Store the result in session state for access in the second tab
                    st.session_state.resultado_processado = resultado
                    
                    # Fix: Handle DataFrame display to prevent PyArrow conversion errors
                    try:
                        # First check for duplicate columns and fix them
                        duplicate_cols = resultado.columns[resultado.columns.duplicated()]
                        if not duplicate_cols.empty:
                            st.warning(f"Colunas duplicadas encontradas: {list(duplicate_cols)}")
                            # Rename duplicates
                            new_cols = []
                            seen = set()
                            for col in resultado.columns:
                                if col in seen:
                                    i = 1
                                    new_col = f"{col}_{i}"
                                    while new_col in seen:
                                        i += 1
                                        new_col = f"{col}_{i}"
                                    new_cols.append(new_col)
                                    seen.add(new_col)
                                else:
                                    new_cols.append(col)
                                    seen.add(col)
                            resultado.columns = new_cols
                            st.success("Colunas duplicadas foram renomeadas.")
                        
                        # First convert problematic columns to string to avoid display issues
                        df_display = resultado.copy()
                        
                        # Ensure all columns have proper types before display
                        for col in df_display.columns:
                            # Convert to string to avoid PyArrow conversion issues
                            df_display[col] = df_display[col].astype(str)
                            
                        # Safe display of DataFrame head
                        st.write("Primeiras linhas do DataFrame:")
                        st.dataframe(df_display.head())
                        
                        # Show column data types
                        st.write("Tipos de dados das colunas:")
                        st.write(pd.DataFrame(df_display.dtypes, columns=['Tipo de Dado']))
                        
                        # Show status distribution
                        st.write("Distribuição dos status:")
                        try:
                            status_counts = df_display['Status'].value_counts().reset_index()
                            status_counts.columns = ['Status', 'Quantidade']
                            st.dataframe(status_counts)
                        except Exception as e:
                            st.write(f"Erro ao mostrar distribuição de status: {str(e)}")
                            st.write("Status únicos:", list(df_display['Status'].unique()))
                        
                        # Show complete dataframe with filtering capability
                        st.subheader("Resultados Preliminares")
                        st.dataframe(df_display)
                        
                        # Summary metrics
                        total_a_pagar = resultado['Valor_Premio'].sum()
                        contagem_por_status = resultado['Status'].value_counts()
                        
                        st.subheader("Resumo")
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("Total a Pagar", f"R$ {total_a_pagar:.2f}")
                        with col2:
                            st.write("Contagem por Status:")
                            st.write(contagem_por_status)
                        
                        # Provide export option
                        excel_data = exportar_novo_excel(resultado)
                        st.download_button(
                            label="Baixar Resultados (Excel)",
                            data=excel_data,
                            file_name=f"resultado_ausencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        st.info("Vá para a aba 'Edição e Exportação' para ajustar os valores e exportar o resultado final.")
                    except Exception as e:
                        st.error(f"Erro ao exibir dados: {str(e)}")
                        st.write("Exibindo informações resumidas para evitar erros:")
                        st.write(f"Total de registros: {len(resultado)}")
                        st.write(f"Colunas: {', '.join(resultado.columns.tolist())}")
                else:
                    st.warning("Nenhum resultado encontrado.")
        else:
            st.warning("Por favor, faça o upload dos arquivos necessários.")

with tab2:
    st.subheader("Edição e Exportação de Resultados")
    if 'resultado_processado' in st.session_state and not st.session_state.resultado_processado.empty:
        # Display editable dataframe (in Streamlit this is just for view, not actual editing)
        resultado_edicao = st.session_state.resultado_processado.copy()
        
        # Check for and handle duplicate columns
        duplicate_cols = resultado_edicao.columns[resultado_edicao.columns.duplicated()]
        if not duplicate_cols.empty:
            st.warning(f"Colunas duplicadas encontradas na aba de edição: {list(duplicate_cols)}")
            # Rename duplicates
            new_cols = []
            seen = set()
            for col in resultado_edicao.columns:
                if col in seen:
                    i = 1
                    new_col = f"{col}_{i}"
                    while new_col in seen:
                        i += 1
                        new_col = f"{col}_{i}"
                    new_cols.append(new_col)
                    seen.add(new_col)
                else:
                    new_cols.append(col)
                    seen.add(col)
            resultado_edicao.columns = new_cols
            st.success("Colunas duplicadas foram renomeadas.")
            
        # Create a safe display version - convert everything to string
        safe_display = resultado_edicao.copy()
        for col in safe_display.columns:
            # Convert all columns to string to avoid display issues
            safe_display[col] = safe_display[col].astype(str)
        
        # Filter options
        st.write("Filtrar por Status:")
        try:
            status_values = safe_display['Status'].unique()
            # Check if we actually got values back
            if len(status_values) > 0:
                all_status = ['Todos'] + list(status_values)
            else:
                all_status = ['Todos']
                st.warning("Não foi possível encontrar valores de status distintos.")
            selected_status = st.selectbox("Selecione um status", all_status)
        except Exception as e:
            st.error(f"Erro ao obter valores de status: {str(e)}")
            all_status = ['Todos']
            selected_status = 'Todos'
        
        # Apply filter
        try:
            if selected_status != 'Todos':
                filtered_df = safe_display[safe_display['Status'] == selected_status]
            else:
                filtered_df = safe_display
        except Exception as e:
            st.error(f"Erro ao filtrar dados: {str(e)}")
            filtered_df = safe_display  # Use unfiltered data as fallback
        
        st.dataframe(filtered_df)
        
        # Export the filtered or full dataset
        if st.button("Exportar Dados Filtrados"):
            excel_filtered = exportar_novo_excel(resultado_edicao if selected_status == 'Todos' 
                                               else resultado_edicao[resultado_edicao['Status'] == selected_status])
            st.download_button(
                label="Baixar Arquivo Excel",
                data=excel_filtered,
                file_name=f"resultado_filtrado_{selected_status}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
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
