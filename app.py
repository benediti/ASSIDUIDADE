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

def renomear_colunas_duplicadas(df):
    """
    Garante que os nomes das colunas sejam únicos ao adicionar sufixos aos duplicados.
    """
    # Verificar se há colunas duplicadas
    if not df.columns.duplicated().any():
        return df
    
    # Cria um novo DataFrame com colunas renomeadas
    df_renamed = df.copy()
    new_cols = []
    seen = {}
    
    for col in df_renamed.columns:
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
    
    df_renamed.columns = new_cols
    return df_renamed

def debug_dataframe(df):
    """
    Exibe informações de depuração sobre o DataFrame de forma segura.
    """
    try:
        # Verifica e renomeia colunas duplicadas
        if df.columns.duplicated().any():
            st.warning("Colunas duplicadas detectadas e renomeadas automaticamente.")
            df = renomear_colunas_duplicadas(df)
            
        st.write("DataFrame Head:")
        df_head = df.head().copy()
        for col in df_head.columns:
            df_head[col] = df_head[col].astype(str)
        st.dataframe(df_head)
        
        st.write("DataFrame Info:")
        info_data = []
        for col in df.columns:
            info_data.append({"Column": col, "Type": str(df[col].dtype)})
        st.dataframe(pd.DataFrame(info_data))
        
        st.write("DataFrame Describe:")
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
    """
    Carrega e processa o arquivo de ausências, tratando colunas duplicadas.
    """
    try:
        df = pd.read_excel(uploaded_file)
        # Verificar e corrigir colunas duplicadas
        df = renomear_colunas_duplicadas(df)
        # Limpa espaços em branco nos nomes das colunas
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
    """
    Carrega e processa o arquivo de funcionários, tratando colunas duplicadas.
    """
    try:
        df = pd.read_excel(uploaded_file)
        # Verificar e corrigir colunas duplicadas
        df = renomear_colunas_duplicadas(df)
        # Limpa espaços em branco nos nomes das colunas
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
    """
    Carrega e processa o arquivo de afastamentos, tratando colunas duplicadas.
    """
    if uploaded_file is None:
        return pd.DataFrame()
    try:
        df = pd.read_excel(uploaded_file)
        # Verificar e corrigir colunas duplicadas
        df = renomear_colunas_duplicadas(df)
        # Limpa espaços em branco nos nomes das colunas
        df.columns = [col.strip() for col in df.columns]
        return df
    except Exception as e:
        st.error(f"Erro ao carregar arquivo de afastamentos: {e}")
        return pd.DataFrame()

def consolidar_dados_funcionario(df_combinado):
    """
    Consolida dados por funcionário, tratando possíveis colunas duplicadas.
    """
    if df_combinado.empty:
        return df_combinado
    
    # Verificar e corrigir colunas duplicadas primeiro
    df_combinado = renomear_colunas_duplicadas(df_combinado)
    
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
    """
    Processa os dados de ausências, funcionários e afastamentos, tratando colunas duplicadas em todas as etapas.
    """
    if df_ausencias.empty or df_funcionarios.empty:
        st.warning("Um ou mais arquivos não puderam ser carregados corretamente.")
        return pd.DataFrame()
    
    # Converter data limite
    if data_limite_admissao:
        data_limite = converter_data_br_para_datetime(data_limite_admissao)
    else:
        data_limite = None
        
    # Filtrar funcionários pela data de admissão
    if data_limite is not None:
        df_funcionarios = df_funcionarios[df_funcionarios['Data Admissão'] <= data_limite]
    
    if 'Matricula' in df_ausencias.columns and 'Matricula' in df_funcionarios.columns:
        # Garantir que não há colunas duplicadas em cada DataFrame antes de mesclar
        df_ausencias = renomear_colunas_duplicadas(df_ausencias)
        df_funcionarios = renomear_colunas_duplicadas(df_funcionarios)
        
        # Criar cópias para manter os originais intactos
        df_ausencias_renamed = df_ausencias.copy()
        df_funcionarios_renamed = df_funcionarios.copy()
        
        # Adicionar sufixo às colunas em df_ausencias para evitar nomes duplicados após mesclagem
        rename_map = {}
        for col in df_ausencias_renamed.columns:
            if col in df_funcionarios_renamed.columns and col != 'Matricula':
                rename_map[col] = f"{col}_ausencia"
        
        if rename_map:
            df_ausencias_renamed = df_ausencias_renamed.rename(columns=rename_map)
        
        # Converter colunas Matricula para string para mesclagem
        df_ausencias_renamed['Matricula'] = df_ausencias_renamed['Matricula'].astype(str)
        df_funcionarios_renamed['Matricula'] = df_funcionarios_renamed['Matricula'].astype(str)
        
        # Mesclar DataFrames
        df_combinado = pd.merge(
            df_funcionarios_renamed,
            df_ausencias_renamed,
            left_on='Matricula',
            right_on='Matricula',
            how='left'
        )
        
        # Verificar e renomear colunas duplicadas após a primeira mesclagem
        df_combinado = renomear_colunas_duplicadas(df_combinado)
        
        # Processar afastamentos se existirem
        if not df_afastamentos.empty:
            if 'Matricula' in df_afastamentos.columns:
                # Garantir que não há colunas duplicadas no DataFrame de afastamentos
                df_afastamentos = renomear_colunas_duplicadas(df_afastamentos)
                
                # Renomear colunas para evitar duplicação
                df_afastamentos_renamed = df_afastamentos.copy()
                rename_map = {}
                for col in df_afastamentos_renamed.columns:
                    if col in df_combinado.columns and col != 'Matricula':
                        rename_map[col] = f"{col}_afastamento"
                
                if rename_map:
                    df_afastamentos_renamed = df_afastamentos_renamed.rename(columns=rename_map)
                
                df_afastamentos_renamed['Matricula'] = df_afastamentos_renamed['Matricula'].astype(str)
                
                # Mesclar com DataFrame combinado
                df_combinado = pd.merge(
                    df_combinado,
                    df_afastamentos_renamed,
                    left_on='Matricula',
                    right_on='Matricula',
                    how='left'
                )
                
                # Verificar e renomear colunas duplicadas após a segunda mesclagem
                df_combinado = renomear_colunas_duplicadas(df_combinado)
        
        # Consolidar dados e aplicar regras
        df_consolidado = consolidar_dados_funcionario(df_combinado)
        df_final = aplicar_regras_pagamento(df_consolidado)
        
        # Verificação final para garantir que não há colunas duplicadas
        df_final = renomear_colunas_duplicadas(df_final)
        
        return df_final
    else:
        st.warning("Aviso: Colunas de Matricula não encontradas em um ou ambos os arquivos.")
        return pd.DataFrame()

def aplicar_regras_pagamento(df):
    """
    Aplica regras de pagamento ao DataFrame consolidado.
    """
    # Verificar e renomear colunas duplicadas primeiro
    df = renomear_colunas_duplicadas(df)
    
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
                pass  # Continua o processamento para determinar o valor
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
            
    # Renomear colunas para o formato final
    df = df.rename(columns={
        'Matricula': 'Matricula',
        'Nome Funcionário': 'Nome',
        'Nome Local': 'Local',
        'Valor a Pagar': 'Valor_Premio'
    })
    
    return df

def exportar_novo_excel(df):
    """
    Exporta dados processados para um arquivo Excel com formatação.
    """
    # Verificar e renomear colunas duplicadas antes da exportação
    df = renomear_colunas_duplicadas(df)
    
    output = BytesIO()
    
    # Determinar quais colunas incluir no arquivo exportado
    colunas_adicionais = []
    for coluna in ['Cargo', 'Salário Mês Atual', 'Qtd Horas Mensais', 'Falta', 'Afastamentos', 'Ausência Integral', 'Ausência Parcial']:
        if coluna in df.columns:
            colunas_adicionais.append(coluna)
            
    colunas_principais = ['Matricula', 'Nome', 'Local', 'Valor_Premio', 'Observacoes']
    colunas_exportar = colunas_adicionais + colunas_principais
    
    # Filtrar DataFrame por status
    df_tem_direito = df[df['Status'] == 'Tem direito'][colunas_exportar].copy()
    df_nao_tem_direito = df[df['Status'] == 'Não tem direito'][colunas_exportar].copy()
    df_aguardando = df[df['Status'] == 'Aguardando decisão'][colunas_exportar].copy()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Aba "Com Direito" (verde)
        df_tem_direito.to_excel(writer, index=False, sheet_name='Com Direito')
        workbook = writer.book
        worksheet = writer.sheets['Com Direito']
        format_green = workbook.add_format({'bg_color': '#CCFFCC'})
        for i in range(len(df_tem_direito)):
            for j in range(len(colunas_exportar)):
                worksheet.write(i+1, j, str(df_tem_direito.iloc[i][colunas_exportar[j]]), format_green)
        
        # Aba "Sem Direito" (vermelho)
        df_nao_tem_direito.to_excel(writer, index=False, sheet_name='Sem Direito')
        worksheet = writer.sheets['Sem Direito']
        format_red = workbook.add_format({'bg_color': '#FFCCCC'})
        for i in range(len(df_nao_tem_direito)):
            for j in range(len(colunas_exportar)):
                worksheet.write(i+1, j, str(df_nao_tem_direito.iloc[i][colunas_exportar[j]]), format_red)
        
        # Aba "Aguardando Decisão" (azul)
        df_aguardando.to_excel(writer, index=False, sheet_name='Aguardando Decisão')
        worksheet = writer.sheets['Aguardando Decisão']
        format_blue = workbook.add_format({'bg_color': '#CCCCFF'})
        for i in range(len(df_aguardando)):
            for j in range(len(colunas_exportar)):
                worksheet.write(i+1, j, str(df_aguardando.iloc[i][colunas_exportar[j]]), format_blue)
        
        # Aba "Resumo"
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

# Configuração da interface do Streamlit
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

# Criar abas para separar processamento e edição
tab1, tab2 = st.tabs(["Processamento Inicial", "Edição e Exportação"])

with tab1:
    processar = st.button("Processar Dados")
    if processar:
        if arquivo_ausencias is not None and arquivo_funcionarios is not None:
            with st.spinner("Processando arquivos..."):
                try:
                    # Carregar arquivos
                    df_ausencias = carregar_arquivo_ausencias(arquivo_ausencias)
                    df_funcionarios = carregar_arquivo_funcionarios(arquivo_funcionarios)
                    df_afastamentos = carregar_arquivo_afastamentos(arquivo_afastamentos)
                    
                    # Verificar se os DataFrames estão vazios
                    if df_ausencias.empty or df_funcionarios.empty:
                        st.error("Um ou mais arquivos não puderam ser carregados corretamente.")
                    else:
                        # Processar dados
                        data_limite_str = data_limite.strftime("%d/%m/%Y")
                        resultado = processar_dados(df_ausencias, df_funcionarios, df_afastamentos, data_limite_admissao=data_limite_str)
                        
                        if not resultado.empty:
                            # Verificar colunas duplicadas antes de armazenar o resultado
                            resultado = renomear_colunas_duplicadas(resultado)
                            
                            # Armazenar resultado na session_state
                            st.session_state.resultado_processado = resultado
                            
                            st.success(f"Processamento concluído com sucesso. {len(resultado)} registros encontrados.")
                            
                            # Exibir informações resumidas sobre o resultado
                            st.write("#### Primeiras linhas do DataFrame:")
                            
                            # Criar versão segura para exibição
                            df_display = resultado.copy()
                            for col in df_display.columns:
                                df_display[col] = df_display[col].astype(str)
                            
                            # Exibir tabela de forma segura usando st.dataframe
                            st.dataframe(df_display.head())
                            
                            # Exibir tipos de dados
                            st.write("#### Tipos de Dados das Colunas:")
                            info_data = []
                            for col in resultado.columns:
                                info_data.append({"Coluna": col, "Tipo": str(resultado[col].dtype)})
                            st.dataframe(pd.DataFrame(info_data))
                            
                            # Exibir distribuição de status
                            st.write("#### Distribuição dos Status:")
                            status_counts = resultado['Status'].value_counts().reset_index()
                            status_counts.columns = ['Status', 'Contagem']
                            st.dataframe(status_counts)
                            
                            # Exibir métricas resumidas
                            st.write("#### Resumo:")
                            col1, col2 = st.columns(2)
                            with col1:
                                st.metric("Total a Pagar", f"R$ {resultado['Valor_Premio'].sum():.2f}")
                            with col2:
                                st.metric("Total de Funcionários", str(len(resultado)))
                            
                            # Opção para exportar
                            excel_data = exportar_novo_excel(resultado)
                            st.download_button(
                                label="Baixar Resultados (Excel)",
                                data=excel_data,
                                file_name=f"resultado_ausencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
