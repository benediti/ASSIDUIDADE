import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
from io import BytesIO
import os

# Ignore pandas warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="Processador de Ausências", layout="wide")
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
            return datetime.strptime(data_str.strip(), '%d/%m/%Y')
        else:
            return data_str
    except (ValueError, TypeError):
        try:
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y']:
                try:
                    return datetime.strptime(str(data_str).strip(), fmt)
                except ValueError:
                    continue
            return None
        except Exception:
            return None

def converter_valor_monetario(valor):
    """
    Converte um valor monetário em formato brasileiro (1.234,56) para float
    """
    if pd.isna(valor) or valor == '' or valor is None:
        return 0.0
    
    if isinstance(valor, (int, float)):
        return float(valor)
    
    try:
        # Remove caracteres não numéricos, exceto ponto e vírgula
        valor_str = str(valor).strip()
        valor_limpo = ''.join(c for c in valor_str if c.isdigit() or c in '.,')
        
        # Se tiver vírgula e não tiver ponto, converte vírgula para ponto
        if ',' in valor_limpo and '.' not in valor_limpo:
            valor_limpo = valor_limpo.replace(',', '.')
        # Se tiver ponto e vírgula, trata como formato brasileiro (ex: 1.234,56)
        elif ',' in valor_limpo and '.' in valor_limpo:
            valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
            
        if valor_limpo:
            return float(valor_limpo)
        return 0.0
    except Exception as e:
        st.warning(f"Erro ao converter valor monetário: {valor} - {e}")
        return 0.0

def carregar_arquivo_ausencias(uploaded_file):
    try:
        if uploaded_file is None:
            st.warning("Arquivo de ausências não fornecido")
            return pd.DataFrame()
            
        st.info(f"Carregando arquivo de ausências: {uploaded_file.name}")
        df = pd.read_excel(uploaded_file)
        st.write("Formato original do arquivo de ausências:")
        st.write(df.head(2))
        
        # Limpar nomes das colunas
        df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
        
        # Converter coluna de data se existir
        if 'Dia' in df.columns:
            st.write("Amostra de valores da coluna 'Dia' antes da conversão:", df['Dia'].head())
            df['Dia'] = df['Dia'].apply(converter_data_br_para_datetime)
            st.write("Amostra após conversão:", df['Dia'].head())
            
        if 'Data de Demissão' in df.columns:
            df['Data de Demissão'] = df['Data de Demissão'].apply(converter_data_br_para_datetime)
            
        st.success("Arquivo de ausências carregado com sucesso")
        return df
    except Exception as e:
        st.error(f"Erro ao carregar arquivo de ausências: {e}")
        return pd.DataFrame()

def carregar_arquivo_funcionarios(uploaded_file):
    try:
        if uploaded_file is None:
            st.warning("Arquivo de funcionários não fornecido")
            return pd.DataFrame()
            
        st.info(f"Carregando arquivo de funcionários: {uploaded_file.name}")
        df = pd.read_excel(uploaded_file)
        st.write("Formato original do arquivo de funcionários:")
        st.write(df.head(2))
        
        # Limpar nomes das colunas
        df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
        
        # Mostrar informações detalhadas sobre o salário
        if 'Salário Mês Atual' in df.columns:
            st.write("### Informações sobre a coluna de salário")
            st.write(f"Tipo da coluna: {df['Salário Mês Atual'].dtype}")
            st.write("Amostra de valores de salário:")
            amostra = df['Salário Mês Atual'].head(5).tolist()
            for i, valor in enumerate(amostra):
                st.write(f"Valor {i+1}: {valor} (tipo: {type(valor)})")
                
            # Converter valores de salário
            df['Salário Mês Atual'] = df['Salário Mês Atual'].apply(converter_valor_monetario)
            st.write("Amostra após conversão:", df['Salário Mês Atual'].head())
            
        # Converter colunas de data
        if 'Data Término Contrato' in df.columns:
            df['Data Término Contrato'] = df['Data Término Contrato'].apply(converter_data_br_para_datetime)
            
        if 'Data Admissão' in df.columns:
            st.write("Amostra de valores da coluna 'Data Admissão' antes da conversão:", df['Data Admissão'].head())
            df['Data Admissão'] = df['Data Admissão'].apply(converter_data_br_para_datetime)
            st.write("Amostra após conversão:", df['Data Admissão'].head())
            
        # Verificar valores únicos de horas mensais
        if 'Qtd Horas Mensais' in df.columns:
            st.write("Valores únicos de Qtd Horas Mensais:", df['Qtd Horas Mensais'].unique())
            # Converter para float
            df['Qtd Horas Mensais'] = pd.to_numeric(df['Qtd Horas Mensais'], errors='coerce').fillna(0)
            
        st.success("Arquivo de funcionários carregado com sucesso")
        return df
    except Exception as e:
        st.error(f"Erro ao carregar arquivo de funcionários: {e}")
        return pd.DataFrame()

def carregar_arquivo_afastamentos(uploaded_file):
    if uploaded_file is None:
        return pd.DataFrame()
    try:
        st.info(f"Carregando arquivo de afastamentos: {uploaded_file.name}")
        df = pd.read_excel(uploaded_file)
        st.write("Formato original do arquivo de afastamentos:")
        st.write(df.head(2))
        
        # Limpar nomes das colunas
        df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
        
        st.success("Arquivo de afastamentos carregado com sucesso")
        return df
    except Exception as e:
        st.error(f"Erro ao carregar arquivo de afastamentos: {e}")
        return pd.DataFrame()

def consolidar_dados_funcionario(df_combinado):
    if df_combinado.empty:
        return df_combinado
        
    st.write("### Consolidando dados de funcionários")
    st.write(f"Número de registros antes da consolidação: {len(df_combinado)}")
    st.write("Colunas disponíveis para consolidação:", df_combinado.columns.tolist())
    
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
    st.write(f"Número de registros após consolidação: {len(df_consolidado)}")
    
    # Mostrar amostra do resultado consolidado
    st.write("Amostra do resultado consolidado:")
    st.write(df_consolidado.head(3))
    
    return df_consolidado

def processar_dados(df_ausencias, df_funcionarios, df_afastamentos, data_limite_admissao=None):
    st.write("### Iniciando processamento de dados")
    
    if df_ausencias.empty or df_funcionarios.empty:
        st.warning("Um ou mais arquivos não puderam ser carregados corretamente.")
        return pd.DataFrame()
        
    st.write(f"Registros de ausências: {len(df_ausencias)}")
    st.write(f"Registros de funcionários: {len(df_funcionarios)}")
    st.write(f"Registros de afastamentos: {len(df_afastamentos) if not df_afastamentos.empty else 0}")
    
    # Converter data limite
    if data_limite_admissao:
        data_limite = converter_data_br_para_datetime(data_limite_admissao)
        st.write(f"Data limite de admissão: {data_limite}")
    else:
        data_limite = None
        
    # Filtrar por data de admissão
    if data_limite is not None:
        contagem_antes = len(df_funcionarios)
        df_funcionarios = df_funcionarios[df_funcionarios['Data Admissão'] <= data_limite]
        st.write(f"Funcionários filtrados por data de admissão: {contagem_antes} -> {len(df_funcionarios)}")
    
    # Verificar e preparar matrículas
    if 'Matricula' in df_ausencias.columns and 'Matricula' in df_funcionarios.columns:
        st.write("Convertendo matrículas para string...")
        
        # Converter matrículas para string
        df_ausencias['Matricula'] = df_ausencias['Matricula'].astype(str)
        df_funcionarios['Matricula'] = df_funcionarios['Matricula'].astype(str)
        
        # Mostrar valores únicos de exemplo
        st.write("Exemplo de matrículas em df_ausencias:", df_ausencias['Matricula'].head(3).tolist())
        st.write("Exemplo de matrículas em df_funcionarios:", df_funcionarios['Matricula'].head(3).tolist())
        
        # Mesclar dados
        st.write("Mesclando dados de funcionários e ausências...")
        df_combinado = pd.merge(
            df_funcionarios,
            df_ausencias,
            on='Matricula',
            how='left'
        )
        
        # Remover colunas duplicadas após o merge
        colunas_duplicadas = df_combinado.columns[df_combinado.columns.duplicated()]
        if not colunas_duplicadas.empty:
            st.write(f"Colunas duplicadas encontradas: {colunas_duplicadas.tolist()}")
            df_combinado = df_combinado.loc[:, ~df_combinado.columns.duplicated()]
        
        st.write(f"Registros após mesclar funcionários e ausências: {len(df_combinado)}")
        
        # Mesclar com afastamentos, se disponível
        if not df_afastamentos.empty:
            if 'Matricula' in df_afastamentos.columns:
                st.write("Mesclando dados de afastamentos...")
                df_afastamentos['Matricula'] = df_afastamentos['Matricula'].astype(str)
                
                df_combinado = pd.merge(
                    df_combinado,
                    df_afastamentos,
                    on='Matricula',
                    how='left',
                    suffixes=('', '_afastamento')
                )
                
                # Remover novamente colunas duplicadas, se houver
                colunas_duplicadas = df_combinado.columns[df_combinado.columns.duplicated()]
                if not colunas_duplicadas.empty:
                    st.write(f"Colunas duplicadas após mesclar afastamentos: {colunas_duplicadas.tolist()}")
                    df_combinado = df_combinado.loc[:, ~df_combinado.columns.duplicated()]
                    
                st.write(f"Registros após mesclar afastamentos: {len(df_combinado)}")
        
        # Consolidar dados
        df_consolidado = consolidar_dados_funcionario(df_combinado)
        
        # Aplicar regras de pagamento
        st.write("Aplicando regras de pagamento...")
        df_final = aplicar_regras_pagamento(df_consolidado)
        
        return df_final
    else:
        st.warning("Aviso: Colunas de Matricula não encontradas em um ou ambos os arquivos.")
        if 'Matricula' not in df_ausencias.columns:
            st.write("Coluna 'Matricula' não encontrada em df_ausencias. Colunas disponíveis:", df_ausencias.columns.tolist())
        if 'Matricula' not in df_funcionarios.columns:
            st.write("Coluna 'Matricula' não encontrada em df_funcionarios. Colunas disponíveis:", df_funcionarios.columns.tolist())
        return pd.DataFrame()

def aplicar_regras_pagamento(df):
    if df.empty:
        return df
        
    st.write("### Aplicando regras de pagamento")
    
    # Adicionar colunas de resultado
    df['Valor a Pagar'] = 0.0
    df['Status'] = ''
    df['Cor'] = ''
    df['Observacoes'] = ''
    
    # Listas de funções e afastamentos
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
    
    # Mostrar estatísticas antes de aplicar regras
    if 'Cargo' in df.columns:
        st.write("Distribuição de cargos:", df['Cargo'].value_counts())
    if 'Salário Mês Atual' in df.columns:
        st.write("Estatísticas de salário:", df['Salário Mês Atual'].describe())
    if 'Qtd Horas Mensais' in df.columns:
        st.write("Distribuição de horas mensais:", df['Qtd Horas Mensais'].value_counts())
    
    # Aplicar regras a cada funcionário
    registros_processados = 0
    registros_com_direito = 0
    registros_sem_direito = 0
    registros_aguardando = 0
    
    for idx, row in df.iterrows():
        # Verificar cargo
        cargo = str(row.get('Cargo', '')).strip()
        if cargo in funcoes_sem_direito:
            df.at[idx, 'Valor a Pagar'] = 0.0
            df.at[idx, 'Status'] = 'Não tem direito'
            df.at[idx, 'Cor'] = 'vermelho'
            df.at[idx, 'Observacoes'] = f'Cargo sem direito: {cargo}'
            registros_sem_direito += 1
            continue
        
        # Verificar salário
        salario = 0
        if 'Salário Mês Atual' in df.columns:
            salario = float(row.get('Salário Mês Atual', 0))
            
        if salario >= 2542.86:
            df.at[idx, 'Valor a Pagar'] = 0.0
            df.at[idx, 'Status'] = 'Não tem direito'
            df.at[idx, 'Cor'] = 'vermelho'
            df.at[idx, 'Observacoes'] = f'Salário acima do limite: R$ {salario:.2f}'
            registros_sem_direito += 1
            continue
        
        # Verificar faltas
        if 'Falta' in df.columns and not pd.isna(row.get('Falta')):
            df.at[idx, 'Valor a Pagar'] = 0.0
            df.at[idx, 'Status'] = 'Não tem direito'
            df.at[idx, 'Cor'] = 'vermelho'
            df.at[idx, 'Observacoes'] = 'Tem falta'
            registros_sem_direito += 1
            continue
        
        # Verificar afastamentos
        if 'Afastamentos' in df.columns and not pd.isna(row.get('Afastamentos')):
            tipo_afastamento = str(row.get('Afastamentos')).strip()
            
            if tipo_afastamento in afastamentos_sem_direito:
                df.at[idx, 'Valor a Pagar'] = 0.0
                df.at[idx, 'Status'] = 'Não tem direito'
                df.at[idx, 'Cor'] = 'vermelho'
                df.at[idx, 'Observacoes'] = f'Afastamento sem direito: {tipo_afastamento}'
                registros_sem_direito += 1
                continue
            elif tipo_afastamento in afastamentos_aguardar_decisao:
                df.at[idx, 'Valor a Pagar'] = 0.0
                df.at[idx, 'Status'] = 'Aguardando decisão'
                df.at[idx, 'Cor'] = 'azul'
                df.at[idx, 'Observacoes'] = f'Afastamento para avaliar: {tipo_afastamento}'
                registros_aguardando += 1
                continue
            elif tipo_afastamento in afastamentos_com_direito:
                pass  # Continua o processamento
            else:
                df.at[idx, 'Valor a Pagar'] = 0.0
                df.at[idx, 'Status'] = 'Aguardando decisão'
                df.at[idx, 'Cor'] = 'azul'
                df.at[idx, 'Observacoes'] = f'Afastamento não classificado: {tipo_afastamento}'
                registros_aguardando += 1
                continue
        
        # Verificar ausências
        if (('Ausência Integral' in df.columns and not pd.isna(row.get('Ausência Integral'))) or 
            ('Ausência Parcial' in df.columns and not pd.isna(row.get('Ausência Parcial')))):
            df.at[idx, 'Valor a Pagar'] = 0.0
            df.at[idx, 'Status'] = 'Aguardando decisão'
            df.at[idx, 'Cor'] = 'azul'
            df.at[idx, 'Observacoes'] = 'Tem ausência - Necessita avaliação'
            registros_aguardando += 1
            continue
        
        # Verificar horas mensais
        horas = 0
        if 'Qtd Horas Mensais' in df.columns:
            horas = float(row.get('Qtd Horas Mensais', 0))
        
        if horas == 220:
            df.at[idx, 'Valor a Pagar'] = 300.00
            df.at[idx, 'Status'] = 'Tem direito'
            df.at[idx, 'Cor'] = 'verde'
            df.at[idx, 'Observacoes'] = 'Sem ausências - 220 horas'
            registros_com_direito += 1
        elif horas in [110, 120]:
            df.at[idx, 'Valor a Pagar'] = 150.00
            df.at[idx, 'Status'] = 'Tem direito'
            df.at[idx, 'Cor'] = 'verde'
            df.at[idx, 'Observacoes'] = f'Sem ausências - {int(horas)} horas'
            registros_com_direito += 1
        else:
            df.at[idx, 'Valor a Pagar'] = 0.00
            df.at[idx, 'Status'] = 'Aguardando decisão'
            df.at[idx, 'Cor'] = 'azul'
            df.at[idx, 'Observacoes'] = f'Sem ausências - verificar horas: {horas}'
            registros_aguardando += 1
        
        registros_processados += 1
    
    # Resumo da aplicação de regras
    st.write(f"Total de registros processados: {registros_processados}")
    st.write(f"Registros com direito: {registros_com_direito}")
    st.write(f"Registros sem direito: {registros_sem_direito}")
    st.write(f"Registros aguardando decisão: {registros_aguardando}")
    
    # Renomear colunas para padrão final
    df = df.rename(columns={
        'Matricula': 'Matricula',
        'Nome Funcionário': 'Nome',
        'Nome Local': 'Local',
        'Valor a Pagar': 'Valor_Premio'
    })
    
    return df

def exportar_novo_excel(df):
    st.write("### Gerando arquivo Excel para exportação")
    
    output = BytesIO()
    
    # Determinar colunas a exportar
    colunas_adicionais = []
    for coluna in ['Cargo', 'Salário Mês Atual', 'Qtd Horas Mensais', 'Falta', 'Afastamentos', 'Ausência Integral', 'Ausência Parcial']:
        if coluna in df.columns:
            colunas_adicionais.append(coluna)
    
    colunas_principais = ['Matricula', 'Nome', 'Local', 'Valor_Premio', 'Observacoes']
    colunas_exportar = colunas_adicionais + colunas_principais
    
    # Filtrar dataframes por status
    df_tem_direito = df[df['Status'] == 'Tem direito'][colunas_exportar]
    df_nao_tem_direito = df[df['Status'] == 'Não tem direito'][colunas_exportar]
    df_aguardando = df[df['Status'] == 'Aguardando decisão'][colunas_exportar]
    
    st.write(f"Registros com direito: {len(df_tem_direito)}")
    st.write(f"Registros sem direito: {len(df_nao_tem_direito)}")
    st.write(f"Registros aguardando decisão: {len(df_aguardando)}")
    
    # Criar arquivo Excel
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Aba de funcionários com direito
        df_tem_direito.to_excel(writer, index=False, sheet_name='Com Direito')
        workbook = writer.book
        worksheet = writer.sheets['Com Direito']
        format_green = workbook.add_format({'bg_color': '#CCFFCC'})
        
        for i in range(len(df_tem_direito)):
            for j in range(len(colunas_exportar)):
                valor = df_tem_direito.iloc[i][colunas_exportar[j]]
                # Converter para string, mas tratar NaN
                if pd.isna(valor):
                    valor_str = ""
                else:
                    valor_str = str(valor)
                worksheet.write(i+1, j, valor_str, format_green)
        
        # Aba de funcionários sem direito
        df_nao_tem_direito.to_excel(writer, index=False, sheet_name='Sem Direito')
        worksheet = writer.sheets['Sem Direito']
        format_red = workbook.add_format({'bg_color': '#FFCCCC'})
        
        for i in range(len(df_nao_tem_direito)):
            for j in range(len(colunas_exportar)):
                valor = df_nao_tem_direito.iloc[i][colunas_exportar[j]]
                if pd.isna(valor):
                    valor_str = ""
                else:
                    valor_str = str(valor)
                worksheet.write(i+1, j, valor_str, format_red)
        
        # Aba de funcionários aguardando decisão
        df_aguardando.to_excel(writer, index=False, sheet_name='Aguardando Decisão')
        worksheet = writer.sheets['Aguardando Decisão']
        format_blue = workbook.add_format({'bg_color': '#CCCCFF'})
        
        for i in range(len(df_aguardando)):
            for j in range(len(colunas_exportar)):
                valor = df_aguardando.iloc[i][colunas_exportar[j]]
                if pd.isna(valor):
                    valor_str = ""
                else:
                    valor_str = str(valor)
                worksheet.write(i+1, j, valor_str, format_blue)
        
        # Aba de resumo
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

# Sidebar para upload de arquivos
st.sidebar.header("Upload de Arquivos")
arquivo_ausencias = st.sidebar.file_uploader("Arquivo de Ausências", type=["xlsx", "xls"])
arquivo_funcionarios = st.sidebar.file_uploader("Arquivo de Funcionários", type=["xlsx", "xls"])
arquivo_afastamentos = st.sidebar.file_uploader("Arquivo de Afastamentos (opcional)", type=["xlsx", "xls"])

# Configuração de data limite
st.sidebar.header("Data Limite de Admissão")
data_limite = st.sidebar.date_input(
    "Considerar apenas funcionários admitidos até:",
    value=datetime(2025, 3, 1),
    format="DD/MM/YYYY"
)

# Abas de interface
tab1, tab2, tab3 = st.tabs(["Processamento Inicial", "Edição e Exportação", "Diagnóstico"])

with tab1:
    processar = st.button("Processar Dados")
    if processar:
        if arquivo_ausencias is not None and arquivo_funcionarios is not None:
            with st.spinner("Processando arquivos..."):
                # Corrigido: Usar o arquivo_funcionarios correto em vez de arquivo_afastamentos
                df_ausencias = carregar_arquivo_ausencias(arquivo_ausencias)
                df_funcionarios = carregar_arquivo_funcionarios(arquivo_funcionarios)
                df_afastamentos = carregar_arquivo_afastamentos(arquivo_afastamentos)
                
                data_limite_str = data_limite.strftime("%d/%m/%Y")
                resultado = processar_dados(df_ausencias, df_funcionarios, df_afastamentos, data_limite_admissao=data_limite_str)
                
                if not resultado.empty:
                    st.success(f"Processamento concluído com sucesso. {len(resultado)} registros encontrados.")
                    
                    # Converter todas as colunas para string para exibição
                    df_display = resultado.copy()
                    for col in df_display.columns:
                        df_display[col] = df_display[col].astype(str)
                    
                    # Exibir informações básicas
                    st.write("### Primeiras linhas do resultado:")
                    st.write(df_display.head())
                    
                    # Mostrar estatísticas
                    total_a_pagar = resultado['Valor_Premio'].sum()
                    contagem_por_status = resultado['Status'].value_counts()
                    
                    st.subheader("Resumo")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Total a Pagar", f"R$ {total_a_pagar:.2f}")
                    with col2:
                        st.write("Contagem por Status:")
                        st.write(contagem_por_status)
                    
                    # Salvar o resultado na session_state para uso na aba de edição
                    st.session_state.resultado_processado = resultado
                    
                    # Botão para exportar
                    st.subheader("Exportar Resultados")
                    if st.button("Gerar Arquivo Excel"):
                        excel_data = exportar_novo_excel(resultado)
                        st.download_button(
                            label="Baixar Arquivo Excel",
                            data=excel_data,
                            file_name=f"resultado_processamento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    st.info("Vá para a aba 'Edição e Exportação' para ajustar os valores manualmente se necessário.")
                else:
                    st.warning("Nenhum resultado encontrado após o processamento.")
        else:
            st.warning("Por favor, faça o upload dos arquivos de ausências e funcionários.")

with tab2:
    st.header("Edição e Exportação dos Resultados")
    
    if 'resultado_processado' in st.session_state:
        resultado = st.session_state.resultado_processado
        
        # Opções de filtro
        st.subheader("Filtrar Resultados")
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.multiselect(
                "Filtrar por Status:",
                options=resultado['Status'].unique(),
                default=resultado['Status'].unique()
            )
        with col2:
            if 'Local' in resultado.columns:
                local_filter = st.multiselect(
                    "Filtrar por Local:",
                    options=resultado['Local'].unique(),
                    default=[]
                )
            
        # Aplicar filtros
        df_filtered = resultado.copy()
        if status_filter:
            df_filtered = df_filtered[df_filtered['Status'].isin(status_filter)]
        if 'Local' in resultado.columns and local_filter:
            df_filtered = df_filtered[df_filtered['Local'].isin(local_filter)]
        
        # Mostrar contagem de registros filtrados
        st.write(f"Exibindo {len(df_filtered)} de {len(resultado)} registros")
        
        # Exibir dataframe com opção de edição
        st.write("### Dados para Edição")
        st.write("Você pode editar os dados diretamente na tabela abaixo:")
        
        # Criar cópia para exibição
        df_display = df_filtered.copy()
        
        # Formatar colunas monetárias
        if 'Valor_Premio' in df_display.columns:
            df_display['Valor_Premio'] = df_display['Valor_Premio'].apply(lambda x: f'R$ {float(x):.2f}')
        if 'Salário Mês Atual' in df_display.columns:
            df_display['Salário Mês Atual'] = df_display['Salário Mês Atual'].apply(lambda x: f'R$ {float(x):.2f}')
        
        # Exibir dataframe
        st.dataframe(df_display)
        
        # Botão para exportar
        st.subheader("Exportar Resultados")
        if st.button("Gerar Arquivo Excel Final"):
            excel_data = exportar_novo_excel(resultado)
            st.download_button(
                label="Baixar Arquivo Excel",
                data=excel_data,
                file_name=f"resultado_processamento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("Por favor, primeiro processe os dados na aba 'Processamento Inicial'.")

with tab3:
    st.header("Diagnóstico e Depuração")
    
    st.write("""
    Esta aba permite diagnosticar problemas nos arquivos de entrada.
    Carregue os arquivos na barra lateral e clique no botão abaixo para analisar.
    """)
    
    if st.button("Analisar Arquivos"):
        if arquivo_ausencias is not None:
            st.subheader("Diagnóstico do Arquivo de Ausências")
            try:
                df = pd.read_excel(arquivo_ausencias)
                st.write(f"Número de linhas: {len(df)}")
                st.write(f"Número de colunas: {len(df.columns)}")
                st.write("Colunas encontradas:", df.columns.tolist())
                
                # Verificar se tem a coluna Matricula
                if 'Matricula' in df.columns:
                    st.write(f"Coluna 'Matricula' encontrada com {df['Matricula'].nunique()} valores únicos")
                    st.write("Primeiros valores de Matricula:", df['Matricula'].head().tolist())
                else:
                    st.error("Coluna 'Matricula' não encontrada no arquivo de ausências!")
                    colunas_similares = [col for col in df.columns if 'matric' in col.lower()]
                    if colunas_similares:
                        st.write("Colunas similares encontradas:", colunas_similares)
                
                # Amostra de dados
                st.write("Amostra de dados do arquivo de ausências:")
                st.dataframe(df.head(3))
                
                # Verificar tipos de dados
                st.write("Tipos de dados:")
                st.write(df.dtypes)
                
            except Exception as e:
                st.error(f"Erro ao analisar arquivo de ausências: {e}")
        
        if arquivo_funcionarios is not None:
            st.subheader("Diagnóstico do Arquivo de Funcionários")
            try:
                df = pd.read_excel(arquivo_funcionarios)
                st.write(f"Número de linhas: {len(df)}")
                st.write(f"Número de colunas: {len(df.columns)}")
                st.write("Colunas encontradas:", df.columns.tolist())
                
                # Verificar se tem a coluna Matricula
                if 'Matricula' in df.columns:
                    st.write(f"Coluna 'Matricula' encontrada com {df['Matricula'].nunique()} valores únicos")
                    st.write("Primeiros valores de Matricula:", df['Matricula'].head().tolist())
                else:
                    st.error("Coluna 'Matricula' não encontrada no arquivo de funcionários!")
                    colunas_similares = [col for col in df.columns if 'matric' in col.lower()]
                    if colunas_similares:
                        st.write("Colunas similares encontradas:", colunas_similares)
                
                # Verificar colunas críticas
                colunas_criticas = ['Salário Mês Atual', 'Qtd Horas Mensais', 'Data Admissão', 'Cargo']
                for coluna in colunas_criticas:
                    if coluna in df.columns:
                        st.write(f"Coluna '{coluna}' encontrada")
                        if coluna == 'Salário Mês Atual':
                            st.write("Tipo da coluna Salário:", df[coluna].dtype)
                            st.write("Amostra de valores:", df[coluna].head().tolist())
                        elif coluna == 'Qtd Horas Mensais':
                            st.write("Valores únicos de horas:", df[coluna].unique())
                    else:
                        st.warning(f"Coluna crítica '{coluna}' não encontrada!")
                
                # Amostra de dados
                st.write("Amostra de dados do arquivo de funcionários:")
                st.dataframe(df.head(3))
                
                # Verificar tipos de dados
                st.write("Tipos de dados:")
                st.write(df.dtypes)
                
            except Exception as e:
                st.error(f"Erro ao analisar arquivo de funcionários: {e}")
        
        if arquivo_afastamentos is not None:
            st.subheader("Diagnóstico do Arquivo de Afastamentos")
            try:
                df = pd.read_excel(arquivo_afastamentos)
                st.write(f"Número de linhas: {len(df)}")
                st.write(f"Número de colunas: {len(df.columns)}")
                st.write("Colunas encontradas:", df.columns.tolist())
                
                # Verificar se tem a coluna Matricula
                if 'Matricula' in df.columns:
                    st.write(f"Coluna 'Matricula' encontrada com {df['Matricula'].nunique()} valores únicos")
                else:
                    st.error("Coluna 'Matricula' não encontrada no arquivo de afastamentos!")
                
                # Verificar coluna de afastamentos
                if 'Afastamentos' in df.columns:
                    st.write("Tipos de afastamentos encontrados:", df['Afastamentos'].unique())
                
                # Amostra de dados
                st.write("Amostra de dados do arquivo de afastamentos:")
                st.dataframe(df.head(3))
                
            except Exception as e:
                st.error(f"Erro ao analisar arquivo de afastamentos: {e}")
        
        if arquivo_ausencias is None and arquivo_funcionarios is None and arquivo_afastamentos is None:
            st.warning("Nenhum arquivo foi carregado para análise.")

# Informações de ajuda na barra lateral
st.sidebar.markdown("---")
st.sidebar.info("""
**Como usar o aplicativo:**
1. Faça o upload dos arquivos necessários.
2. Defina a data limite de admissão.
3. Na aba 'Processamento Inicial', clique em 'Processar Dados'.
4. Revise os resultados e vá para a aba 'Edição e Exportação' para ajustar valores se necessário.
5. O relatório final terá abas para:
   - Funcionários com Direito (verde)
   - Funcionários sem Direito (vermelho)
   - Funcionários Aguardando Decisão (azul)
   - Resumo do processamento

**Solução de Problemas:**
- Se encontrar erros, use a aba 'Diagnóstico' para analisar seus arquivos.
- Verifique se as colunas essenciais estão presentes e com o formato correto.
- Para problemas de formato de data, certifique-se que estão no padrão DD/MM/AAAA.
""")
