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
        # Se for uma string, tenta converter do formato brasileiro DD/MM/YYYY
        if isinstance(data_str, str):
            return datetime.strptime(data_str, '%d/%m/%Y')
        else:
            return data_str  # Se já for outro tipo (como timestamp), retorna como está
    except (ValueError, TypeError):
        try:
            # Caso falhe, tenta outros formatos comuns
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
    Carrega o arquivo de ausências e converte colunas de data corretamente.
    """
    try:
        df = pd.read_excel(uploaded_file)
        # Normaliza os nomes das colunas
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
    Carrega o arquivo de funcionários e converte colunas de data corretamente.
    """
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
    """
    Carrega o arquivo de afastamentos, se existir.
    """
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
    """
    Consolida os dados para que cada funcionário apareça apenas uma vez,
    mantendo todas as informações da tabela de ausências.
    """
    if df_combinado.empty:
        return df_combinado

    dados_consolidados = []
    
    # Agrupa por Matricula para que cada funcionário apareça uma vez
    for matricula, grupo in df_combinado.groupby('Matricula'):
        funcionario = grupo.iloc[0].copy()
        
        # Verifica se há ausência, falta ou afastamento em qualquer registro do grupo
        tem_falta = False
        if 'Falta' in df_combinado.columns:
            tem_falta = grupo['Falta'].notna().any()
        
        tem_afastamento = False
        if 'Afastamentos' in df_combinado.columns:
            tem_afastamento = grupo['Afastamentos'].notna().any()
        
        tem_ausencia = False
        if ('Ausência Integral' in df_combinado.columns) and ('Ausência Parcial' in df_combinado.columns):
            tem_ausencia = grupo['Ausência Integral'].notna().any() or grupo['Ausência Parcial'].notna().any()
        
        # Acrescenta as colunas auxiliares
        funcionario['Tem Falta'] = tem_falta
        funcionario['Tem Afastamento'] = tem_afastamento
        funcionario['Tem Ausência'] = tem_ausencia
        
        dados_consolidados.append(funcionario)
    
    df_consolidado = pd.DataFrame(dados_consolidados)
    return df_consolidado

def processar_dados(df_ausencias, df_funcionarios, df_afastamentos, data_limite_admissao=None):
    """
    Processa os dados sem filtro de mês.
    """
    if df_ausencias.empty or df_funcionarios.empty:
        st.warning("Um ou mais arquivos não puderam ser carregados corretamente.")
        return pd.DataFrame()
    
    # Converte a data limite de admissão
    if data_limite_admissao:
        data_limite = converter_data_br_para_datetime(data_limite_admissao)
    else:
        data_limite = None
    
    # Filtra funcionários pela data limite de admissão, se especificado
    if data_limite is not None:
        df_funcionarios = df_funcionarios[df_funcionarios['Data Admissão'] <= data_limite]
    
    # Mescla os dados de ausências com os dados de funcionários usando "Matricula" como chave
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
        
        # Inclui os dados de afastamentos, se existirem
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
        
        df_consolidado = consolidar_dados_funcionario(df_combinado)
        df_final = aplicar_regras_pagamento(df_consolidado)
        return df_final
    else:
        st.warning("Aviso: Colunas de Matricula não encontradas em um ou ambos os arquivos.")
        return pd.DataFrame()

def aplicar_regras_pagamento(df):
    """
    Aplica as regras de pagamento conforme os critérios especificados:
    
    - Se o cargo for um dos que não têm direito, marca imediatamente como "Não tem direito".
    - Se o salário for igual ou superior a R$ 2.542,86, marca como "Não tem direito".
    - Se houver falta, marca como "Não tem direito".
    - Se houver registro de afastamento, verifica:
         * Se o afastamento for dos que NÃO dão direito, marca como "Não tem direito".
         * Se o afastamento for de "Atraso", marca como "Aguardando decisão".
         * Se o afastamento for dos que dão direito (Abonado Gerencia Loja ou Abono Administrativo),
           segue para o cálculo baseado nas horas.
         * Se o tipo de afastamento não for identificado, marca como "Aguardando decisão".
    - Se houver ausência (Ausência Integral ou Parcial), marca como "Aguardando decisão".
    - Caso não haja nenhum impeditivo, calcula o valor a pagar com base na quantidade de horas:
         * 220 horas = R$ 300,00
         * 110 ou 120 horas = R$ 150,00
         * Caso contrário, marca como "Aguardando decisão" para verificação de horas.
    """
    # Inicializa as colunas necessárias
    df['Valor a Pagar'] = 0.0
    df['Status'] = ''
    df['Cor'] = ''
    df['Observacoes'] = ''

    # Lista de cargos que não têm direito
    funcoes_sem_direito = [
        'AUX DE SERV GERAIS (INT)', 
        'AUX DE LIMPEZA (INT)',
        'LIMPADOR DE VIDROS INT', 
        'RECEPCIONISTA INTERMITENTE', 
        'PORTEIRO INTERMITENTE'
    ]
    
    # Listas de afastamentos
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
    
    # Processa cada funcionário
    for idx, row in df.iterrows():
        # 1. Verifica o cargo
        cargo = str(row.get('Cargo', '')).strip()
        if cargo in funcoes_sem_direito:
            df.at[idx, 'Valor a Pagar'] = 0.0
            df.at[idx, 'Status'] = 'Não tem direito'
            df.at[idx, 'Cor'] = 'vermelho'
            df.at[idx, 'Observacoes'] = f'Cargo sem direito: {cargo}'
            continue
        
        # 2. Verifica o salário
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

        # 3. Verifica faltas
        if 'Falta' in df.columns and not pd.isna(row.get('Falta')):
            df.at[idx, 'Valor a Pagar'] = 0.0
            df.at[idx, 'Status'] = 'Não tem direito'
            df.at[idx, 'Cor'] = 'vermelho'
            df.at[idx, 'Observacoes'] = 'Tem falta'
            continue

        # 4. Verifica afastamentos
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
                # Se o afastamento dá direito, continua para a verificação das horas
                pass
            else:
                df.at[idx, 'Valor a Pagar'] = 0.0
                df.at[idx, 'Status'] = 'Aguardando decisão'
                df.at[idx, 'Cor'] = 'azul'
                df.at[idx, 'Observacoes'] = f'Afastamento não classificado: {tipo_afastamento}'
                continue

        # 5. Verifica ausências (se houver)
        if (('Ausência Integral' in df.columns and not pd.isna(row.get('Ausência Integral'))) or 
            ('Ausência Parcial' in df.columns and not pd.isna(row.get('Ausência Parcial')))):
            df.at[idx, 'Valor a Pagar'] = 0.0
            df.at[idx, 'Status'] = 'Aguardando decisão'
            df.at[idx, 'Cor'] = 'azul'
            df.at[idx, 'Observacoes'] = 'Tem ausência - Necessita avaliação'
            continue

        # 6. Se não houve nenhum impeditivo, calcula com base na quantidade de horas
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

    # Renomeia colunas para compatibilidade na exportação
    df = df.rename(columns={
        'Matricula': 'Matricula',
        'Nome Funcionário': 'Nome',
        'Nome Local': 'Local',
        'Valor a Pagar': 'Valor_Premio'
    })
    
    return df

def exportar_novo_excel(df):
    output = BytesIO()
    
    # Define as colunas adicionais: além das que já temos, inclui as da ausências
    colunas_adicionais = []
    for coluna in ['Cargo', 'Salário Mês Atual', 'Qtd Horas Mensais', 'Falta', 'Afastamentos', 'Ausência Integral', 'Ausência Parcial']:
        if coluna in df.columns:
            colunas_adicionais.append(coluna)
    
    # Colunas principais que queremos sempre incluir
    colunas_principais = ['Matricula', 'Nome', 'Local', 'Valor_Premio', 'Observacoes']
    colunas_exportar = colunas_adicionais + colunas_principais
    
    # Filtra os DataFrames para cada status
    df_tem_direito = df[df['Status'] == 'Tem direito'][colunas_exportar]
    df_nao_tem_direito = df[df['Status'] == 'Não tem direito'][colunas_exportar]
    df_aguardando = df[df['Status'] == 'Aguardando decisão'][colunas_exportar]
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Aba: Funcionários com direito
        df_tem_direito.to_excel(writer, index=False, sheet_name='Com Direito')
        workbook = writer.book
        worksheet = writer.sheets['Com Direito']
        format_green = workbook.add_format({'bg_color': '#CCFFCC'})
        for i in range(len(df_tem_direito)):
            for j in range(len(colunas_exportar)):
                worksheet.write(i+1, j, str(df_tem_direito.iloc[i][colunas_exportar[j]]), format_green)
        
        # Aba: Funcionários sem direito
        df_nao_tem_direito.to_excel(writer, index=False, sheet_name='Sem Direito')
        worksheet = writer.sheets['Sem Direito']
        format_red = workbook.add_format({'bg_color': '#FFCCCC'})
        for i in range(len(df_nao_tem_direito)):
            for j in range(len(colunas_exportar)):
                worksheet.write(i+1, j, str(df_nao_tem_direito.iloc[i][colunas_exportar[j]]), format_red)
        
        # Aba: Funcionários aguardando decisão
        df_aguardando.to_excel(writer, index=False, sheet_name='Aguardando Decisão')
        worksheet = writer.sheets['Aguardando Decisão']
        format_blue = workbook.add_format({'bg_color': '#CCCCFF'})
        for i in range(len(df_aguardando)):
            for j in range(len(colunas_exportar)):
                worksheet.write(i+1, j, str(df_aguardando.iloc[i][colunas_exportar[j]]), format_blue)
        
        # Aba adicional: Resumo
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

# Interface principal do Streamlit
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
                    
                    # Imprime as primeiras linhas do DataFrame
                    st.write("Primeiras linhas do DataFrame:")
                    st.write(resultado.head())
                    
                    # Verifica os tipos de dados de cada coluna
                    st.write("Tipos de dados das colunas:")
                    st.write(resultado.dtypes)
                    
                    # Exibe a distribuição dos status
                    st.write("Distribuição dos status:", resultado['Status'].value_counts())
                    
                    st.subheader("Resultados Preliminares")
                    st.dataframe(resultado)
                    
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
