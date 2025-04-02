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
        # Se for uma string, tenta converter
        if isinstance(data_str, str):
            # Tenta converter do formato brasileiro DD/MM/YYYY
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
        # Lê o arquivo Excel
        df = pd.read_excel(uploaded_file)
        
        # Converte as colunas de data
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
        # Lê o arquivo Excel
        df = pd.read_excel(uploaded_file)
        
        # Exibe informações de debug sobre a coluna de salário
        if 'Salário Mês Atual' in df.columns:
            # Mostra informações sobre os tipos e valores da coluna de salário
            st.write("### Informações sobre a coluna de salário")
            st.write(f"Tipo da coluna: {df['Salário Mês Atual'].dtype}")
            
            # Amostra de valores
            st.write("Amostra de valores de salário:")
            amostra = df['Salário Mês Atual'].head(5).tolist()
            for i, valor in enumerate(amostra):
                st.write(f"Valor {i+1}: {valor} (tipo: {type(valor)})")
        
        # Converte as colunas de data
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
        # Lê o arquivo Excel
        df = pd.read_excel(uploaded_file)
        return df
    except Exception as e:
        st.error(f"Erro ao carregar arquivo de afastamentos: {e}")
        return pd.DataFrame()

def consolidar_dados_funcionario(df_combinado):
    """
    Consolida os dados para que cada funcionário apareça apenas uma vez no relatório.
    """
    if df_combinado.empty:
        return df_combinado
    
    # Verifica colunas necessárias
    colunas_funcionario = ['Matrícula', 'Nome Funcionário', 'Cargo', 'Qtd Horas Mensais', 
                          'Salário Mês Atual', 'Data Admissão', 'Nome Local', 'Tipo Contrato']
    
    # Lista para guardar os dados consolidados
    dados_consolidados = []
    
    # Agrupa por matrícula para processar cada funcionário individualmente
    for matricula, grupo in df_combinado.groupby('Matrícula'):
        # Pega os dados básicos do funcionário (mesmos para todas as ocorrências)
        funcionario = grupo.iloc[0].copy()
        
        # Verifica se tem faltas
        tem_falta = False
        if 'Falta' in df_combinado.columns:
            tem_falta = grupo['Falta'].notna().any()
        
        # Verifica se tem afastamentos
        tem_afastamento = False
        if 'Afastamentos' in df_combinado.columns:
            tem_afastamento = grupo['Afastamentos'].notna().any()
        
        # Verifica se tem ausências
        tem_ausencia = False
        if 'Ausência Integral' in df_combinado.columns and 'Ausência Parcial' in df_combinado.columns:
            tem_ausencia = grupo['Ausência Integral'].notna().any() or grupo['Ausência Parcial'].notna().any()
        
        # Cria uma nova linha para o funcionário
        nova_linha = pd.Series({
            'Tem Falta': tem_falta,
            'Tem Afastamento': tem_afastamento,
            'Tem Ausência': tem_ausencia
        })
        
        # Adiciona os dados do funcionário
        for coluna in colunas_funcionario:
            if coluna in funcionario.index:
                nova_linha[coluna] = funcionario[coluna]
        
        dados_consolidados.append(nova_linha)
    
    # Cria o DataFrame consolidado
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
        df_funcionarios = df_funcionarios[
            df_funcionarios['Data Admissão'] <= data_limite
        ]
    
    # Mescla os dados de ausências com os dados de funcionários
    # Usando a "Matricula" como chave de junção
    if 'Matricula' in df_ausencias.columns and 'Matrícula' in df_funcionarios.columns:
        # Converte a coluna Matrícula para o mesmo tipo em ambos os DataFrames
        df_ausencias['Matricula'] = df_ausencias['Matricula'].astype(str)
        df_funcionarios['Matrícula'] = df_funcionarios['Matrícula'].astype(str)
        
        # Mescla os DataFrames
        df_combinado = pd.merge(
            df_funcionarios,
            df_ausencias,
            left_on='Matrícula',
            right_on='Matricula',
            how='left'
        )
        
        # Se tiver dados de afastamentos, inclui também
        if not df_afastamentos.empty:
            if 'Matricula' in df_afastamentos.columns:
                df_afastamentos['Matricula'] = df_afastamentos['Matricula'].astype(str)
                df_combinado = pd.merge(
                    df_combinado,
                    df_afastamentos,
                    left_on='Matrícula',
                    right_on='Matricula',
                    how='left',
                    suffixes=('', '_afastamento')
                )
        
        # Consolida os dados para que cada funcionário apareça apenas uma vez
        df_consolidado = consolidar_dados_funcionario(df_combinado)
        
        # Aplicar as regras de cálculo
        df_final = aplicar_regras_pagamento(df_consolidado)
        
        return df_final
    else:
        st.warning("Aviso: Colunas de matrícula não encontradas em um ou ambos os arquivos.")
        return pd.DataFrame()

def aplicar_regras_pagamento(df):
    """
    Aplica as regras de cálculo de pagamento conforme os critérios especificados.
    
    Regras:
    - Apenas funcionários com salário abaixo de R$ 2.542,86 têm direito a receber
    - Se tem falta -> 0,00 (vermelho)
    - Se tem afastamento -> 0,00 (vermelho)
    - Se tem ausência -> Avaliar (azul)
    - Se horas = 220 -> R$ 300,00 (verde)
    - Se horas = 110 ou 120 -> R$ 150,00 (verde)
    - Se linha em branco -> Paga conforme horas
    """
    # Adiciona coluna de resultado
    df['Valor a Pagar'] = 0.0
    df['Status'] = ''
    df['Cor'] = ''
    df['Observacoes'] = ''  # Adiciona coluna de observações
    
    # Cria uma tabela de depuração para exibir valores processados de salário
    debug_data = []
    
    # Processa cada linha
    for idx, row in df.iterrows():
        # Verifica salário - FUNÇÃO MELHORADA
        salario = 0
        salario_original = row.get('Salário Mês Atual', None)
        
        try:
            # Tenta converter direto se for um tipo numérico
            if isinstance(salario_original, (int, float)):
                salario = float(salario_original)
            # Caso seja string, limpa e converte
            elif isinstance(salario_original, str):
                salario_str = salario_original
                # Remove R$ e espaços se existirem
                salario_limpo = salario_str.replace('R$', '').replace(' ', '')
                # Se usa vírgula como separador decimal, converte para ponto
                if ',' in salario_limpo and '.' not in salario_limpo:
                    salario_limpo = salario_limpo.replace(',', '.')
                # Converte para float
                salario = float(salario_limpo)
            
            # Debug information
            debug_item = {
                'Matricula': row.get('Matrícula', '') or row.get('Matricula', ''),
                'Nome': row.get('Nome Funcionário', '') or row.get('Nome', ''),
                'Salario_Original': salario_original,
                'Salario_Tipo': type(salario_original).__name__,
                'Salario_Convertido': salario
            }
            debug_data.append(debug_item)
            
        except (ValueError, TypeError) as e:
            # Se ocorrer erro na conversão, registra para debug
            debug_item = {
                'Matricula': row.get('Matrícula', '') or row.get('Matricula', ''),
                'Nome': row.get('Nome Funcionário', '') or row.get('Nome', ''),
                'Salario_Original': salario_original,
                'Salario_Tipo': type(salario_original).__name__,
                'Erro': str(e)
            }
            debug_data.append(debug_item)
            salario = 0
        
        # Verifica horas
        horas = 0
        if 'Qtd Horas Mensais' in df.columns:
            try:
                horas_valor = row['Qtd Horas Mensais']
                if isinstance(horas_valor, (int, float)):
                    horas = float(horas_valor)
                elif isinstance(horas_valor, str) and horas_valor.strip():
                    horas = float(horas_valor.strip())
            except (ValueError, TypeError):
                horas = 0
        
        # Verifica ocorrências
        tem_falta = row.get('Tem Falta', False)
        tem_afastamento = row.get('Tem Afastamento', False)
        tem_ausencia = row.get('Tem Ausência', False)
        
        # Aplicar regras na ordem correta e definir status conforme utils.py
        if salario >= 2542.86:
            df.at[idx, 'Valor a Pagar'] = 0.00
            df.at[idx, 'Status'] = 'Não tem direito'
            df.at[idx, 'Cor'] = 'vermelho'
            df.at[idx, 'Observacoes'] = f'Salário acima do limite: R$ {salario:.2f}'
        elif tem_falta:
            df.at[idx, 'Valor a Pagar'] = 0.00
            df.at[idx, 'Status'] = 'Não tem direito'
            df.at[idx, 'Cor'] = 'vermelho'
            df.at[idx, 'Observacoes'] = 'Tem falta'
        elif tem_afastamento:
            df.at[idx, 'Valor a Pagar'] = 0.00
            df.at[idx, 'Status'] = 'Não tem direito'
            df.at[idx, 'Cor'] = 'vermelho'
            df.at[idx, 'Observacoes'] = 'Tem afastamento'
        elif tem_ausencia:
            df.at[idx, 'Valor a Pagar'] = 0.00
            df.at[idx, 'Status'] = 'Aguardando decisão'
            df.at[idx, 'Cor'] = 'azul'
            df.at[idx, 'Observacoes'] = 'Tem ausência - Necessita avaliação'
        else:
            # Paga conforme as horas
            if horas == 220:
                df.at[idx, 'Valor a Pagar'] = 300.00
                df.at[idx, 'Status'] = 'Tem direito'
                df.at[idx, 'Cor'] = 'verde'
                df.at[idx, 'Observacoes'] = '220 horas'
            elif horas == 110 or horas == 120:
                df.at[idx, 'Valor a Pagar'] = 150.00
                df.at[idx, 'Status'] = 'Tem direito'
                df.at[idx, 'Cor'] = 'verde'
                df.at[idx, 'Observacoes'] = f'{int(horas)} horas'
            else:
                df.at[idx, 'Valor a Pagar'] = 0.00
                df.at[idx, 'Status'] = 'Aguardando decisão'
                df.at[idx, 'Cor'] = ''
                df.at[idx, 'Observacoes'] = f'Verificar horas: {horas}'
    
    # Exibe a tabela de depuração de salários
    st.write("### Debug: Conversão de Salários")
    df_debug = pd.DataFrame(debug_data)
    if not df_debug.empty:
        st.dataframe(df_debug.head(10))  # Mostra apenas as primeiras 10 linhas para não sobrecarregar a UI
        
        # Mostra estatísticas de salários
        salarios_acima_limite = sum(1 for item in debug_data if item.get('Salario_Convertido', 0) >= 2542.86)
        salarios_abaixo_limite = sum(1 for item in debug_data if item.get('Salario_Convertido', 0) < 2542.86)
        
        st.write(f"Total de registros: {len(debug_data)}")
        st.write(f"Salários acima do limite (≥ R$2.542,86): {salarios_acima_limite}")
        st.write(f"Salários abaixo do limite (< R$2.542,86): {salarios_abaixo_limite}")
    
    # Renomeia colunas para compatibilidade com utils.py
    df = df.rename(columns={
        'Matrícula': 'Matricula',
        'Nome Funcionário': 'Nome',
        'Nome Local': 'Local',
        'Valor a Pagar': 'Valor_Premio'
    })
    
    return df

# Função para converter ou limpar valores para exportação
def preparar_para_excel(df):
    """
    Prepara o DataFrame para exportação para Excel, 
    convertendo datas para strings e limpando valores problemáticos.
    """
    df_limpo = df.copy()
    
    # Converte datas para strings
    for coluna in df_limpo.columns:
        # Identifica colunas com datas
        if df_limpo[coluna].dtype == 'datetime64[ns]' or 'datetime' in str(df_limpo[coluna].dtype):
            # Converte para string no formato DD/MM/YYYY
            df_limpo[coluna] = df_limpo[coluna].apply(
                lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) and not pd.isna(x) else ''
            )
        
        # Trata valores NaT, NaN e None
        df_limpo[coluna] = df_limpo[coluna].apply(
            lambda x: '' if pd.isna(x) or x is None or (isinstance(x, float) and np.isnan(x)) else x
        )
    
    return df_limpo

def exportar_novo_excel(df):
    """
    Exporta o DataFrame para Excel com formatação adequada.
    Função baseada no utils.py.
    """
    output = BytesIO()
    
    df_direito = df[df['Status'].str.contains('Tem direito')].copy()
    df_exportar = df_direito[['Matricula', 'Nome', 'Local', 'Valor_Premio', 'Observacoes']]
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_exportar.to_excel(writer, index=False, sheet_name='Funcionários com Direito')
        
        resumo_data = [
            ['RESUMO DO PROCESSAMENTO'],
            [f'Data de Geração: {pd.Timestamp.now().strftime("%d/%m/%Y %H:%M:%S")}'],
            [''],
            ['Métricas Gerais'],
            [f'Total de Funcionários Processados: {len(df)}'],
            [f'Total de Funcionários com Direito: {len(df_direito)}'],
            [f'Valor Total dos Prêmios: R$ {df_direito["Valor_Premio"].sum():,.2f}'],
        ]
        
        pd.DataFrame(resumo_data).to_excel(
            writer, 
            index=False, 
            header=False, 
            sheet_name='Resumo'
        )
    
    output.seek(0)
    return output.getvalue()

def salvar_alteracoes(idx, novo_status, novo_valor, nova_obs, nome):
    """Função auxiliar para salvar alterações (baseada no utils.py)"""
    st.session_state.modified_df.at[idx, 'Status'] = novo_status
    st.session_state.modified_df.at[idx, 'Valor_Premio'] = novo_valor
    st.session_state.modified_df.at[idx, 'Observacoes'] = nova_obs
    st.session_state.expanded_item = idx
    st.session_state.last_saved = nome
    st.session_state.show_success = True

def editar_valores_status(df):
    """Função para editar os valores de pagamento e status (baseada no utils.py)"""
    if 'modified_df' not in st.session_state:
        st.session_state.modified_df = df.copy()
    
    if 'expanded_item' not in st.session_state:
        st.session_state.expanded_item = None
        
    if 'show_success' not in st.session_state:
        st.session_state.show_success = False
        
    if 'last_saved' not in st.session_state:
        st.session_state.last_saved = None
    
    st.subheader("Filtro Principal")
    
    status_options = ["Todos", "Tem direito", "Não tem direito", "Aguardando decisão"]
    
    status_principal = st.selectbox(
        "Selecione o status para visualizar:",
        options=status_options,
        index=0,
        key="status_principal_filter_unique"
    )
    
    df_filtrado = st.session_state.modified_df.copy()
    if status_principal != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Status'] == status_principal]
    
    st.subheader("Buscar Funcionários")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        matricula_busca = st.text_input("Buscar por Matrícula", key="matricula_search_unique")
    with col2:
        nome_busca = st.text_input("Buscar por Nome", key="nome_search_unique")
    with col3:
        ordem = st.selectbox(
            "Ordenar por:",
            options=["Nome (A-Z)", "Nome (Z-A)", "Matrícula (Crescente)", "Matrícula (Decrescente)"],
            key="ordem_select_unique"
        )
    
    if matricula_busca:
        df_filtrado = df_filtrado[df_filtrado['Matricula'].astype(str).str.contains(matricula_busca)]
    if nome_busca:
        df_filtrado = df_filtrado[df_filtrado['Nome'].str.contains(nome_busca, case=False)]
    
    # Ordenação
    if ordem == "Nome (A-Z)":
        df_filtrado = df_filtrado.sort_values('Nome')
    elif ordem == "Nome (Z-A)":
        df_filtrado = df_filtrado.sort_values('Nome', ascending=False)
    elif ordem == "Matrícula (Crescente)":
        df_filtrado = df_filtrado.sort_values('Matricula')
    elif ordem == "Matrícula (Decrescente)":
        df_filtrado = df_filtrado.sort_values('Matricula', ascending=False)
    
    #
