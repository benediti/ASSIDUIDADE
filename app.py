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

def filtrar_ausencias_por_periodo(df_ausencias, data_inicio, data_fim):
    """
    Filtra ausências para um determinado período.
    """
    # Converte data_inicio e data_fim para datetime se forem strings
    if isinstance(data_inicio, str):
        data_inicio = converter_data_br_para_datetime(data_inicio)
    if isinstance(data_fim, str):
        data_fim = converter_data_br_para_datetime(data_fim)
    
    # Verifica se a conversão ocorreu com sucesso
    if data_inicio is None or data_fim is None:
        st.error("Erro: formato de data inválido.")
        return pd.DataFrame()
    
    # Filtra as ausências dentro do período
    try:
        # Primeiro verifica se a coluna Dia existe e tem valores válidos
        if 'Dia' not in df_ausencias.columns or df_ausencias['Dia'].isna().all():
            st.error("Erro ao filtrar ausências por período: 'Dia'")
            return df_ausencias
        
        # Filtra apenas registros com data válida
        df_com_data = df_ausencias.dropna(subset=['Dia'])
        
        ausencias_periodo = df_com_data[
            (df_com_data['Dia'] >= data_inicio) & 
            (df_com_data['Dia'] <= data_fim)
        ]
        return ausencias_periodo
    except Exception as e:
        st.error(f"Erro ao filtrar ausências por período: {e}")
        return df_ausencias

def processar_dados(df_ausencias, df_funcionarios, df_afastamentos, mes=None, ano=None, data_limite_admissao=None):
    """
    Processa os dados de ambos os arquivos com tratamento correto de datas.
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
    
    # Se mês e ano forem fornecidos, filtra para esse período
    if mes is not None and ano is not None:
        # Determine o primeiro e último dia do mês
        primeiro_dia = datetime(ano, mes, 1)
        if mes == 12:
            ultimo_dia = datetime(ano + 1, 1, 1) - timedelta(days=1)
        else:
            ultimo_dia = datetime(ano, mes + 1, 1) - timedelta(days=1)
        
        # Filtra as ausências para o mês/ano especificado
        df_ausencias = filtrar_ausencias_por_periodo(df_ausencias, primeiro_dia, ultimo_dia)
    
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
        
        # Aplicar as regras de cálculo
        df_combinado = aplicar_regras_pagamento(df_combinado)
        
        return df_combinado
    else:
        st.warning("Aviso: Colunas de matrícula não encontradas em um ou ambos os arquivos.")
        return df_ausencias  # Retorna apenas as ausências se não for possível combinar

def aplicar_regras_pagamento(df):
    """
    Aplica as regras de cálculo de pagamento conforme os critérios especificados.
    
    Regras:
    - Se salário > R$ 2.542,86 -> Não paga (vermelho)
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
    
    # Processa cada linha
    for idx, row in df.iterrows():
        # Verifica salário
        salario = 0
        if 'Salário Mês Atual' in df.columns:
            try:
                # Converte para float, tratando possíveis strings
                salario_str = str(row['Salário Mês Atual']).replace('R$', '').replace('.', '').replace(',', '.')
                salario = float(salario_str)
            except (ValueError, TypeError):
                salario = 0
        
        # Verifica horas
        horas = 0
        if 'Qtd Horas Mensais' in df.columns:
            try:
                horas = float(row['Qtd Horas Mensais'])
            except (ValueError, TypeError):
                horas = 0
        
        # Verifica ocorrências
        tem_falta = False
        if 'Falta' in df.columns and not pd.isna(row['Falta']):
            tem_falta = True
        
        tem_afastamento = False
        if 'Afastamentos' in df.columns and not pd.isna(row['Afastamentos']):
            tem_afastamento = True
        
        tem_ausencia = False
        if ('Ausência Integral' in df.columns and not pd.isna(row['Ausência Integral'])) or \
           ('Ausência Parcial' in df.columns and not pd.isna(row['Ausência Parcial'])):
            tem_ausencia = True
        
        # Aplicar regras na ordem correta
        if salario > 2542.86:
            df.at[idx, 'Valor a Pagar'] = 0.00
            df.at[idx, 'Status'] = 'Não paga - Salário acima do limite'
            df.at[idx, 'Cor'] = 'vermelho'
        elif tem_falta:
            df.at[idx, 'Valor a Pagar'] = 0.00
            df.at[idx, 'Status'] = 'Não paga - Tem falta'
            df.at[idx, 'Cor'] = 'vermelho'
        elif tem_afastamento:
            df.at[idx, 'Valor a Pagar'] = 0.00
            df.at[idx, 'Status'] = 'Não paga - Tem afastamento'
            df.at[idx, 'Cor'] = 'vermelho'
        elif tem_ausencia:
            df.at[idx, 'Valor a Pagar'] = 0.00
            df.at[idx, 'Status'] = 'Avaliar - Tem ausência'
            df.at[idx, 'Cor'] = 'azul'
        else:
            # Paga conforme as horas
            if horas == 220:
                df.at[idx, 'Valor a Pagar'] = 300.00
                df.at[idx, 'Status'] = 'Paga - 220 horas'
                df.at[idx, 'Cor'] = 'verde'
            elif horas == 110 or horas == 120:
                df.at[idx, 'Valor a Pagar'] = 150.00
                df.at[idx, 'Status'] = f'Paga - {horas} horas'
                df.at[idx, 'Cor'] = 'verde'
            else:
                df.at[idx, 'Valor a Pagar'] = 0.00
                df.at[idx, 'Status'] = 'Verificar horas trabalhadas'
                df.at[idx, 'Cor'] = ''
    
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

# Interface Streamlit
st.sidebar.header("Upload de Arquivos")

# Upload dos arquivos Excel
arquivo_ausencias = st.sidebar.file_uploader("Arquivo de Ausências", type=["xlsx", "xls"])
arquivo_funcionarios = st.sidebar.file_uploader("Arquivo de Funcionários", type=["xlsx", "xls"])
arquivo_afastamentos = st.sidebar.file_uploader("Arquivo de Afastamentos (opcional)", type=["xlsx", "xls"])

# Seleção de período
st.sidebar.header("Filtrar por Período")
meses = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 
         6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 
         10: "Outubro", 11: "Novembro", 12: "Dezembro"}

mes_selecionado = st.sidebar.selectbox("Mês", options=list(meses.keys()), format_func=lambda x: meses[x])
ano_selecionado = st.sidebar.number_input("Ano", min_value=2020, max_value=2030, value=2025)

# Data limite de admissão
st.sidebar.header("Data Limite de Admissão")
data_limite = st.sidebar.date_input(
    "Considerar apenas funcionários admitidos até:",
    value=datetime(2025, 3, 1),
    format="DD/MM/YYYY"
)

# Botão para processar
processar = st.sidebar.button("Processar Dados")

# Processamento
if processar:
    if arquivo_ausencias is not None and arquivo_funcionarios is not None:
        with st.spinner("Processando arquivos..."):
            # Carrega os dados
            df_ausencias = carregar_arquivo_ausencias(arquivo_ausencias)
            df_funcionarios = carregar_arquivo_funcionarios(arquivo_funcionarios)
            df_afastamentos = carregar_arquivo_afastamentos(arquivo_afastamentos)
            
            # Converte data limite para string no formato brasileiro
            data_limite_str = data_limite.strftime("%d/%m/%Y")
            
            # Processa os dados
            resultado = processar_dados(
                df_ausencias, 
                df_funcionarios, 
                df_afastamentos,
                mes=mes_selecionado, 
                ano=ano_selecionado,
                data_limite_admissao=data_limite_str
            )
            
            if not resultado.empty:
                st.success(f"Processamento concluído com sucesso. {len(resultado)} registros encontrados.")
                
                # Estiliza o DataFrame para exibição
                def highlight_row(row):
                    if row['Cor'] == 'vermelho':
                        return ['background-color: #FFCCCC'] * len(row)
                    elif row['Cor'] == 'verde':
                        return ['background-color: #CCFFCC'] * len(row)
                    elif row['Cor'] == 'azul':
                        return ['background-color: #CCCCFF'] * len(row)
                    else:
                        return [''] * len(row)
                
                # Remove a coluna Cor antes de exibir
                df_exibir = resultado.copy()
                
                # Exibe os resultados
                st.subheader("Resultados")
                st.dataframe(df_exibir.style.apply(highlight_row, axis=1))
                
                # Resumo de valores
                total_a_pagar = df_exibir['Valor a Pagar'].sum()
                contagem_por_status = df_exibir['Status'].value_counts()
                
                st.subheader("Resumo")
                col1, col2 = st.columns(2)
                
                with col1:
                    st.metric("Total a Pagar", f"R$ {total_a_pagar:.2f}")
                
                with col2:
                    st.write("Contagem por Status:")
                    st.write(contagem_por_status)
                
                # Botão para download CSV
                csv = df_exibir.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download CSV",
                    data=csv,
                    file_name=f"resultado_ausencias_{meses[mes_selecionado]}_{ano_selecionado}.csv",
                    mime="text/csv"
                )
                
                # Preparar dados para Excel (limpar valores problemáticos)
                df_excel = preparar_para_excel(df_exibir)
                
                # Botão para download do Excel - VERSÃO CORRIGIDA
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_excel.to_excel(writer, index=False, sheet_name='Resultado')
                    
                    # Acessa o objeto workbook e worksheet
                    workbook = writer.book
                    worksheet = writer.sheets['Resultado']
                    
                    # Adiciona formatos condicionais
                    formato_vermelho = workbook.add_format({'bg_color': '#FFCCCC'})
                    formato_verde = workbook.add_format({'bg_color': '#CCFFCC'})
                    formato_azul = workbook.add_format({'bg_color': '#CCCCFF'})
                    
                    # Métodos alternativos para aplicar formatação condicional
                    for i, row in df_excel.iterrows():
                        cor = df_exibir.iloc[i]['Cor']
                        if cor == 'vermelho':
                            formato = formato_vermelho
                        elif cor == 'verde':
                            formato = formato_verde
                        elif cor == 'azul':
                            formato = formato_azul
                        else:
                            continue  # Pula linhas sem cor definida
                            
                        # Aplica o formato a toda a linha
                        for j in range(len(df_excel.columns)):
                            # Use formato básico para todos os valores
                            worksheet.write(i+1, j, str(df_excel.iloc[i, j]), formato)
                
                # Move o cursor para o início do buffer
                buffer.seek(0)
                
                # Botão de download para o Excel
                st.download_button(
                    label="Download Excel",
                    data=buffer,
                    file_name=f"resultado_ausencias_{meses[mes_selecionado]}_{ano_selecionado}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Nenhum resultado encontrado para o período selecionado.")
    else:
        st.warning("Por favor, faça o upload dos arquivos necessários.")

# Exibe informações de uso
st.sidebar.markdown("---")
st.sidebar.info("""
**Como usar:**
1. Faça o upload do arquivo de ausências
2. Faça o upload do arquivo de funcionários
3. Faça o upload do arquivo de afastamentos (opcional)
4. Defina a data limite de admissão
5. Selecione o mês e ano desejados
6. Clique em "Processar Dados"
7. Veja os resultados e faça o download se necessário

**Regras de pagamento:**
- Vermelho: Não paga (0,00)
- Verde: Pagamento normal (R$ 300,00 ou R$ 150,00)
- Azul: Avaliar caso a caso
""")
