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
    - Caso não haja nenhum problema (ausência/falta/afastamento impeditivo), calcula o valor
      a pagar com base na quantidade de horas:
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
            continue  # Não processa mais este funcionário
        
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
                # Se o afastamento dá direito, prossegue para a verificação das horas
                pass
            else:
                # Se o tipo não for reconhecido, marca como "Aguardando decisão"
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
        'Matrícula': 'Matricula',
        'Nome Funcionário': 'Nome',
        'Nome Local': 'Local',
        'Valor a Pagar': 'Valor_Premio'
    })
    
    return df


def exportar_novo_excel(df):
    output = BytesIO()
    
    # Definir as colunas para exportação
    colunas_exportar = ['Matricula', 'Nome', 'Local', 'Valor_Premio', 'Observacoes']
    for coluna in ['Cargo', 'Salário Mês Atual', 'Qtd Horas Mensais']:
        if coluna in df.columns:
            colunas_exportar.insert(0, coluna)
    
    # Filtrar os DataFrames para cada status
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
    
    # Métricas
    st.subheader("Métricas do Filtro Atual")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Funcionários exibidos", len(df_filtrado))
    with col2:
        st.metric("Total com direito", len(df_filtrado[df_filtrado['Status'] == 'Tem direito']))
    with col3:
        st.metric("Valor total dos prêmios", f"R$ {df_filtrado['Valor_Premio'].sum():,.2f}")
    
    # Mostrar mensagem de sucesso se houver
    if st.session_state.show_success:
        st.success(f"✅ Alterações salvas com sucesso para {st.session_state.last_saved}!")
        st.session_state.show_success = False
    
    # Editor de dados por linhas individuais
    st.subheader("Editor de Dados")
    
    for idx, row in df_filtrado.iterrows():
        with st.expander(
            f"🧑‍💼 {row['Nome']} - Matrícula: {row['Matricula']}", 
            expanded=st.session_state.expanded_item == idx
        ):
            col1, col2 = st.columns(2)
            
            with col1:
                novo_status = st.selectbox(
                    "Status",
                    options=status_options[1:],
                    index=status_options[1:].index(row['Status']) if row['Status'] in status_options[1:] else 0,
                    key=f"status_{idx}_{row['Matricula']}"
                )
                
                novo_valor = st.number_input(
                    "Valor do Prêmio",
                    min_value=0.0,
                    max_value=1000.0,
                    value=float(row['Valor_Premio']),
                    step=50.0,
                    format="%.2f",
                    key=f"valor_{idx}_{row['Matricula']}"
                )
            
            with col2:
                nova_obs = st.text_area(
                    "Observações",
                    value=row.get('Observacoes', ''),
                    key=f"obs_{idx}_{row['Matricula']}"
                )
            
            if st.button("Salvar Alterações", key=f"save_{idx}_{row['Matricula']}"):
                salvar_alteracoes(idx, novo_status, novo_valor, nova_obs, row['Nome'])
    
    # Botões de ação geral
    st.subheader("Ações Gerais")
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("Reverter Todas as Alterações", key="revert_all_unique"):
            st.session_state.modified_df = df.copy()
            st.session_state.expanded_item = None
            st.session_state.show_success = False
            st.warning("⚠️ Todas as alterações foram revertidas!")
    
    with col2:
        if st.button("Exportar Arquivo Final", key="export_unique"):
            output = exportar_novo_excel(st.session_state.modified_df)
            st.download_button(
                label="📥 Baixar Arquivo Excel",
                data=output,
                file_name="funcionarios_premios.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_unique"
            )
    
    return st.session_state.modified_df

# Interface Streamlit
st.sidebar.header("Upload de Arquivos")

# Upload dos arquivos Excel
arquivo_ausencias = st.sidebar.file_uploader("Arquivo de Ausências", type=["xlsx", "xls"])
arquivo_funcionarios = st.sidebar.file_uploader("Arquivo de Funcionários", type=["xlsx", "xls"])
arquivo_afastamentos = st.sidebar.file_uploader("Arquivo de Afastamentos (opcional)", type=["xlsx", "xls"])

# Data limite de admissão
st.sidebar.header("Data Limite de Admissão")
data_limite = st.sidebar.date_input(
    "Considerar apenas funcionários admitidos até:",
    value=datetime(2025, 3, 1),
    format="DD/MM/YYYY"
)

# Tabs para diferenciar processamento e edição
tab1, tab2 = st.tabs(["Processamento Inicial", "Edição e Exportação"])

with tab1:
    # Botão para processar
    processar = st.button("Processar Dados")

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
                
                # Processa os dados sem filtro de mês
                resultado = processar_dados(
                    df_ausencias, 
                    df_funcionarios, 
                    df_afastamentos,
                    data_limite_admissao=data_limite_str
                )
                
                if not resultado.empty:
                    st.success(f"Processamento concluído com sucesso. {len(resultado)} registros encontrados.")
                    
                    # Guardamos o resultado no session_state para a tab de edição
                    st.session_state.resultado_processado = resultado
                    
                    # Colunas a serem removidas da exibição
                    colunas_esconder = ['Tem Falta', 'Tem Afastamento', 'Tem Ausência']

                    # Colunas visualizar a contagem dos status
                    st.success(f"Processamento concluído com sucesso. {len(resultado)} registros encontrados.")
                    st.write("Distribuição dos status:", st.session_state.resultado_processado['Status'].value_counts())

                    
                    # Criando uma cópia do DataFrame para exibição, mantendo a coluna 'Cor'
                    df_exibir = resultado.drop(columns=[col for col in colunas_esconder if col in resultado.columns])
                    
                    # Função de highlight baseada na coluna 'Cor' que ainda está presente
                    def highlight_row(row):
                        cor = row['Cor']
                        if cor == 'vermelho':
                            return ['background-color: #FFCCCC'] * len(row)
                        elif cor == 'verde':
                            return ['background-color: #CCFFCC'] * len(row)
                        elif cor == 'azul':
                            return ['background-color: #CCCCFF'] * len(row)
                        else:
                            return [''] * len(row)
                    
                    # Exibe os resultados com a coluna 'Cor' ainda presente para formatação
                    st.subheader("Resultados Preliminares")
                    st.dataframe(df_exibir.style.apply(highlight_row, axis=1))
                    
                    # Resumo de valores
                    total_a_pagar = df_exibir['Valor_Premio'].sum()
                    contagem_por_status = df_exibir['Status'].value_counts()
                    
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
        # Usa a função editar_valores_status do utils.py
        df_final = editar_valores_status(st.session_state.resultado_processado)
    else:
        st.info("Por favor, primeiro processe os dados na aba 'Processamento Inicial'.")

# Exibe informações de uso
st.sidebar.markdown("---")
st.sidebar.info("""
**Como usar:**
1. Faça o upload dos arquivos necessários
2. Defina a data limite de admissão
3. Na aba 'Processamento Inicial', clique em 'Processar Dados'
4. Na aba 'Edição e Exportação', revise e ajuste os valores individuais
5. Exporte o resultado final usando o botão 'Exportar Arquivo Final'

**Regras de pagamento:**
- Funcionários com salário acima de R$ 2.542,86 não têm direito
- Faltas e afastamentos bloqueiam o pagamento
- Ausências necessitam avaliação
- Pagamentos: 220h = R$300,00, 110/120h = R$150,00
""")
