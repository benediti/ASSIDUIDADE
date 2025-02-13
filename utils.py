import streamlit as st
import pandas as pd

def editar_valores_status(df):
    # Função para editar valores e status das pessoas que têm direito ao prêmio
    # Retorna o DataFrame atualizado
    st.subheader("Editar Valores e Status")
    
    # Inicializar estado
    if 'df_original' not in st.session_state:
        st.session_state.df_original = df.copy()
    if 'df_editado' not in st.session_state:
        st.session_state.df_editado = df.copy()
    
    # Filtro principal por status
    status_filter = st.selectbox("Filtrar por Status", options=["Todos", "Tem direito", "Não tem direito", "Aguardando decisão"], key="status_principal_selectbox")
    if status_filter != "Todos":
        df = st.session_state.df_editado[st.session_state.df_editado['Status'].str.contains(status_filter)]
    else:
        df = st.session_state.df_editado
    
    # Filtros de pesquisa
    matricula_filter = st.text_input("Filtrar por Matrícula", key="matricula_busca")
    if matricula_filter:
        df = df[df['Matricula'].astype(str).str.contains(matricula_filter)]
    
    nome_filter = st.text_input("Filtrar por Nome", key="nome_busca")
    if nome_filter:
        df = df[df['Nome'].str.contains(nome_filter, case=False)]
    
    # Ordenação
    ordenar_por = st.selectbox("Ordenar por", options=["Nome", "Matricula"], key="ordem_selectbox")
    df = df.sort_values(by=ordenar_por)
    
    # Mostrar métricas
    st.metric("Total de Funcionários no Filtro Atual", len(df), key="metric_total")
    st.metric("Total de Funcionários com Direito no Filtro Atual", len(df[df['Status'].str.contains("Tem direito")]), key="metric_direito")
    st.metric("Valor Total dos Prêmios no Filtro Atual", f"R$ {df['Valor_Premio'].sum():,.2f}", key="metric_valor")
    
    if not df.empty:
        # Configurar colunas editáveis
        for index, row in df.iterrows():
            st.write(f"Funcionário: {row['Nome']}")
            status_options = ["Tem direito", "Não tem direito", "Aguardando decisão"]
            current_status = next((opt for opt in status_options if opt in row['Status']), "Tem direito")
            st.session_state.df_editado.at[index, 'Status'] = st.selectbox(
                "Status",
                options=status_options,
                index=status_options.index(current_status),
                key=f"status_{index}"
            )
            st.session_state.df_editado.at[index, 'Valor_Premio'] = st.number_input(
                "Valor do Prêmio",
                min_value=0.0,
                value=row['Valor_Premio'],
                step=50.0,
                key=f"valor_premio_{index}"
            )
            st.session_state.df_editado.at[index, 'Observações'] = st.text_input(
                "Observações",
                value=row.get('Observações', ''),
                key=f"observacoes_{index}"
            )
            st.write("---")
    
    # Botões de ação
    if st.button("Salvar Alterações"):
        st.session_state.df_original = st.session_state.df_editado.copy()
        st.success("Alterações salvas com sucesso!")
    
    if st.button("Reverter Alterações"):
        st.session_state.df_editado = st.session_state.df_original.copy()
        st.warning("Alterações revertidas!")
    
    return st.session_state.df_editado

def exportar_novo_excel(df, caminho_arquivo):
    # Filtrar apenas as pessoas que têm direito ao prêmio
    df_direito = df[df['Status'].str.contains('Tem direito')]
    
    # Selecionar as colunas desejadas
    df_exportar = df_direito[['CPF', 'SomaDeVALOR', 'Nome', 'CNPJ']]
    
    # Exportar para um novo arquivo Excel
    with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
        df_exportar.to_excel(writer, index=False, sheet_name='Funcionarios com Direito')
        
        # Adicionar aba de resumo
        resumo = pd.DataFrame({
            "Métrica": ["Total de Funcionários", "Funcionários com Direito", "Valor Total dos Prêmios"],
            "Valor": [len(df), len(df_direito), f"R$ {df_direito['SomaDeVALOR'].sum():,.2f}"]
        })
        resumo.to_excel(writer, index=False, sheet_name='Resumo')
