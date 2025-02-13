import streamlit as st
import pandas as pd

def editar_valores_status(df):
    # Função para editar valores e status das pessoas que têm direito ao prêmio
    # Retorna o DataFrame atualizado
    st.subheader("Editar Valores e Status")
    
    # Filtro principal por status
    status_filter = st.selectbox("Filtrar por Status", options=["Todos", "Tem direito", "Não tem direito", "Aguardando decisão"])
    if status_filter != "Todos":
        df = df[df['Status'].str.contains(status_filter)]
    
    # Filtros de pesquisa
    matricula_filter = st.text_input("Filtrar por Matrícula")
    if matricula_filter:
        df = df[df['Matricula'].astype(str).str.contains(matricula_filter)]
    
    nome_filter = st.text_input("Filtrar por Nome")
    if nome_filter:
        df = df[df['Nome'].str.contains(nome_filter, case=False)]
    
    # Ordenação
    ordenar_por = st.selectbox("Ordenar por", options=["Nome", "Matricula"])
    df = df.sort_values(by=ordenar_por)
    
    # Mostrar métricas
    st.metric("Total de Funcionários no Filtro Atual", len(df))
    st.metric("Total de Funcionários com Direito no Filtro Atual", len(df[df['Status'].str.contains("Tem direito")]))
    st.metric("Valor Total dos Prêmios no Filtro Atual", f"R$ {df['Valor_Premio'].sum():,.2f}")
    
    if not df.empty:
        # Configurar colunas editáveis
        for index, row in df.iterrows():
            st.write(f"Funcionário: {row['Nome']}")
            status_options = ["Tem direito", "Não tem direito", "Aguardando decisão"]
            current_status = next((opt for opt in status_options if opt in row['Status']), "Tem direito")
            df.at[index, 'Status'] = st.selectbox(
                "Status",
                options=status_options,
                index=status_options.index(current_status),
                key=f"status_{index}"
            )
            df.at[index, 'Valor_Premio'] = st.number_input(
                "Valor do Prêmio",
                min_value=0.0,
                value=row['Valor_Premio'],
                step=50.0,
                key=f"valor_premio_{index}"
            )
            df.at[index, 'Observações'] = st.text_input(
                "Observações",
                value=row.get('Observações', ''),
                key=f"observacoes_{index}"
            )
            st.write("---")
    
    return df

def exportar_novo_excel(df, caminho_arquivo):
    # Filtrar apenas as pessoas que têm direito ao prêmio
    df_direito = df[df['Status'].str.contains('Tem direito')]
    
    # Selecionar as colunas desejadas
    df_exportar = df_direito[['CPF', 'SomaDeVALOR', 'Nome', 'CNPJ']]
    
    # Exportar para um novo arquivo Excel
    df_exportar.to_excel(caminho_arquivo, index=False)
