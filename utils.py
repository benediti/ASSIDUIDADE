import streamlit as st

def editar_valores_status(df):
    # Função para editar valores e status das pessoas que têm direito ao prêmio
    # Retorna o DataFrame atualizado
    st.subheader("Editar Valores e Status")
    
    # Filtrar apenas as pessoas que têm direito ao prêmio
    df_direito = df[df['Status'] == 'Tem direito'].copy()
    
    # Exibir tabela editável
    edited_df = st.dataframe(df_direito)
    
    return edited_df

def exportar_novo_excel(df, caminho_arquivo):
    # Filtrar apenas as pessoas que têm direito ao prêmio
    df_direito = df[df['Status'] == 'Tem direito']
    
    # Selecionar as colunas desejadas
    df_exportar = df_direito[['CPF', 'SomaDeVALOR', 'Nome', 'CNPJ']]
    
    # Exportar para um novo arquivo Excel
    df_exportar.to_excel(caminho_arquivo, index=False)
