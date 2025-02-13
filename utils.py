import streamlit as st
import pandas as pd

def editar_valores_status(df):
    # Função para editar valores e status das pessoas que têm direito ao prêmio
    # Retorna o DataFrame atualizado
    st.subheader("Editar Valores e Status")
    
    # Filtrar apenas as pessoas que têm direito ao prêmio
    df_direito = df.copy()
    
    if not df_direito.empty:
        # Configurar colunas editáveis
        for index, row in df_direito.iterrows():
            st.write(f"Funcionário: {row['Nome']}")
            df_direito.at[index, 'Status'] = st.selectbox(
                "Status",
                options=["Tem direito", "Não tem direito", "Aguardando decisão"],
                index=["Tem direito", "Não tem direito", "Aguardando decisão"].index(row['Status']),
                key=f"status_{index}"
            )
            df_direito.at[index, 'Valor_Premio'] = st.number_input(
                "Valor do Prêmio",
                min_value=0.0,
                value=row['Valor_Premio'],
                step=50.0,
                key=f"valor_premio_{index}"
            )
            df_direito.at[index, 'Observações'] = st.text_input(
                "Observações",
                value=row.get('Observações', ''),
                key=f"observacoes_{index}"
            )
            st.write("---")
        
        # Atualizar o DataFrame original com as edições
        df.update(df_direito)
    
    return df

def exportar_novo_excel(df, caminho_arquivo):
    # Filtrar apenas as pessoas que têm direito ao prêmio
    df_direito = df[df['Status'] == 'Tem direito']
    
    # Selecionar as colunas desejadas
    df_exportar = df_direito[['CPF', 'SomaDeVALOR', 'Nome', 'CNPJ']]
    
    # Exportar para um novo arquivo Excel
    df_exportar.to_excel(caminho_arquivo, index=False)
