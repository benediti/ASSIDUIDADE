import streamlit as st
import pandas as pd

def editar_valores_status(df):
    if 'modified_df' not in st.session_state:
        st.session_state.modified_df = df.copy()
        
    if 'edited_rows' not in st.session_state:
        st.session_state.edited_rows = {}
    
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
        df_filtrado = df_filtrado[df_filtrado['Status'].str.contains(status_principal)]
    
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
        st.metric("Total com direito", len(df_filtrado[df_filtrado['Status'].str.contains('Tem direito')]))
    with col3:
        st.metric("Valor total dos prêmios", f"R$ {df_filtrado['Valor_Premio'].sum():,.2f}")
    
    # Editor de dados por linhas individuais
    st.subheader("Editor de Dados")
    
    # Adicionar chave para controlar qual expansor está aberto
    if 'expanded_item' not in st.session_state:
        st.session_state.expanded_item = None
    
    for idx, row in df_filtrado.iterrows():
        with st.expander(f"🧑‍💼 {row['Nome']} - Matrícula: {row['Matricula']}", 
                        expanded=st.session_state.expanded_item == row['Matricula']):
            col1, col2 = st.columns(2)
            
            with col1:
                novo_status = st.selectbox(
                    "Status",
                    options=status_options[1:],
                    index=status_options[1:].index(row['Status']) if row['Status'] in status_options[1:] else 0,
                    key=f"status_{row['Matricula']}"
                )
                
                novo_valor = st.number_input(
                    "Valor do Prêmio",
                    min_value=0.0,
                    max_value=1000.0,
                    value=float(row['Valor_Premio']),
                    step=50.0,
                    format="%.2f",
                    key=f"valor_{row['Matricula']}"
                )
            
            with col2:
                nova_obs = st.text_area(
                    "Observações",
                    value=row.get('Observacoes', ''),
                    key=f"obs_{row['Matricula']}"
                )
            
            if st.button("Salvar Alterações", key=f"save_{row['Matricula']}"):
                st.session_state.modified_df.at[idx, 'Status'] = novo_status
                st.session_state.modified_df.at[idx, 'Valor_Premio'] = novo_valor
                st.session_state.modified_df.at[idx, 'Observacoes'] = nova_obs
                st.session_state.expanded_item = row['Matricula']
                st.success(f"✅ Alterações salvas com sucesso para {row['Nome']}!")
                st.rerun()
    
    # Botões de ação geral
    st.subheader("Ações Gerais")
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("Reverter Todas as Alterações", key="revert_all_unique"):
            st.session_state.modified_df = df.copy()
            st.session_state.expanded_item = None
            st.warning("⚠️ Todas as alterações foram revertidas!")
            st.rerun()
    
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

def exportar_novo_excel(df):
    import io
    output = io.BytesIO()
    
    df_direito = df[df['Status'].str.contains('Tem direito')].copy()
    df_exportar = df_direito[['Matricula', 'Nome', 'Local', 'Valor_Premio', 'Observacoes']]
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
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
