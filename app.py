import streamlit as st
import pandas as pd
import pdfkit
from datetime import datetime
import os
import logging

# Configuração do logging
logging.basicConfig(
    filename='sistema_premios.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def carregar_tipos_afastamento():
    if os.path.exists("tipos_afastamento.pkl"):
        return pd.read_pickle("tipos_afastamento.pkl")
    return pd.DataFrame({"tipo": [], "categoria": []})

def salvar_tipos_afastamento(df):
    df.to_pickle("tipos_afastamento.pkl")

def processar_ausencias(df):
    df['Matricula'] = pd.to_numeric(df['Matricula'], errors='coerce')
    df = df.dropna(subset=['Matricula'])
    df['Matricula'] = df['Matricula'].astype(int)
    
    def converter_para_horas(tempo):
        if pd.isna(tempo) or tempo == '' or tempo == '00:00':
            return 0
        try:
            if ':' in str(tempo):
                horas, minutos = map(int, str(tempo).split(':'))
                return horas + minutos / 60
            return 0
        except:
            return 0
    
    df['Horas_Atraso'] = df['Ausencia_Parcial'].apply(converter_para_horas)
    df['Afastamentos'] = df['Afastamentos'].fillna('').astype(str)
    
    resultado = df.groupby('Matricula').agg({
        'Horas_Atraso': 'sum',
        'Afastamentos': lambda x: '; '.join(sorted(set(filter(None, x))))
    }).reset_index()
    
    resultado['Atrasos'] = resultado['Horas_Atraso'].apply(
        lambda x: f"{int(x)}:{int((x % 1) * 60):02d}" if x > 0 else ""
    )
    resultado = resultado.drop('Horas_Atraso', axis=1)
    
    return resultado

def calcular_premio(df_funcionarios, df_ausencias, data_limite_admissao):
    df_funcionarios = df_funcionarios[
        pd.to_datetime(df_funcionarios['Data_Admissao']) <= pd.to_datetime(data_limite_admissao)
    ]
    
    resultados = []
    for _, func in df_funcionarios.iterrows():
        ausencias = df_ausencias[df_ausencias['Matricula'] == func['Matricula']]
        tem_afastamento_impeditivo = not ausencias.empty
        
        valor_premio = 300.00 if func['Qtd_Horas_Mensais'] == 220 else 150.00
        status = "Tem direito" if not tem_afastamento_impeditivo else "Não tem direito"
        
        resultados.append({
            'Matricula': func['Matricula'],
            'Nome': func['Nome_Funcionario'],
            'Cargo': func['Cargo'],
            'Local': func['Nome_Local'],
            'Horas_Mensais': func['Qtd_Horas_Mensais'],
            'Data_Admissao': func['Data_Admissao'],
            'Valor_Premio': valor_premio if status == "Tem direito" else 0,
            'Status': status,
            'Detalhes_Afastamentos': ausencias['Afastamentos'].iloc[0] if not ausencias.empty else ''
        })
    
    return pd.DataFrame(resultados)

def main():
    st.set_page_config(page_title="Sistema de Verificação de Prêmios", page_icon="🏆", layout="wide")
    st.title("Sistema de Verificação de Prêmios")

    with st.sidebar:
        st.header("Configurações")
        
        data_limite = st.date_input("Data Limite de Admissão")
        uploaded_func = st.file_uploader("Carregar base de funcionários", type=['xlsx'])
        uploaded_ausencias = st.file_uploader("Carregar base de ausências", type=['xlsx'])

    if uploaded_func and uploaded_ausencias and data_limite:
        try:
            df_funcionarios = pd.read_excel(uploaded_func)
            df_funcionarios['Data_Admissao'] = pd.to_datetime(df_funcionarios['Data_Admissao'], format='%d/%m/%Y')

            df_ausencias = pd.read_excel(uploaded_ausencias)
            df_ausencias = processar_ausencias(df_ausencias)

            df_resultado = calcular_premio(df_funcionarios, df_ausencias, data_limite)

            st.subheader("Resultado do Cálculo de Prêmios")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Funcionários", len(df_resultado))
            with col2:
                st.metric("Funcionários com Direito", len(df_resultado[df_resultado['Status'] == 'Tem direito']))
            with col3:
                st.metric("Valor Total Pago", f"R$ {df_resultado['Valor_Premio'].sum():,.2f}")

            # Geração de HTML
            html_content = f"""
            <html>
                <head>
                    <style>
                        body {{ font-family: Arial, sans-serif; margin: 20px; }}
                        h1 {{ color: #1f77b4; text-align: center; }}
                        table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                        th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
                        th {{ background-color: #1f77b4; color: white; }}
                        tr:nth-child(even) {{ background-color: #f8f9fa; }}
                    </style>
                </head>
                <body>
                    <h1>RELATÓRIO DE PRÊMIOS</h1>
                    <p>Data do relatório: {datetime.now().strftime('%d/%m/%Y')}</p>
                    <table>
                        <tr>
                            <th>Matrícula</th>
                            <th>Nome</th>
                            <th>Cargo</th>
                            <th>Local</th>
                            <th>Valor Prêmio</th>
                        </tr>
            """
            for _, row in df_resultado.iterrows():
                html_content += f"""
                    <tr>
                        <td>{row['Matricula']}</td>
                        <td>{row['Nome']}</td>
                        <td>{row['Cargo']}</td>
                        <td>{row['Local']}</td>
                        <td style="text-align: right;">R$ {row['Valor_Premio']:,.2f}</td>
                    </tr>
                """
            html_content += "</table></body></html>"

            # Botão para exportação PDF
            if st.button("📑 Exportar Relatório como PDF"):
                try:
                    options = {
                        'page-size': 'A4',
                        'margin-top': '20mm',
                        'margin-right': '20mm',
                        'margin-bottom': '20mm',
                        'margin-left': '20mm',
                        'encoding': 'UTF-8',
                        'orientation': 'Landscape'
                    }
                    pdf = pdfkit.from_string(html_content, False, options=options)

                    st.download_button(
                        label="⬇️ Baixar PDF",
                        data=pdf,
                        file_name="relatorio_premios.pdf",
                        mime="application/pdf"
                    )
                except Exception as e:
                    st.error(f"Erro ao gerar PDF: {str(e)}")

        except Exception as e:
            st.error(f"Erro ao processar os dados: {str(e)}")

if __name__ == "__main__":
    main()

