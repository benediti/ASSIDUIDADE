import streamlit as st
import pandas as pd
import pdfkit
from datetime import datetime
import io
import os
import logging

# Configuração do logging
logging.basicConfig(
    filename='sistema_premios.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def carregar_tipos_afastamento():
    """Carrega tipos de afastamento do arquivo"""
    if os.path.exists("tipos_afastamento.pkl"):
        return pd.read_pickle("tipos_afastamento.pkl")
    return pd.DataFrame({"tipo": [], "categoria": []})

def salvar_tipos_afastamento(df):
    """Salva tipos de afastamento em arquivo"""
    df.to_pickle("tipos_afastamento.pkl")

def contar_faltas(valor):
    """Conta faltas: X = 1, vazio ou outros = 0"""
    if isinstance(valor, str) and valor.strip().upper() == 'X':
        return 1
    return 0

def processar_ausencias(df):
    """Processa e agrupa as ausências por matrícula"""
    # Converter matrícula para inteiro
    df['Matricula'] = pd.to_numeric(df['Matricula'], errors='coerce')
    df = df.dropna(subset=['Matricula'])
    df['Matricula'] = df['Matricula'].astype(int)
    
    # Processar faltas
    df['Faltas'] = df['Falta'].apply(contar_faltas)
    
    # Processar Ausência Parcial (atrasos)
    def converter_para_horas(tempo):
        if pd.isna(tempo) or tempo == '' or tempo == '00:00':
            return 0
        try:
            if ':' in str(tempo):
                horas, minutos = map(int, str(tempo).split(':'))
                return horas + minutos/60
            return 0
        except:
            return 0
    
    df['Horas_Atraso'] = df['Ausencia_Parcial'].apply(converter_para_horas)
    df['Afastamentos'] = df['Afastamentos'].fillna('').astype(str)
    
    # Agrupar por matrícula
    resultado = df.groupby('Matricula').agg({
        'Faltas': 'sum',
        'Horas_Atraso': 'sum',
        'Afastamentos': lambda x: '; '.join(sorted(set(filter(None, x))))
    }).reset_index()
    
    # Formatar horas de atraso
    resultado['Atrasos'] = resultado['Horas_Atraso'].apply(
        lambda x: f"{int(x)}:{int((x % 1) * 60):02d}" if x > 0 else ""
    )
    resultado = resultado.drop('Horas_Atraso', axis=1)
    
    return resultado

def calcular_premio(df_funcionarios, df_ausencias, data_limite_admissao):
    """Calcula prêmio baseado nas regras definidas"""
    # Lista de afastamentos que impedem o prêmio
    afastamentos_impeditivos = [
        "Declaração Acompanhante", "Feriado", "Emenda Feriado", 
        "Licença Maternidade", "Declaração INSS (dias)", 
        "Comparecimento Medico INSS", "Aposentado por Invalidez",
        "Atestado Médico", "Atestado de Óbito", "Licença Paternidade",
        "Licença Casamento", "Acidente de Trabalho", "Auxilio Doença",
        "Primeira Suspensão", "Segunda Suspensão", "Férias",
        "Falta não justificada", "Processo",
        "Falta não justificada (dias)", "Atestado Médico (dias)"
    ]
    
    # Afastamentos que precisam de decisão
    afastamentos_decisao = ["Abono", "Atraso"]
    
    # Afastamentos permitidos
    afastamentos_permitidos = [
        "Folga Gestor", "Abonado Gerencia Loja",
        "Confraternização universal", "Aniversario de São Paulo"
    ]
    
    # Filtrar funcionários pela data de admissão
    df_funcionarios = df_funcionarios[
        pd.to_datetime(df_funcionarios['Data_Admissao']) <= pd.to_datetime(data_limite_admissao)
    ]
    
    resultados = []
    for _, func in df_funcionarios.iterrows():
        # Buscar ausências do funcionário
        ausencias = df_ausencias[df_ausencias['Matricula'] == func['Matricula']]
        
        tem_afastamento_impeditivo = False
        tem_afastamento_decisao = False
        tem_apenas_permitidos = False
        
        if not ausencias.empty:
            afastamentos = ' '.join(ausencias['Afastamentos'].fillna('').astype(str)).lower()
            
            # Verificar afastamentos impeditivos
            for afastamento in afastamentos_impeditivos:
                if afastamento.lower() in afastamentos:
                    tem_afastamento_impeditivo = True
                    break
            
            # Verificar afastamentos que precisam de decisão
            if not tem_afastamento_impeditivo:
                for afastamento in afastamentos_decisao:
                    if afastamento.lower() in afastamentos:
                        tem_afastamento_decisao = True
                        break
            
            # Verificar se tem apenas afastamentos permitidos
            tem_apenas_permitidos = not tem_afastamento_impeditivo and not tem_afastamento_decisao
        
        # Calcular valor do prêmio
        valor_premio = 0
        if func['Qtd_Horas_Mensais'] == 220:
            valor_premio = 300.00
        elif func['Qtd_Horas_Mensais'] <= 110:
            valor_premio = 150.00
        
        # Determinar status
        status = "Não tem direito"
        if not tem_afastamento_impeditivo:
            if tem_afastamento_decisao:
                status = "Aguardando decisão"
                if not ausencias.empty and 'Atrasos' in ausencias:
                    atrasos = ausencias['Atrasos'].iloc[0]
                    if atrasos:
                        status = f"Aguardando decisão (Total Atrasos: {atrasos})"
            elif tem_apenas_permitidos or ausencias.empty:
                status = "Tem direito"
        
        # Adicionar ao resultado
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
        
        # Data limite
        data_limite = st.date_input(
            "Data Limite de Admissão",
            help="Funcionários admitidos após esta data não terão direito ao prêmio",
            format="DD/MM/YYYY"
        )
        
        # Uploads
        st.subheader("Base de Funcionários")
        uploaded_func = st.file_uploader("Carregar base de funcionários", type=['xlsx'])
        
        st.subheader("Base de Ausências")
        uploaded_ausencias = st.file_uploader("Carregar base de ausências", type=['xlsx'])
        
        # Upload dos tipos de afastamento
        st.subheader("Tipos de Afastamento")
        uploaded_tipos = st.file_uploader("Atualizar tipos de afastamento", type=['xlsx'])
        
        if uploaded_tipos is not None:
            try:
                df_tipos_novo = pd.read_excel(uploaded_tipos)
                if 'Nome' in df_tipos_novo.columns and 'Categoria' in df_tipos_novo.columns:
                    df_tipos = df_tipos_novo.rename(columns={'Nome': 'tipo', 'Categoria': 'categoria'})
                    salvar_tipos_afastamento(df_tipos)
                    st.success("Tipos de afastamento atualizados!")
                else:
                    st.error("Arquivo deve conter colunas 'Nome' e 'Categoria'")
            except Exception as e:
                st.error(f"Erro ao processar arquivo: {str(e)}")
    
    if uploaded_func and uploaded_ausencias and data_limite:
        try:
            # Carregar base de funcionários
            df_funcionarios = pd.read_excel(uploaded_func)
            df_funcionarios.columns = [
                "Matricula", "Nome_Funcionario", "Cargo", 
                "Codigo_Local", "Nome_Local", "Qtd_Horas_Mensais",
                "Tipo_Contrato", "Data_Termino_Contrato", 
                "Dias_Experiencia", "Salario_Mes_Atual", "Data_Admissao"
            ]
            
            # Converter datas
            df_funcionarios['Data_Admissao'] = pd.to_datetime(df_funcionarios['Data_Admissao'], format='%d/%m/%Y')
            df_funcionarios['Data_Termino_Contrato'] = pd.to_datetime(df_funcionarios['Data_Termino_Contrato'], format='%d/%m/%Y', errors='coerce')
            
            # Carregar base de ausências
            df_ausencias = pd.read_excel(uploaded_ausencias)
            df_ausencias = df_ausencias.rename(columns={
                "Matrícula": "Matricula",
                "Centro de Custo": "Centro_de_Custo",
                "Ausência Integral": "Ausencia_Integral",
                "Ausência Parcial": "Ausencia_Parcial",
                "Data de Demissão": "Data_de_Demissao"
            })
            df_ausencias = processar_ausencias(df_ausencias)
            
            # Calcular prêmios
            df_resultado = calcular_premio(df_funcionarios, df_ausencias, data_limite)

            # Mostrar resultado na tela
            st.subheader("Resultado do Cálculo de Prêmios")
            
            # Cards com métricas
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Analisados", len(df_resultado))
            with col2:
                st.metric("Com Direito", len(df_resultado[df_resultado['Status'] == 'Tem direito']))
            with col3:
                st.metric("Aguardando Decisão", 
                         len(df_resultado[df_resultado['Status'].str.contains('Aguardando decisão', na=False)]))
            with col4:
                st.metric("Valor Total", f"R$ {df_resultado['Valor_Premio'].sum():,.2f}")

            # Detalhamento por status
            for status in sorted(df_resultado['Status'].unique()):
                df_status = df_resultado[df_resultado['Status'] == status]
                with st.expander(f"Status: {status}", expanded=True):
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Quantidade de Funcionários", len(df_status))
                    with col2:
                        st.metric("Valor Total", f"R$ {df_status['Valor_Premio'].sum():,.2f}")
                    
                    st.subheader("Locais Afetados")
                    st.write(", ".join(sorted(df_status['Local'].unique())))
                    
                    st.subheader("Lista de Funcionários")
                    st.dataframe(
                        df_status[['Matricula', 'Nome', 'Cargo', 'Local', 'Valor_Premio']],
                        hide_index=True,
                        column_config={
                            "Matricula": st.column_config.NumberColumn("Matrícula", format="%d"),
                            "Nome": st.column_config.TextColumn("Nome"),
                            "Cargo": st.column_config.TextColumn("Cargo"),
                            "Local": st.column_config.TextColumn("Local"),
                            "Valor_Premio": st.column_config.NumberColumn("Valor Prêmio", format="R$ %.2f")
                        }
                    )

            # Geração de HTML para o relatório
            html_content = f"""
            <html>
                <head>
                    <style>
                        body {{
                            font-family: Arial, sans-serif;
                            margin: 20px;
                        }}
                        h1 {{
                            color: #1f77b4;
                            text-align: center;
                        }}
                        .resumo {{
                            margin: 20px 0;
                            padding: 10px;
                            background-color: #f8f9fa;
                            border-radius: 5px;
                        }}
                        table {{
                            width: 100%;
                            border-collapse: collapse;
                            margin-top: 20px;
                        }}
                        th, td {{
                            border: 1px solid #ddd;
                            padding: 12px;
                            text-align: left;
                        }}
                        th {{
                            background-color: #1f77b4;
                            color: white;
                        }}
                        tr:nth-child(even) {{
                            background-color: #f8f9fa;
                        }}
                    </style>
                </head>
                <body>
                    <h1>RELATÓRIO DE PRÊMIOS - VISÃO EXECUTIVA</h1>
                    <p style="text-align: right;">Data do relatório: {datetime.now().strftime('%d/%m/%Y')}</p>
                    
                    <div class="resumo">
                        <h2>Resumo Geral</h2>
                        <p>Total Analisados: {len(df_resultado)}</p>
                        <p>Com Direito: {len(df_resultado[df_resultado['Status'] == 'Tem direito'])}</p>
                        <p>Aguardando Decisão: {len(df_resultado[df_resultado['Status'].str.contains('Aguardando decisão', na=False)])}</p>
                        resultado['Valor_Premio'].sum():,.2f}</p>
                   </div>
           """
           
           # Adicionar seções por status
           for status in sorted(df_resultado['Status'].unique()):
               df_status = df_resultado[df_resultado['Status'] == status]
               html_content += f"""
                   <h2>Status: {status}</h2>
                   <p>Quantidade de Funcionários: {len(df_status)}</p>
                   <p>Valor Total: R$ {df_status['Valor_Premio'].sum():,.2f}</p>
                   <table>
                       <tr>
                           <th>Matrícula</th>
                           <th>Nome</th>
                           <th>Cargo</th>
                           <th>Local</th>
                           <th>Valor Prêmio</th>
                       </tr>
               """
               
               for _, row in df_status.iterrows():
                   html_content += f"""
                       <tr>
                           <td>{int(row['Matricula'])}</td>
                           <td>{row['Nome']}</td>
                           <td>{row['Cargo']}</td>
                           <td>{row['Local']}</td>
                           <td style="text-align: right;">R$ {row['Valor_Premio']:,.2f}</td>
                       </tr>
                   """
               
               html_content += "</table>"
           
           html_content += """
               </body>
           </html>
           """

           # Botão para exportar para PDF
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
                   st.error("Certifique-se de que o wkhtmltopdf está instalado no sistema.")

       except Exception as e:
           st.error(f"Erro ao processar dados: {str(e)}")

if __name__ == "__main__":
   main()
