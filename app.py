import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io
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

def contar_faltas(valor):
    if isinstance(valor, str) and valor.strip().upper() == 'X':
        return 1
    return 0

def verificar_estrutura_dados_funcionarios(df):
    colunas_esperadas = {
        "Matricula": "numeric",
        "Nome_Funcionario": "string",
        "Cargo": "string",
        "Codigo_Local": "numeric",
        "Nome_Local": "string",
        "Qtd_Horas_Mensais": "numeric",
        "Tipo_Contrato": "string",
        "Data_Termino_Contrato": "datetime",
        "Dias_Experiencia": "numeric",
        "Salario_Mes_Atual": "numeric",
        "Data_Admissao": "datetime"
    }
    
    colunas_permitir_nulos = ["Data_Termino_Contrato", "Dias_Experiencia"]
    info = []
    erros = []
    
    for coluna in colunas_esperadas.keys():
        if coluna not in df.columns:
            erros.append(f"Coluna ausente: {coluna}")
    
    if erros:
        return False, erros
    
    for coluna, tipo in colunas_esperadas.items():
        try:
            nulos = df[coluna].isnull().sum()
            if nulos > 0:
                if coluna in colunas_permitir_nulos:
                    info.append(f"Informação: Coluna {coluna} contém {nulos} valores em branco (permitido)")
                else:
                    erros.append(f"Coluna {coluna} contém {nulos} valores nulos")
            
            if tipo == "numeric":
                if coluna in colunas_permitir_nulos:
                    df[coluna] = pd.to_numeric(df[coluna], errors='coerce')
                else:
                    df[coluna] = pd.to_numeric(df[coluna], errors='raise')
            elif tipo == "datetime":
                if coluna in colunas_permitir_nulos:
                    df[coluna] = pd.to_datetime(df[coluna], format='%d/%m/%Y', errors='coerce')
                else:
                    df[coluna] = pd.to_datetime(df[coluna], format='%d/%m/%Y')
            elif tipo == "string":
                df[coluna] = df[coluna].astype(str)
        except Exception as e:
            erros.append(f"Erro na coluna {coluna}: {str(e)}")
    
    return len(erros) == 0, info + erros

def verificar_estrutura_dados_ausencias(df):
    colunas_esperadas = {
        "Matricula": "numeric",
        "Nome": "string",
        "Centro_de_Custo": "numeric",
        "Dia": "datetime",
        "Ausencia_Integral": "string",
        "Ausencia_Parcial": "string",
        "Afastamentos": "string",
        "Falta": "string",
        "Data_de_Demissao": "datetime"
    }
    
    info = []
    erros = []
    
    for coluna in colunas_esperadas.keys():
        if coluna not in df.columns:
            erros.append(f"Coluna ausente: {coluna}")
    
    if not erros:
        try:
            df = df[df['Matricula'].notna() & (df['Matricula'].astype(str).str.strip() != '')]
            df['Matricula'] = df['Matricula'].astype(int)
            df['Centro_de_Custo'] = pd.to_numeric(df['Centro_de_Custo'], errors='coerce')
            df['Dia'] = pd.to_datetime(df['Dia'], format='%d/%m/%Y', errors='coerce')
            df['Data_de_Demissao'] = pd.to_datetime(df['Data_de_Demissao'], format='%d/%m/%Y', errors='coerce')
            
            df['Falta'] = df['Falta'].fillna('')
            df['Falta'] = df['Falta'].apply(lambda x: 1 if x.lower() == 'x' else 0)
            
            info.append("Estrutura de dados validada com sucesso!")
            
        except Exception as e:
            erros.append(f"Erro ao processar dados: {str(e)}")
    
    return len(erros) == 0, info + erros

def processar_ausencias(df):
    df['Matricula'] = pd.to_numeric(df['Matricula'], errors='coerce')
    df = df.dropna(subset=['Matricula'])
    df['Matricula'] = df['Matricula'].astype(int)
    
    df['Faltas'] = df['Falta'].apply(contar_faltas)
    
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
    
    resultado = df.groupby('Matricula').agg({
        'Faltas': 'sum',
        'Horas_Atraso': 'sum',
        'Afastamentos': lambda x: '; '.join(sorted(set(filter(None, x))))
    }).reset_index()
    
    resultado['Atrasos'] = resultado['Horas_Atraso'].apply(
        lambda x: f"{int(x)}:{int((x % 1) * 60):02d}" if x > 0 else ""
    )
    resultado = resultado.drop('Horas_Atraso', axis=1)
    
    df_tipos = carregar_tipos_afastamento()
    for tipo in df_tipos['tipo'].unique():
        df[f'count_{tipo}'] = df['Afastamentos'].str.contains(tipo, case=False).astype(int)
        contagem = df.groupby('Matricula')[f'count_{tipo}'].sum().reset_index()
        resultado = resultado.merge(contagem, on='Matricula', how='left')
        resultado = resultado.rename(columns={f'count_{tipo}': tipo})
        resultado[tipo] = resultado[tipo].apply(lambda x: x if x > 0 else "")
    
    return resultado

def calcular_premio(df_funcionarios, df_ausencias, data_limite_admissao):
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
    
    afastamentos_decisao = ["Abono", "Atraso"]
    
    afastamentos_permitidos = [
        "Folga Gestor", "Abonado Gerencia Loja",
        "Confraternização universal", "Aniversario de São Paulo"
    ]
    
    df_funcionarios = df_funcionarios[
        pd.to_datetime(df_funcionarios['Data_Admissao']) <= pd.to_datetime(data_limite_admissao)
    ]
    
    resultados = []
    for _, func in df_funcionarios.iterrows():
        ausencias = df_ausencias[df_ausencias['Matricula'] == func['Matricula']]
        
        tem_afastamento_impeditivo = False
        tem_afastamento_decisao = False
        tem_apenas_permitidos = False
        
        if not ausencias.empty:
            afastamentos = ' '.join(ausencias['Afastamentos'].fillna('').astype(str)).lower()
            
            for afastamento in afastamentos_impeditivos:
                if afastamento.lower() in afastamentos:
                    tem_afastamento_impeditivo = True
                    break
            
            if not tem_afastamento_impeditivo:
                for afastamento in afastamentos_decisao:
                    if afastamento.lower() in afastamentos:
                        tem_afastamento_decisao = True
                        break
            
            tem_apenas_permitidos = not tem_afastamento_impeditivo and not tem_afastamento_decisao
        
        valor_premio = 0
        if func['Qtd_Horas_Mensais'] == 220:
            valor_premio = 300.00
        elif func['Qtd_Horas_Mensais'] <= 110:
            valor_premio = 150.00
        
        status = "Não tem direito"
        if not tem_afastamento_impeditivo:
            if tem_afastamento_decisao:
                status = "Aguardando decisão"
                if not ausencias.empty and 'Atrasos' in ausencias.columns:
                    atrasos = ausencias['Atrasos'].iloc[0]
                    if atrasos:
                        status = f"Aguardando decisão (Total Atrasos: {atrasos})"
            elif tem_apenas_permitidos or ausencias.empty:
                status = "Tem direito"
        
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
    
    st.title("RELATÓRIO DE PRÊMIOS - VISÃO EXECUTIVA")
    st.caption(f"Data do relatório: {datetime.now().strftime('%d/%m/%Y')}")
    
    with st.sidebar:
        st.header("Configurações")
        
        data_limite = st.date_input(
            "Data Limite de Admissão",
            help="Funcionários admitidos após esta data não terão direito ao prêmio",
            format="DD/MM/YYYY"
        )
        
        st.subheader("Base de Funcionários")
        uploaded_func = st.file_uploader("Carregar base de funcionários", type=['xlsx'])
        
        st.subheader("Base de Ausências")
        uploaded_ausencias = st.file_uploader("Carregar base de ausências", type=['xlsx'])
        
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

    if uploaded_func is not None and uploaded_ausencias is not None and data_limite is not None:
        try:
            # Carregar e processar base de funcionários
            df_funcionarios = pd.read_excel(uploaded_func)
            df_funcionarios.columns = [
                "Matricula", "Nome_Funcionario", "Cargo", 
                "Codigo_Local", "Nome_Local", "Qtd_Horas_Mensais",
                "Tipo_Contrato", "Data_Termino_Contrato", 
                "Dias_Experiencia", "Salario_Mes_Atual", "Data_Admissao"
            ]
            
            df_funcionarios['Data_Admissao'] = pd.to_datetime(df_funcionarios['Data_Admissao'], format='%d/%m/%Y')
            df_funcionarios['Data_Termino_Contrato'] = pd.to_datetime(df_funcionarios['Data_Termino_Contrato'], format='%d/%m/%Y', errors='coerce')
            
            sucesso_func, msg_func = verificar_estrutura_dados_funcionarios(df_funcionarios)
            with st.expander("Log Base Funcionários", expanded=True):
                for msg in msg_func:
                    if msg.startswith("Informação:"):
                        st.info(msg)
                    elif sucesso_func:
                        st.success(msg)
                    else:
                        st.error(msg)
            
            # Carregar e processar base de ausências
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
            
            st.markdown("---")
            
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
            
            # Botão para exportar PDF
            if st.button("📑 Exportar Relatório como PDF"):
                from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
                from reportlab.lib import colors
                from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                from reportlab.lib.pagesizes import A4
                from io import BytesIO
                
                try:
                    # Criar PDF
                    buffer = BytesIO()
                    doc = SimpleDocTemplate(
                        buffer,
                        pagesize=A4,
                        rightMargin=30,
                        leftMargin=30,
                        topMargin=30,
                        bottomMargin=30
                    )
                    story = []
                    
                    # Criar estilos
                    styles = getSampleStyleSheet()
                    title_style = ParagraphStyle(
                        'CustomTitle',
                        parent=styles['Heading1'],
                        fontSize=24,
                        alignment=1,
                        spaceAfter=30
                    )
                    
                    # Título
                    story.append(Paragraph("RELATÓRIO DE PRÊMIOS - VISÃO EXECUTIVA", title_style))
                    story.append(Spacer(1, 20))
                    
                    # Resumo Geral
                    resumo_data = [
                        ["RESUMO GERAL"],
                        [f"Total Analisados: {len(df_resultado)}"],
                        [f"Com Direito: {len(df_resultado[df_resultado['Status'] == 'Tem direito'])}"],
                        [f"Aguardando Decisão: {len(df_resultado[df_resultado['Status'].str.contains('Aguardando decisão', na=False)])}"],
                        [f"Valor Total: R$ {df_resultado['Valor_Premio'].sum():,.2f}"]
                    ]
                    
                    t = Table(resumo_data, colWidths=[450])
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 14),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)
                    ]))
                    story.append(t)
                    story.append(Spacer(1, 20))
                    
                    # Detalhamento por status
                    for status in sorted(df_resultado['Status'].unique()):
                        df_status = df_resultado[df_resultado['Status'] == status]
                        
                        story.append(Paragraph(f"Status: {status}", styles['Heading2']))
                        
                        info_data = [
                            [f"Quantidade de Funcionários: {len(df_status)}"],
                            [f"Valor Total: R$ {df_status['Valor_Premio'].sum():,.2f}"],
                            ["Locais Afetados:"],
                            [", ".join(sorted(df_status['Local'].unique()))]
                        ]
                        
                        t = Table(info_data, colWidths=[450])
                        t.setStyle(TableStyle([
                            ('GRID', (0, 0), (-1, -1), 1, colors.black),
                            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                            ('FONTSIZE', (0, 0), (-1, -1), 10)
                        ]))
                        story.append(t)
                        story.append(Spacer(1, 10))
                        
                        if len(df_status) > 0:
                            funcionarios_data = [[
                                "Matrícula", "Nome", "Cargo", "Local", "Valor Prêmio"
                            ]]
                            for _, row in df_status.iterrows():
                                funcionarios_data.append([
                                    str(int(row['Matricula'])),
                                    row['Nome'],
                                    row['Cargo'],
                                    row['Local'],
                                    f"R$ {row['Valor_Premio']:,.2f}"
                                ])
                            
                            t = Table(funcionarios_data, colWidths=[60, 100, 90, 140, 60])
                            t.setStyle(TableStyle([
                                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                ('FONTSIZE', (0, 0), (-1, -1), 8),
                                ('ALIGN', (-1, 0), (-1, -1), 'RIGHT')
                            ]))
                            story.append(t)
                        
                        story.append(PageBreak())
                    
                    doc.build(story)
                    
                    st.download_button(
                        "⬇️ Download PDF",
                        data=buffer.getvalue(),
                        file_name="relatorio_premios.pdf",
                        mime="application/pdf"
                    )
                
                except Exception as e:
                    st.error(f"Erro ao gerar PDF: {str(e)}")
        
        except Exception as e:
            st.error(f"Erro ao processar dados: {str(e)}")

if __name__ == "__main__":
    main()
