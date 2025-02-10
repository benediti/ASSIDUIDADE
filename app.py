import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io
import logging
from reportlab.platypus import PageBreak

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
    # Renomear colunas com acentos
    df = df.rename(columns={
        "Matrícula": "Matricula",
        "Centro de Custo": "Centro_de_Custo",
        "Ausência Integral": "Ausencia_Integral",
        "Ausência Parcial": "Ausencia_Parcial",
        "Data de Demissão": "Data_de_Demissao"
    })
    
    # Converter Matricula para inteiro
    df['Matricula'] = pd.to_numeric(df['Matricula'], errors='coerce')
    df = df.dropna(subset=['Matricula'])
    df['Matricula'] = df['Matricula'].astype(int)
    
    # Processar faltas (X = 1, vazio = 0)
    df['Faltas'] = df['Falta'].fillna('')
    df['Faltas'] = df['Faltas'].apply(lambda x: 1 if str(x).upper().strip() == 'X' else 0)
    
    # Processar atrasos
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
    
    # Processar tipos de afastamento
    df_tipos = carregar_tipos_afastamento()
    for tipo in df_tipos['tipo'].unique():
        resultado[tipo] = resultado['Afastamentos'].str.contains(tipo, case=False).astype(int)
        resultado[tipo] = resultado[tipo].apply(lambda x: x if x > 0 else "")
    
    return resultado

def calcular_premio(df_funcionarios, df_ausencias, data_limite_admissao):
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
    
    # Filtrar pela data de admissão
    df_funcionarios['Data_Admissao'] = pd.to_datetime(df_funcionarios['Data_Admissao'], format='%d/%m/%Y')
    df_funcionarios = df_funcionarios[df_funcionarios['Data_Admissao'] <= pd.to_datetime(data_limite_admissao)]
    
    resultados = []
    for _, func in df_funcionarios.iterrows():
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
        total_atrasos = ""
        
        if not tem_afastamento_impeditivo:
            if tem_afastamento_decisao:
                status = "Aguardando decisão"
                if not ausencias.empty and 'Atrasos' in ausencias.columns:
                    total_atrasos = ausencias['Atrasos'].iloc[0]
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
            'Status': f"{status} (Total Atrasos: {total_atrasos})" if status == "Aguardando decisão" and total_atrasos else status,
            'Detalhes_Afastamentos': ausencias['Afastamentos'].iloc[0] if not ausencias.empty else ''
        })
    
    return pd.DataFrame(resultados)

def main():
    st.set_page_config(page_title="Sistema de Verificação de Prêmios", page_icon="🏆", layout="wide")
    st.title("Sistema de Verificação de Prêmios")
    
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
            # Carregar base de funcionários
            df_funcionarios = pd.read_excel(uploaded_func)
            df_funcionarios.columns = [
                "Matricula", "Nome_Funcionario", "Cargo", 
                "Codigo_Local", "Nome_Local", "Qtd_Horas_Mensais",
                "Tipo_Contrato", "Data_Termino_Contrato", 
                "Dias_Experiencia", "Salario_Mes_Atual", "Data_Admissao"
            ]
            
            # Carregar e processar base de ausências
            df_ausencias = pd.read_excel(uploaded_ausencias)
            df_ausencias = processar_ausencias(df_ausencias)
            
            # Calcular prêmios
            df_resultado = calcular_premio(df_funcionarios, df_ausencias, data_limite)
            
            st.subheader("Resultado do Cálculo de Prêmios")
            
            # Filtros
            col1, col2 = st.columns(2)
            with col1:
                filtro_status = st.multiselect(
                    "Filtrar por Status",
                    options=sorted(df_resultado['Status'].unique())
                )
            
            with col2:
                filtro_local = st.multiselect(
                    "Filtrar por Local",
                    options=sorted(df_resultado['Local'].unique())
                )
            
            # Aplicar filtros
            df_mostrar = df_resultado
            if filtro_status:
                df_mostrar = df_mostrar[df_mostrar['Status'].isin(filtro_status)]
            if filtro_local:
                df_mostrar = df_mostrar[df_mostrar['Local'].isin(filtro_local)]
            
            # Relatório formatado na interface
            st.markdown("---")
            st.subheader("Relatório Executivo", divider="rainbow")
            
            # Criar conteúdo HTML para o relatório
            html_content = f"""
            <div style="font-family: Arial; padding: 20px;">
                <h1 style="color: #1f77b4; text-align: center;">RELATÓRIO DE PRÊMIOS - VISÃO EXECUTIVA</h1>
                <p style="text-align: right; font-size: 14px;">Data do relatório: {datetime.now().strftime('%d/%m/%Y')}</p>
                
                <div style="display: flex; justify-content: space-between; margin: 20px 0;">
                    <div style="text-align: center; padding: 10px; background-color: #f8f9fa; border-radius: 5px;">
                        <h3>Total Analisados</h3>
                        <h2>{len(df_mostrar):,}</h2>
                    </div>
                    <div style="text-align: center; padding: 10px; background-color: #f8f9fa; border-radius: 5px;">
                        <h3>Com Direito</h3>
                        <h2>{len(df_mostrar[df_mostrar['Status'] == 'Tem direito']):,}</h2>
                    </div>
                    <div style="text-align: center; padding: 10px; background-color: #f8f9fa; border-radius: 5px;">
                        <h3>Aguardando Decisão</h3>
                        <h2>{len(df_mostrar[df_mostrar['Status'].str.contains('Aguardando decisão', na=False)]):,}</h2>
                    </div>
                    <div style="text-align: center; padding: 10px; background-color: #f8f9fa; border-radius: 5px;">
                        <h3>Valor Total</h3>
                        <h2>R$ {df_mostrar['Valor_Premio'].sum():,.2f}</h2>
                    </div>
                </div>
            """
            
            # Adicionar detalhamento por status
            for status in sorted(df_mostrar['Status'].unique()):
                df_status = df_mostrar[df_mostrar['Status'] == status]
                html_content += f"""
                <div style="margin: 20px 0; padding: 15px; background-color: #f8f9fa; border-radius: 5px;">
                    <h2 style="color: #1f77b4;">Status: {status}</h2>
                    <p><strong>Quantidade de Funcionários:</strong> {len(df_status):,}</p>
                    <p><strong>Valor Total:</strong> R$ {df_status['Valor_Premio'].sum():,.2f}</p>
                    <p><strong>Locais Afetados:</strong></p>
                    <p>{', '.join(sorted(df_status['Local'].unique()))}</p>
                    
                    <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
                        <tr style="background-color: #1f77b4; color: white;">
                            <th style="padding: 8px; text-align: left;">Matrícula</th>
                            <th style="padding: 8px; text-align: left;">Nome</th>
                            <th style="padding: 8px; text-align: left;">Cargo</th>
                            <th style="padding: 8px; text-align: left;">Local</th>
                            <th style="padding: 8px; text-align: right;">Valor Prêmio</th>
                        </tr>
                """
                
                for _, row in df_status.iterrows():
                    html_content += f"""
                        <tr style="border-bottom: 1px solid #ddd;">
                            <td style="padding: 8px;">{int(row['Matricula'])}</td>
                            <td style="padding: 8px;">{row['Nome']}</td>
                            <td style="padding: 8px;">{row['Cargo']}</td>
                            <td style="padding: 8px;">{row['Local']}</td>
                            <td style="padding: 8px; text-align: right;">R$ {row['Valor_Premio']:,.2f}</td>
                        </tr>
                    """
                
                html_content += """
                    </table>
                </div>
                """
            
            html_content += "</div>"
            
            # Mostrar relatório na interface
            st.write(html_content, unsafe_allow_html=True)
            
            # Exportar para PDF
            if st.button("📑 Exportar Relatório como PDF"):
                try:
                    from reportlab.lib import colors
                    from reportlab.lib.pagesizes import A4
                    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
                    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                    from io import BytesIO
                    
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
                    styles = getSampleStyleSheet()
                    
                    # Estilos personalizados
                    styles.add(ParagraphStyle(
                        'CustomTitle',
                        parent=styles['Heading1'],
                        fontSize=24,
                        spaceAfter=30,
                        alignment=1,
                        textColor=colors.HexColor('#1f77b4')
                    ))
                    
                    styles.add(ParagraphStyle(
                        'SectionHeader',
                        parent=styles['Heading2'],
                        fontSize=16,
                        spaceBefore=15,
                        spaceAfter=10,
                        textColor=colors.HexColor('#2c3e50')
                    ))
                    
                    # Título e data
                    story.append(Paragraph("RELATÓRIO DE PRÊMIOS - VISÃO EXECUTIVA", styles['CustomTitle']))
                    story.append(Paragraph(
                        f"Data do relatório: {datetime.now().strftime('%d/%m/%Y')}",
                        ParagraphStyle(
                            'Date',
                            parent=styles['Normal'],
                            alignment=2,
                            fontSize=12,
                            textColor=colors.grey
                        )
                    ))
                    story.append(Spacer(1, 20))
                    
                    # Resumo geral
                    resumo_data = [
                        ['RESUMO GERAL'],
                        [f'Total Analisados: {len(df_mostrar):,}'],
                        [f'Com Direito: {len(df_mostrar[df_mostrar["Status"] == "Tem direito"]):,}'],
                        [f'Aguardando Decisão: {len(df_mostrar[df_mostrar["Status"].str.contains("Aguardando decisão", na=False)]):,}'],
                        [f'Valor Total: R$ {df_mostrar["Valor_Premio"].sum():,.2f}']
                    ]
                    
                    t = Table(resumo_data, colWidths=[480])
                    t.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 14),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8f9fa')),
                        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 1), (-1, -1), 12),
                        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
                        ('ROWHEIGHT', (0, 0), (-1, -1), 30),
                    ]))
                    story.append(t)
                    story.append(Spacer(1, 30))
                    
                    # Detalhamento por status
                    for status in sorted(df_mostrar['Status'].unique()):
                        df_status = df_mostrar[df_mostrar['Status'] == status]
                        
                        story.append(Paragraph(f'Status: {status}', styles['SectionHeader']))
                        
                        info_data = [
                            [f'Quantidade de Funcionários: {len(df_status):,}'],
                            [f'Valor Total: R$ {df_status["Valor_Premio"].sum():,.2f}'],
                            ['Locais Afetados:'],
                            [', '.join(sorted(df_status['Local'].unique()))]
                        ]
                        
                        t = Table(info_data, colWidths=[480])
                        t.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#f8f9fa')),
                            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                            ('FONTSIZE', (0, 0), (-1, -1), 10),
                            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                            ('GRID', (0, 0), (-1, -1), 1, colors.grey),
                            ('ROWHEIGHT', (0, 0), (-1, -1), 25),
                        ]))
                        story.append(t)
                        story.append(Spacer(1, 10))
                        
                        if len(df_status) > 0:
                            data = [['Matrícula', 'Nome', 'Cargo', 'Local', 'Valor Prêmio']]
                            for _, row in df_status.iterrows():
                                data.append([
                                    str(int(row['Matricula'])),
                                    row['Nome'],
                                    row['Cargo'],
                                    row['Local'],
                                    f'R$ {row["Valor_Premio"]:,.2f}'
                                ])
                            
                            col_widths = [60, 140, 100, 120, 60]
                            t = Table(data, colWidths=col_widths)
                            t.setStyle(TableStyle([
                                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                ('FONTSIZE', (0, 0), (-1, 0), 10),
                                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                                ('FONTSIZE', (0, 1), (-1, -1), 8),
                                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                                ('ALIGN', (-1, 0), (-1, -1), 'RIGHT'),
                                ('GRID', (0, 0), (-1, -1), 1, colors.grey),
                                ('ROWHEIGHT', (0, 0), (-1, -1), 20),
                            ]))
                            story.append(t)
                        
                        story.append(PageBreak())
                    
                    doc.build(story)
                    
                    # Resumo Geral
                    data = [
                        ["RESUMO GERAL"],
                        [f"Total Analisados: {len(df_mostrar):,}"],
                        [f"Com Direito: {len(df_mostrar[df_mostrar['Status'] == 'Tem direito']):,}"],
                        [f"Aguardando Decisão: {len(df_mostrar[df_mostrar['Status'].str.contains('Aguardando decisão', na=False)]):,}"],
                        [f"Valor Total: R$ {df_mostrar['Valor_Premio'].sum():,.2f}"]
                    ]
                    
                    t = Table(data, colWidths=[450])
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 14),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 1), (-1, -1), 12),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)
                    ]))
                    story.append(t)
                    story.append(Spacer(1, 20))
                    
                    # Detalhamento por status
                    for status in sorted(df_mostrar['Status'].unique()):
                        df_status = df_mostrar[df_mostrar['Status'] == status]
                        
                        story.append(Paragraph(f"Status: {status}", styles['Heading2']))
                        story.append(Spacer(1, 10))
                        
                        data = [
                            [f"Quantidade de Funcionários: {len(df_status):,}"],
                            [f"Valor Total: R$ {df_status['Valor_Premio'].sum():,.2f}"],
                            ["Locais Afetados:"],
                            [', '.join(sorted(df_status['Local'].unique()))]
                        ]
                        
                        t = Table(data, colWidths=[450])
                        t.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, -1), colors.lightgrey),
                            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                            ('FONTSIZE', (0, 0), (-1, -1), 10),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ]))
                        story.append(t)
                        story.append(Spacer(1, 10))
                        
                        # Lista de funcionários
                        if len(df_status) > 0:
                            data = [[
                                "Matrícula", "Nome", "Cargo", "Local", "Valor Prêmio"
                            ]]
                            for _, row in df_status.iterrows():
                                data.append([
                                    str(int(row['Matricula'])),
                                    row['Nome'],
                                    row['Cargo'],
                                    row['Local'],
                                    f"R$ {row['Valor_Premio']:,.2f}"
                                ])
                            
                            t = Table(data, colWidths=[60, 120, 100, 100, 70])
                            t.setStyle(TableStyle([
                                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                ('FONTSIZE', (0, 0), (-1, 0), 10),
                                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                                ('ALIGN', (-1, 0), (-1, -1), 'RIGHT'),
                                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                                ('FONTSIZE', (0, 1), (-1, -1), 8),
                                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER')
                            ]))
                            story.append(t)
                        
                        story.append(Spacer(1, 20))
                    
                    doc.build(story)
                    
                    # Botão de download
                    st.download_button(
                        label="⬇️ Download PDF",
                        data=buffer.getvalue(),
                        file_name="relatorio_premios.pdf",
                        mime="application/pdf"
                    )
                    
                except Exception as e:
                    st.error(f"Erro ao gerar PDF: {str(e)}")

            
            # Estatísticas
            st.subheader("Estatísticas")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total de Funcionários", len(df_mostrar))
            with col2:
                st.metric("Total com Direito", len(df_mostrar[df_mostrar['Status'] == "Tem direito"]))
            with col3:
                st.metric("Valor Total Prêmios", f"R$ {df_mostrar['Valor_Premio'].sum():.2f}")
            
            # Exportar resultados
            if st.button("Exportar Resultados"):
                output = io.BytesIO()
                
                # Preparar dados para exportação
                df_export = df_mostrar.copy()
                df_export['Salario'] = df_funcionarios.set_index('Matricula').loc[df_export['Matricula'], 'Salario_Mes_Atual'].values
                
                # Exportar planilha detalhada
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Resultados Detalhados')
                    
                    # Criar relatório para diretoria
                    relatorio_diretoria = pd.DataFrame([
                        ["RELATÓRIO DE PRÊMIOS - VISÃO EXECUTIVA", ""],
                        [f"Data do relatório: {datetime.now().strftime('%d/%m/%Y')}", ""],
                        ["", ""],
                        ["RESUMO GERAL", ""],
                        [f"Total de Funcionários Analisados: {len(df_export)}", ""],
                        [f"Funcionários com Direito: {len(df_export[df_export['Status'] == 'Tem direito'])}", ""],
                        [f"Funcionários Aguardando Decisão: {len(df_export[df_export['Status'].str.contains('Aguardando decisão', na=False)])}", ""],
                        [f"Valor Total dos Prêmios: R$ {df_export['Valor_Premio'].sum():,.2f}", ""],
                        ["", ""],
                        ["DETALHAMENTO POR STATUS", ""],
                    ])
                    
                    # Adicionar detalhamento por status
                    for status in df_export['Status'].unique():
                        df_status = df_export[df_export['Status'] == status]
                        relatorio_diretoria = pd.concat([relatorio_diretoria, pd.DataFrame([
                            [f"\nStatus: {status}", ""],
                            [f"Quantidade de Funcionários: {len(df_status)}", ""],
                            [f"Valor Total: R$ {df_status['Valor_Premio'].sum():,.2f}", ""],
                            ["Locais Afetados:", ""],
                            [", ".join(df_status['Local'].unique()), ""],
                            ["", ""]
                        ])])
                    
                    # Salvar relatório em nova aba
                    relatorio_diretoria.to_excel(writer, index=False, header=False, sheet_name='Relatório Executivo')
                
                st.download_button(
                    label="Download Excel",
                    data=output.getvalue(),
                    file_name="resultado_premios.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        except Exception as e:
            st.error(f"Erro ao processar dados: {str(e)}")

if __name__ == "__main__":
    main()
