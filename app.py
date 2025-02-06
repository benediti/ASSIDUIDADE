I'll modify the code to ensure logs are available for download:

```python
import streamlit as st
import pandas as pd
import unicodedata
import traceback
from io import BytesIO
import logging

# Configuração de logging
logging.basicConfig(level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s: %(message)s')

# Lista para armazenar logs
log_messages = []

def add_to_log(message, level='info'):
    """Adiciona mensagens ao log"""
    global log_messages
    log_entry = f"{level.upper()}: {message}"
    log_messages.append(log_entry)
    
    # Log usando o módulo logging
    if level == 'debug':
        logging.debug(message)
    elif level == 'warning':
        logging.warning(message)
    elif level == 'error':
        logging.error(message)
    else:
        logging.info(message)

def generate_log_file():
    """Gera um arquivo de log em memória"""
    log_data = "\n".join(log_messages)
    return BytesIO(log_data.encode('utf-8'))

# [Restante do código anterior permanece o mesmo, apenas adicionando log_messages em pontos críticos]

def main():
    st.title("Processador de Prêmio Assiduidade")
    
    base_file = st.file_uploader("Arquivo Base", type=['xlsx'], key='base')
    absence_file = st.file_uploader("Arquivo de Ausências", type=['xlsx'], key='ausencia')
    model_file = st.file_uploader("Modelo de Exportação", type=['xlsx'], key='modelo')

    # Área para exibir logs
    log_area = st.expander("Logs de Processamento")

    if base_file and absence_file and model_file:
        if st.button("Processar Dados"):
            # Limpa logs anteriores
            log_messages.clear()
            
            add_to_log("Iniciando processamento de dados")
            
            df_resultado = process_data(base_file, absence_file, model_file)
            
            if df_resultado is not None:
                st.success("Dados processados com sucesso!")
                
                # Exportação do resultado
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Resultado')
                output.seek(0)

                st.download_button("Baixar Relatório", data=output, file_name="resultado_assiduidade.xlsx")
                
                # Mostra log na interface
                with log_area:
                    for log in log_messages:
                        st.text(log)
                
                # Download de logs
                log_file = generate_log_file()
                st.download_button("Baixar Log Detalhado", 
                                   data=log_file, 
                                   file_name="log_processamento.txt", 
                                   key="download_log")
            else:
                st.error("Falha no processamento. Verifique os logs.")
                
                # Mostra log na interface em caso de erro
                with log_area:
                    for log in log_messages:
                        st.text(log)
                
                # Download de logs em caso de erro
                log_file = generate_log_file()
                st.download_button("Baixar Log de Erro", 
                                   data=log_file, 
                                   file_name="log_erro.txt", 
                                   key="download_error_log")

if __name__ == "__main__":
    main()
```

Principais mudanças:
- Adicionei uma área expansível para logs
- Limpa logs anteriores antes de cada processamento
- Mostra logs na interface
- Garante download de logs em caso de sucesso ou erro
- Adiciona log de início do processamento

Isso deve resolver o problema de visualização e download dos logs.
