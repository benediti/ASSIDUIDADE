import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
import os

# Ignora avisos do pandas
warnings.filterwarnings('ignore')

def converter_data_br_para_datetime(data_str):
    """
    Converte uma string de data no formato brasileiro (DD/MM/YYYY) para um objeto datetime.
    Retorna None se a conversão falhar.
    """
    if pd.isna(data_str) or data_str == '':
        return None
    
    if isinstance(data_str, datetime):
        return data_str
        
    try:
        # Se for uma string, tenta converter
        if isinstance(data_str, str):
            # Tenta converter do formato brasileiro DD/MM/YYYY
            return datetime.strptime(data_str, '%d/%m/%Y')
        else:
            return data_str  # Se já for outro tipo (como timestamp), retorna como está
    except (ValueError, TypeError):
        try:
            # Caso falhe, tenta outros formatos comuns
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y']:
                try:
                    return datetime.strptime(str(data_str), fmt)
                except ValueError:
                    continue
            return None
        except Exception:
            return None

def carregar_arquivo_ausencias(caminho_arquivo):
    """
    Carrega o arquivo de ausências e converte colunas de data corretamente.
    """
    try:
        # Lê o arquivo Excel
        df = pd.read_excel(caminho_arquivo)
        
        # Converte as colunas de data
        if 'Dia' in df.columns:
            df['Dia'] = df['Dia'].astype(str)
            df['Dia'] = df['Dia'].apply(converter_data_br_para_datetime)
        
        if 'Data de Demissão' in df.columns:
            df['Data de Demissão'] = df['Data de Demissão'].astype(str)
            df['Data de Demissão'] = df['Data de Demissão'].apply(converter_data_br_para_datetime)
        
        return df
    except Exception as e:
        print(f"Erro ao carregar arquivo de ausências: {e}")
        return pd.DataFrame()

def carregar_arquivo_funcionarios(caminho_arquivo):
    """
    Carrega o arquivo de funcionários e converte colunas de data corretamente.
    """
    try:
        # Lê o arquivo Excel
        df = pd.read_excel(caminho_arquivo)
        
        # Converte as colunas de data
        if 'Data Término Contrato' in df.columns:
            df['Data Término Contrato'] = df['Data Término Contrato'].astype(str)
            df['Data Término Contrato'] = df['Data Término Contrato'].apply(converter_data_br_para_datetime)
        
        if 'Data Admissão' in df.columns:
            df['Data Admissão'] = df['Data Admissão'].astype(str)
            df['Data Admissão'] = df['Data Admissão'].apply(converter_data_br_para_datetime)
        
        return df
    except Exception as e:
        print(f"Erro ao carregar arquivo de funcionários: {e}")
        return pd.DataFrame()

def filtrar_ausencias_por_periodo(df_ausencias, data_inicio, data_fim):
    """
    Filtra ausências para um determinado período.
    
    Args:
        df_ausencias: DataFrame com ausências
        data_inicio: Data de início do período (datetime ou string no formato DD/MM/YYYY)
        data_fim: Data de fim do período (datetime ou string no formato DD/MM/YYYY)
    
    Returns:
        DataFrame filtrado com ausências no período
    """
    # Converte data_inicio e data_fim para datetime se forem strings
    if isinstance(data_inicio, str):
        data_inicio = converter_data_br_para_datetime(data_inicio)
    if isinstance(data_fim, str):
        data_fim = converter_data_br_para_datetime(data_fim)
    
    # Verifica se a conversão ocorreu com sucesso
    if data_inicio is None or data_fim is None:
        print("Erro: formato de data inválido.")
        return pd.DataFrame()
    
    # Filtra as ausências dentro do período
    try:
        ausencias_periodo = df_ausencias[
            (df_ausencias['Dia'] >= data_inicio) & 
            (df_ausencias['Dia'] <= data_fim)
        ]
        return ausencias_periodo
    except Exception as e:
        print(f"Erro ao filtrar ausências por período: {e}")
        return pd.DataFrame()

def processar_dados(arquivo_ausencias, arquivo_funcionarios, mes=None, ano=None):
    """
    Processa os dados de ambos os arquivos com tratamento correto de datas.
    
    Args:
        arquivo_ausencias: Caminho para o arquivo de ausências
        arquivo_funcionarios: Caminho para o arquivo de funcionários
        mes: Mês para filtrar (1-12)
        ano: Ano para filtrar (ex: 2025)
    
    Returns:
        DataFrame com os dados processados
    """
    # Carrega os dados
    df_ausencias = carregar_arquivo_ausencias(arquivo_ausencias)
    df_funcionarios = carregar_arquivo_funcionarios(arquivo_funcionarios)
    
    if df_ausencias.empty or df_funcionarios.empty:
        print("Erro: Um ou mais arquivos não puderam ser carregados corretamente.")
        return pd.DataFrame()
    
    # Se mês e ano forem fornecidos, filtra para esse período
    if mes is not None and ano is not None:
        # Determine o primeiro e último dia do mês
        primeiro_dia = datetime(ano, mes, 1)
        if mes == 12:
            ultimo_dia = datetime(ano + 1, 1, 1) - timedelta(days=1)
        else:
            ultimo_dia = datetime(ano, mes + 1, 1) - timedelta(days=1)
        
        # Filtra as ausências para o mês/ano especificado
        df_ausencias = filtrar_ausencias_por_periodo(df_ausencias, primeiro_dia, ultimo_dia)
    
    # Mescla os dados de ausências com os dados de funcionários
    # Usando a "Matricula" como chave de junção
    if 'Matricula' in df_ausencias.columns and 'Matrícula' in df_funcionarios.columns:
        # Converte a coluna Matrícula para o mesmo tipo em ambos os DataFrames
        df_ausencias['Matricula'] = df_ausencias['Matricula'].astype(str)
        df_funcionarios['Matrícula'] = df_funcionarios['Matrícula'].astype(str)
        
        # Mescla os DataFrames
        df_combinado = pd.merge(
            df_ausencias,
            df_funcionarios,
            left_on='Matricula',
            right_on='Matrícula',
            how='left'
        )
        
        return df_combinado
    else:
        print("Erro: Colunas de matrícula não encontradas em um ou ambos os arquivos.")
        return df_ausencias  # Retorna apenas as ausências se não for possível combinar

def main():
    """
    Função principal para processar arquivos de ausências e funcionários.
    """
    # Defina os caminhos para os arquivos
    arquivo_ausencias = 'AUSENCIAS 0325.xlsx'
    arquivo_funcionarios = 'EQUIPPE  Base Funcionarios.xlsx'
    
    # Verifique se os arquivos existem
    if not os.path.exists(arquivo_ausencias):
        print(f"Erro: Arquivo {arquivo_ausencias} não encontrado.")
        return
    
    if not os.path.exists(arquivo_funcionarios):
        print(f"Erro: Arquivo {arquivo_funcionarios} não encontrado.")
        return
    
    # Processamento para março de 2025
    resultado = processar_dados(arquivo_ausencias, arquivo_funcionarios, mes=3, ano=2025)
    
    if not resultado.empty:
        print(f"Processamento concluído com sucesso. {len(resultado)} registros encontrados.")
        
        # Opcional: salvar o resultado em um novo arquivo Excel
        resultado.to_excel('resultado_ausencias_março_2025.xlsx', index=False)
        print("Arquivo de resultado salvo: resultado_ausencias_março_2025.xlsx")
    else:
        print("Nenhum resultado encontrado ou ocorreu um erro durante o processamento.")

if __name__ == "__main__":
    main()
