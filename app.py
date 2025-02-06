def read_excel(file, sheet_name=0):
    """Lê arquivo Excel com tratamento robusto"""
    try:
        # Primeira tentativa: leitura direta
        df = pd.read_excel(
            file,
            sheet_name=sheet_name,
            engine='openpyxl'
        )
        
        # Se DataFrame está vazio, tenta com header=None
        if df.empty:
            df = pd.read_excel(
                file,
                sheet_name=sheet_name,
                header=None,
                engine='openpyxl'
            )
        
        # Limpa o DataFrame
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        
        # Se tem dados, normaliza as colunas
        if not df.empty:
            # Converte todas as colunas para string
            df.columns = df.columns.astype(str)
            # Remove espaços extras
            df.columns = df.columns.str.strip()
            
            add_to_log(f"Arquivo {file.name} lido com sucesso", 'info')
            return df
            
        add_to_log(f"Arquivo {file.name} está vazio", 'error')
        return None
        
    except Exception as e:
        add_to_log(f"Erro ao ler {file.name}: {str(e)}", 'error')
        return None

def process_faltas(df):
    """Processa faltas com tratamento robusto"""
    try:
        # Lista de possíveis nomes de coluna para faltas
        falta_cols = ['Falta', 'falta', 'Faltas', 'FALTA', 'FALTAS']
        
        # Encontra a coluna correta
        falta_col = None
        for col in df.columns:
            if any(fc.lower() in col.lower() for fc in falta_cols):
                falta_col = col
                break
        
        if falta_col:
            # Converte para string e conta 'x' ou '1'
            df['Falta'] = df[falta_col].fillna('').astype(str).apply(
                lambda x: x.lower().count('x') + x.count('1')
            )
        else:
            df['Falta'] = 0
            
        return df
        
    except Exception as e:
        add_to_log(f"Erro no processamento de faltas: {str(e)}", 'error')
        df['Falta'] = 0
        return df

def process_data(base_file, absence_file, model_file):
    """Processa dados com validações adicionais"""
    try:
        # Lê arquivos
        df_base = read_excel(base_file)
        df_ausencias = read_excel(absence_file)
        df_model = read_excel(model_file)

        if any(df is None for df in [df_base, df_ausencias, df_model]):
            add_to_log("Um ou mais arquivos não foram lidos", 'error')
            return None

        # Verifica colunas obrigatórias
        required_cols = {'Matrícula'}
        for df, name in [(df_base, 'base'), (df_ausencias, 'ausencias')]:
            missing = required_cols - set(df.columns)
            if missing:
                add_to_log(f"Colunas faltando em {name}: {missing}", 'error')
                return None

        # Processa faltas
        df_ausencias = process_faltas(df_ausencias)

        # Merge com tratamento de tipos
        df_merge = pd.merge(
            df_base,
            df_ausencias[['Matrícula', 'Falta', 'Afastamentos', 'Ausência Integral', 'Ausência Parcial']],
            on='Matrícula',
            how='left'
        )

        # Preenche valores ausentes
        df_merge = df_merge.fillna({
            'Falta': 0,
            'Afastamentos': False,
            'Ausência Integral': False,
            'Ausência Parcial': False
        })

        # Processa resultados
        resultados = []
        for _, row in df_merge.iterrows():
            status, valor = calcular_premio(row)
            row_dict = row.to_dict()
            row_dict.update({
                'Status Prêmio': status,
                'Valor Prêmio': valor
            })
            resultados.append(row_dict)

        # Cria DataFrame final
        df_resultado = pd.DataFrame(resultados)

        # Ajusta ao modelo
        for col in df_model.columns:
            if col not in df_resultado.columns:
                df_resultado[col] = None

        return df_resultado[df_model.columns]

    except Exception as e:
        add_to_log(f"Erro no processamento: {str(e)}", 'error')
        return None
