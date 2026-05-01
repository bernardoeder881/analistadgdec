def tratar_e_filtrar_preservando_nomes(caminho, d_ini_str, d_fim_str):
    """
    v6.6: Filtra Turno + Filtra Tipos de Eventos Específicos
    """
    try:
        # 1. Identifica cabeçalho original
        df_temp = pd.read_excel(caminho)
        linha_cabecalho = 0
        for i, row in df_temp.iterrows():
            if "Data Ocorrência" in str(row.values) or "Protocolo" in str(row.values):
                linha_cabecalho = i + 1
                break
        
        # 2. Carrega mantendo nomes originais
        df = pd.read_excel(caminho, skiprows=linha_cabecalho)
        
        # 3. LISTA DE DESEJADOS (Nomes exatos do SisGeO)
        naturezas_desejadas = [
            'Corte de Árvore', 
            'Salvamento de Pessoa', 
            'DESLIZAMENTO / DESABAMENTO', 
            'ENCHENTE / INUNDAÇÃO', 
            'FOGO EM VEGETAÇÃO'
        ]

        # 4. FILTRAGEM POR NATUREZA (Busca na coluna 6, geralmente chamada de 'Natureza' ou similar)
        # Identificamos a coluna pelo conteúdo se o nome variar
        col_natureza = None
        for col in df.columns:
            if df[col].astype(str).str.contains('Corte de Árvore|Fogo em Vegetação', case=False, na=False).any():
                col_natureza = col
                break
        
        if col_natureza:
            # Mantém apenas as linhas que estão na nossa lista
            df = df[df[col_natureza].isin(naturezas_desejadas)].copy()
        
        # 5. FILTRAGEM POR TURNO (Fuso -3h)
        limite_ini = pd.to_datetime(d_ini_str, dayfirst=True)
        limite_fim = pd.to_datetime(d_f_str if 'd_f_str' in locals() else d_fim_str, dayfirst=True)

        if 'Data Ocorrência' in df.columns:
            df['Data Ocorrência'] = pd.to_datetime(df['Data Ocorrência'], dayfirst=True, errors='coerce')
            df['Data Ocorrência'] = df['Data Ocorrência'] - pd.Timedelta(hours=3)
            df = df[(df['Data Ocorrência'] >= limite_ini) & (df['Data Ocorrência'] <= limite_fim)]

        # 6. Formatação Final e Gravação
        cols_data = ['Data Ocorrência', 'Data Despacho', 'Data Deslocamento', 'Data Chegada', 'Data Fechamento']
        for col in cols_data:
            if col in df.columns:
                if col != 'Data Ocorrência':
                    df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce') - pd.Timedelta(hours=3)
                df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M')

        with pd.ExcelWriter(caminho, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, startrow=2, sheet_name='Sheet1')
            ws = writer.sheets['Sheet1']
            ws.write(0, 0, "SisGeO - Filtro Operacional (v6.6)")
            ws.write(1, 0, f"Período: {d_ini_str} a {d_fim_str}")
            
        return True
    except Exception as e:
        st.error(f"Erro no filtro de natureza v6.6: {e}")
        return False
