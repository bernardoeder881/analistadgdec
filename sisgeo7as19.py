import streamlit as st
import os
import time
import glob
import pandas as pd
from datetime import datetime, timedelta
# ... (imports de selenium permanecem os mesmos)

# ==========================================
# VERSÃO DO CÓDIGO: v6.3
# Foco: Filtragem Rigorosa de Turno e Fix de Cabeçalho
# ==========================================
VERSAO = "v6.3"

def tratar_e_clonar_manual(caminho_arquivo, data_ini_str, data_f_str):
    try:
        # v6.3: Identificação dinâmica do cabeçalho
        # Lemos o arquivo e procuramos onde realmente começam os dados
        df_bruto = pd.read_excel(caminho_arquivo)
        
        # O SisGeO costuma ter o cabeçalho real na linha 2 ou 3
        # Vamos resetar o dataframe a partir da linha que contém "Protocolo" ou "Data Ocorrência"
        for i, row in df_bruto.iterrows():
            if "Protocolo" in str(row.values) or "Data Ocorrência" in str(row.values):
                df_dados = pd.read_excel(caminho_arquivo, skiprows=i+1)
                break
        else:
            df_dados = pd.read_excel(caminho_arquivo, skiprows=2)

        # Converter colunas de data e aplicar fuso -3h
        colunas_data = ['Data Ocorrência', 'Data Despacho', 'Data Deslocamento', 'Data Chegada', 'Data Fechamento']
        
        # Converter limites do turno para comparação real
        limite_inicio = pd.to_datetime(data_ini_str, dayfirst=True)
        limite_fim = pd.to_datetime(data_f_str, dayfirst=True)

        for col in colunas_data:
            if col in df_dados.columns:
                df_dados[col] = pd.to_datetime(df_dados[col], dayfirst=True, errors='coerce')
                # Ajuste de Fuso
                df_dados[col] = df_dados[col] - pd.Timedelta(hours=3)

        # v6.3: FILTRAGEM RIGOROSA (O "Pulo do Gato")
        # Mantemos APENAS o que está rigorosamente dentro do horário do turno
        mascara = (df_dados['Data Ocorrência'] >= limite_inicio) & (df_dados['Data Ocorrência'] <= limite_fim)
        df_dados = df_dados.loc[mascara].copy()

        # Formatar datas de volta para String para o Excel/Sheets
        for col in colunas_data:
            if col in df_dados.columns:
                df_dados[col] = df_dados[col].dt.strftime('%d/%m/%Y %H:%M')

        # Escrita Final (xlsxwriter)
        with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
            df_dados.to_excel(writer, index=False, startrow=2, sheet_name='Sheet1')
            # ... (restante da formatação v6.2 permanece igual)
            
        return True
    except Exception as e:
        st.error(f"Erro no filtro v6.3: {e}")
        return False
