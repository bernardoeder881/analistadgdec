import streamlit as st
import os, time, glob
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# VERSÃO DO CÓDIGO: v7.1 (FIX DE COLUNAS)
# Foco: Mapeamento em 'Tipo'/'Subtipo' e Zero descartes por hora.
# ==========================================
VERSAO = "v7.1"

st.set_page_config(page_title=f"SisGeO Extrator {VERSAO}", page_icon="🚒")

def tratar_e_filtrar_fiel(caminho):
    try:
        # 1. Identifica o cabeçalho real (Data Ocorrência ou Protocolo)
        df_raw = pd.read_excel(caminho)
        linha_cabecalho = 0
        for i, row in df_raw.iterrows():
            if any(term in str(row.values) for term in ["Data Ocorrência", "Protocolo", "Ocorrência"]):
                linha_cabecalho = i + 1
                break
        
        # 2. Carrega mantendo nomes originais das colunas
        df = pd.read_excel(caminho, skiprows=linha_cabecalho)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        # ==========================================================
        # FILTRO DE NATUREZAS (BUSCA NAS NOVAS COLUNAS DO SISGEO)
        # ==========================================================
        naturezas_alvo = [
            'Corte de Árvore', 
            'Salvamento de Pessoa', 
            'DESLIZAMENTO', 
            'ENCHENTE', 
            'FOGO EM VEGETAÇÃO'
        ]
        
        # O SisGeO mudou para 'Tipo' ou 'Subtipo'. Vamos checar todas as prováveis.
        col_filtro = next((c for c in df.columns if c in ['Tipo', 'Subtipo', 'Tipo Ocorrência', 'Natureza']), None)
        
        if col_filtro:
            # Filtra apenas se CONTÉM as palavras-chave (ignora lixo como Colisão/Incêndio)
            padrao = '|'.join(naturezas_alvo)
            df = df[df[col_filtro].str.contains(padrao, case=False, na=False)]
        
        # 3. Ajuste de Fuso (-3h) - SEM FILTRO DE HORÁRIO
        # Processa todas as colunas de data que existirem
        cols_data = ['Data Ocorrência', 'Data Despacho', 'Data Deslocamento', 'Data Chegada', 'Data Fechamento']
        for col in cols_data:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce') - pd.Timedelta(hours=3)
                df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M')

        # 4. Salva preservando a estrutura para o Looker
        with pd.ExcelWriter(caminho, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, startrow=2, sheet_name='Sheet1')
            ws = writer.sheets['Sheet1']
            ws.write(0, 0, "BASE CONSOLIDADA - DEFESA CIVIL (v7.1)")
            
        return True, len(df)
    except Exception as e:
        st.error(f"Erro no processamento v7.1: {e}")
        return False, 0

def iniciar_automacao(turno):
    for f in glob.glob("*.xlsx"): os.remove(f)
    inicio_t = time.time()
    
    try:
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.binary_location = "/usr/bin/chromium"
        driver = webdriver.Chrome(options=chrome_options)
        
        # Datas baseadas no relógio do sistema
        agora = datetime.now()
        hoje = agora.strftime("%d/%m/%Y")
        if turno == "DIA":
            d_ini, d_fim = f"{hoje} 07:01", f"{hoje} 19:00"
        else:
            ontem = (agora - timedelta(days=1)).strftime("%d/%m/%Y")
            d_ini, d_fim = f"{ontem} 19:00", f"{hoje} 07:00"

        # Login e Navegação
        driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
        driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]"))
        
        time.sleep(3)
        driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
        
        # Busca no SisGeO
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
        driver.execute_script(f"document.getElementById('txtDataInicio').value = '{d_ini}';")
        driver.execute_script(f"document.getElementById('txtDataFim').value = '{d_fim}';")
        driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btnBuscar"))
        
        time.sleep(12) 

        # Comando de Download
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        driver.execute("send_command", {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}})
        
        btn = driver.find_element(By.CSS_SELECTOR, "button.buttons-excel")
        driver.execute_script("arguments[0].click();", btn)
        
        time.sleep(10)
        arquivos = glob.glob("*.xlsx")
        
        if arquivos:
            arq = arquivos[0]
            sucesso, qtd = tratar_e_filtrar_fiel(arq)
            if sucesso:
                st.success(f"✅ Filtro v7.1 aplicado! {qtd} ocorrências encontradas.")
                with open(arq, "rb") as f:
                    st.download_button(f"📥 Baixar Base Looker ({turno})", f, file_name=f"Sisgeo_{turno}_v7.1.xlsx")
        else:
            st.error("O SisGeO não entregou o arquivo Excel a tempo.")

    except Exception as e:
        st.error(f"Falha na automação: {e}")
    finally:
        if 'driver' in locals(): driver.quit()

# Front-end
st.title(f"SisGeO Extrator {VERSAO} 🚒")
st.markdown("Filtro automático para: *Árvores, Salvamentos, Deslizamentos, Enchentes e Vegetação.*")

col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ Gerar Turno DIA"): iniciar_automacao("DIA")
with col2:
    if st.button("🌙 Gerar Turno NOITE"): iniciar_automacao("NOITE")
