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
# VERSÃO DO CÓDIGO: v7.0 (MODO ESPELHO)
# Foco: Zero descartes por horário. Limpeza apenas por Natureza.
# ==========================================
VERSAO = "v7.0"

st.set_page_config(page_title=f"SisGeO Extrator {VERSAO}", page_icon="🚒")

def tratar_e_filtrar_v7(caminho):
    try:
        # 1. Identifica o cabeçalho original
        df_temp = pd.read_excel(caminho)
        linha_cabecalho = 0
        for i, row in df_temp.iterrows():
            if "Data Ocorrência" in str(row.values) or "Protocolo" in str(row.values):
                linha_cabecalho = i + 1
                break
        
        # 2. Carrega mantendo os nomes de colunas da sua planilha perfeita
        df = pd.read_excel(caminho, skiprows=linha_cabecalho)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        # ==========================================================
        # FILTRO DE NATUREZAS (BUSCA FLEXÍVEL)
        # ==========================================================
        naturezas_alvo = [
            'Corte de Árvore', 
            'Salvamento de Pessoa', 
            'DESLIZAMENTO', 
            'ENCHENTE', 
            'FOGO EM VEGETAÇÃO'
        ]
        
        # Localiza a coluna de natureza
        col_tipo = next((c for c in df.columns if c in ['Tipo Ocorrência', 'Natureza']), None)
        
        if col_tipo:
            # Filtra se CONTÉM qualquer uma das palavras acima (evita erro de espaços/barras)
            padrao = '|'.join(naturezas_alvo)
            df = df[df[col_tipo].str.contains(padrao, case=False, na=False)]

        # 3. Ajuste de Fuso (-3h) - SEM FILTRAR HORÁRIO
        cols_data = ['Data Ocorrência', 'Data Despacho', 'Data Deslocamento', 'Data Chegada', 'Data Fechamento']
        
        for col in cols_data:
            if col in df.columns:
                # Converte e subtrai 3h, mas NÃO deleta ninguém
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce') - pd.Timedelta(hours=3)
                df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M')

        # 4. Salva no formato que o Looker aceita
        with pd.ExcelWriter(caminho, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, startrow=2, sheet_name='Sheet1')
            ws = writer.sheets['Sheet1']
            ws.write(0, 0, "SisGeO - Extração Técnica (Base Conferida)")
            
        return True
    except Exception as e:
        st.error(f"Erro no processamento v7.0: {e}")
        return False

def iniciar_automacao(turno):
    for f in glob.glob("*.xlsx"): os.remove(f)
    inicio_t = time.time()
    
    try:
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.binary_location = "/usr/bin/chromium"
        driver = webdriver.Chrome(options=chrome_options)
        
        # Datas para o preenchimento do site
        agora = datetime.now()
        hoje = agora.strftime("%d/%m/%Y")
        if turno == "DIA":
            d_ini, d_fim = f"{hoje} 07:01", f"{hoje} 19:00"
        else:
            ontem = (agora - timedelta(days=1)).strftime("%d/%m/%Y")
            d_ini, d_fim = f"{ontem} 19:00", f"{hoje} 07:00"

        driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
        driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]"))
        
        time.sleep(3)
        driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
        
        # Preenche os campos de busca do SisGeO
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
        driver.execute_script(f"document.getElementById('txtDataInicio').value = '{d_ini}';")
        driver.execute_script(f"document.getElementById('txtDataFim').value = '{d_fim}';")
        driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btnBuscar"))
        
        time.sleep(12) # Tempo para o SisGeO processar a busca

        # Download
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        driver.execute("send_command", {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}})
        
        btn = driver.find_element(By.CSS_SELECTOR, "button.buttons-excel")
        driver.execute_script("arguments[0].click();", btn)
        
        time.sleep(8)
        arq = glob.glob("*.xlsx")[0]
        
        if tratar_e_filtrar_v7(arq):
            st.success(f"✅ Sucesso! Planilha gerada com base na sua conferência.")
            with open(arq, "rb") as f:
                st.download_button(f"📥 Baixar Arquivo Final ({turno})", f, file_name=f"Sisgeo_{turno}_v7.xlsx")

    except Exception as e:
        st.error(f"Erro Crítico: {e}")
    finally:
        if 'driver' in locals(): driver.quit()

st.title(f"Extrator SisGeO {VERSAO} 🚒")
c1, c2 = st.columns(2)
with c1:
    if st.button("☀️ Turno DIA"): iniciar_automacao("DIA")
with c2:
    if st.button("🌙 Turno NOITE"): iniciar_automacao("NOITE")
