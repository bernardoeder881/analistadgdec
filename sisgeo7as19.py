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
# VERSÃO DO CÓDIGO: v6.4
# Analista: Estabilidade Total & Filtro Turno
# ==========================================
VERSAO = "v6.4"

st.set_page_config(page_title=f"SisGeO Extrator {VERSAO}", page_icon="🚒", layout="centered")

# CSS para garantir que a UI apareça corretamente
st.markdown("""
    <style>
    .reportview-container { background: #f0f2f6; }
    .cronometro {
        font-size: 22px; font-weight: bold; color: white;
        background-color: #d32f2f; padding: 15px;
        border-radius: 10px; text-align: center; margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title(f"SisGeO Extrator {VERSAO} 🚒")
st.sidebar.markdown(f"### 📊 Dashboard Status\n**Versão:** {VERSAO}\n**Especialista:** Python & BI")

def tratar_e_filtrar_dados(caminho, d_ini, d_fim):
    """ Filtra os dados para garantir que APENAS o turno selecionado seja salvo """
    try:
        # Tenta identificar onde o cabeçalho real começa
        raw_df = pd.read_excel(caminho)
        skip = 0
        for i, row in raw_df.iterrows():
            if "Protocolo" in str(row.values) or "Data Ocorrência" in str(row.values):
                skip = i + 1
                break
        
        df = pd.read_excel(caminho, skiprows=skip)
        
        # Datas para comparação (Fuso já ajustado mentalmente aqui)
        limite_ini = pd.to_datetime(d_ini, dayfirst=True)
        limite_fim = pd.to_datetime(d_fim, dayfirst=True)

        col_data = 'Data Ocorrência'
        if col_data in df.columns:
            df[col_data] = pd.to_datetime(df[col_data], dayfirst=True, errors='coerce')
            # Ajuste de Fuso Horário (-3h)
            df[col_data] = df[col_data] - pd.Timedelta(hours=3)
            
            # FILTRO REAL: Só o que está entre os horários do turno
            df = df[(df[col_data] >= limite_ini) & (df[col_data] <= limite_fim)]

        # Outras colunas de data (apenas fuso)
        outras_cols = ['Data Despacho', 'Data Deslocamento', 'Data Chegada', 'Data Fechamento']
        for c in outras_cols:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], dayfirst=True, errors='coerce') - pd.Timedelta(hours=3)
                df[c] = df[c].dt.strftime('%d/%m/%Y %H:%M')
        
        if col_data in df.columns:
            df[col_data] = df[col_data].dt.strftime('%d/%m/%Y %H:%M')

        # Gravação final com XlsxWriter (Se falhar, grava comum)
        try:
            with pd.ExcelWriter(caminho, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, startrow=2, sheet_name='Sheet1')
                ws = writer.sheets['Sheet1']
                ws.write(0, 0, "SisGeO - Consulta Ocorrência (Filtro v6.4)")
                ws.write(1, 0, f"Período: {d_ini} a {d_fim}")
        except:
            df.to_excel(caminho, index=False)
        return True
    except Exception as e:
        st.error(f"Erro no ETL: {e}")
        return False

def bot_extracao(turno):
    for f in glob.glob("*.xlsx"): os.remove(f) if os.path.exists(f) else None
    
    inicio = time.time()
    barra_progresso = st.empty()
    
    try:
        opts = Options()
        opts.add_argument('--headless')
        opts.add_argument('--no-sandbox')
        opts.binary_location = "/usr/bin/chromium"
        
        driver = webdriver.Chrome(options=opts)
        wait = WebDriverWait(driver, 40)

        # Configuração de Datas
        agora = datetime.now()
        hoje = agora.strftime("%d/%m/%Y")
        if turno == "DIA":
            d_ini, d_fim = f"{hoje} 07:01", f"{hoje} 19:00"
        else:
            ontem = (agora - timedelta(days=1)).strftime("%d/%m/%Y")
            d_ini, d_fim = f"{ontem} 19:00", f"{hoje} 07:00"

        driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
        wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
        driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]"))
        
        time.sleep(4)
        driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
        
        # Preenche datas
        wait.until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
        driver.execute_script(f"document.getElementById('txtDataInicio').value = '{d_ini}';")
        driver.execute_script(f"document.getElementById('txtDataFim').value = '{d_fim}';")
        
        driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btnBuscar"))
        
        # Cronômetro visual
        for i in range(25):
            time.sleep(1)
            barra_progresso.markdown(f'<div class="cronometro">⏱️ Extraindo Dados... {round(time.time()-inicio,1)}s</div>', unsafe_allow_html=True)

        # Download
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        driver.execute("send_command", {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}})
        
        btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel")))
        driver.execute_script("arguments[0].click();", btn)
        
        time.sleep(5)
        arquivos = glob.glob("*.xlsx")
        if arquivos:
            arquivo_ok = arquivos[0]
            if tratar_e_filtrar_dados(arquivo_ok, d_ini, d_fim):
                st.success(f"Concluído em {round(time.time()-inicio,1)}s")
                with open(arquivo_ok, "rb") as f:
                    st.download_button(f"📥 Baixar Planilha {turno}", f, file_name=f"Sisgeo_{turno}.xlsx")
        else:
            st.warning("Nenhum dado encontrado para este período.")
            
    except Exception as e:
        st.error(f"Falha na v6.4: {e}")
    finally:
        driver.quit()

col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ Turno DIA"): bot_extracao("DIA")
with col2:
    if st.button("🌙 Turno NOITE"): bot_extracao("NOITE")
