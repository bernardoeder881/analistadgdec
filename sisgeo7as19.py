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
# VERSÃO DO CÓDIGO: v6.5
# Foco: Preservação de Nomes Originais e Estabilidade UI
# ==========================================
VERSAO = "v6.5"

st.set_page_config(page_title=f"SisGeO Extrator {VERSAO}", page_icon="🚒")

# Interface Visual
st.markdown("""
    <style>
    .cronometro {
        font-size: 22px; font-weight: bold; color: white;
        background-color: #007bff; padding: 15px;
        border-radius: 10px; text-align: center; margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title(f"SisGeO Extrator {VERSAO} 🚒")
st.sidebar.info(f"Analista: Python & BI\nVersão: {VERSAO}")

def tratar_e_filtrar_preservando_nomes(caminho, d_ini_str, d_fim_str):
    """
    v6.5: Filtra o período, corrige fuso, mas MANTÉM NOMES ORIGINAIS.
    """
    try:
        # 1. Identifica onde o cabeçalho original começa
        df_temp = pd.read_excel(caminho)
        linha_cabecalho = 0
        for i, row in df_temp.iterrows():
            if "Data Ocorrência" in str(row.values) or "Protocolo" in str(row.values):
                linha_cabecalho = i + 1
                break
        
        # 2. Carrega com os nomes originais do SisGeO
        df = pd.read_excel(caminho, skiprows=linha_cabecalho)
        
        # 3. Preparação para Filtro Rigoroso
        limite_ini = pd.to_datetime(d_ini_str, dayfirst=True)
        limite_fim = pd.to_datetime(d_fim_str, dayfirst=True)

        # Colunas de data para processamento (nomes literais do SisGeO)
        cols_data = ['Data Ocorrência', 'Data Despacho', 'Data Deslocamento', 'Data Chegada', 'Data Fechamento']
        
        # Filtra baseado na 'Data Ocorrência' original
        if 'Data Ocorrência' in df.columns:
            df['Data Ocorrência'] = pd.to_datetime(df['Data Ocorrência'], dayfirst=True, errors='coerce')
            df['Data Ocorrência'] = df['Data Ocorrência'] - pd.Timedelta(hours=3)
            
            # Aplica o filtro de turno
            df = df[(df['Data Ocorrência'] >= limite_ini) & (df['Data Ocorrência'] <= limite_fim)]

        # 4. Ajusta fuso nas demais colunas sem mudar os nomes
        for col in cols_data:
            if col in df.columns and col != 'Data Ocorrência':
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce') - pd.Timedelta(hours=3)
            
            if col in df.columns:
                df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M')

        # 5. Salva mantendo a estrutura SisGeO para Google Planilhas/Looker
        with pd.ExcelWriter(caminho, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, startrow=2, sheet_name='Sheet1')
            ws = writer.sheets['Sheet1']
            fmt = writer.book.add_format({'align': 'left'})
            ws.write(0, 0, "SisGeO - Consulta Ocorrência", fmt)
            ws.write(1, 0, f"Período: {d_ini_str} a {d_fim_str}", fmt)
            
        return True
    except Exception as e:
        st.error(f"Erro no processamento v6.5: {e}")
        return False

def iniciar_automacao(turno):
    # Limpa vestígios
    for f in glob.glob("*.xlsx"): os.remove(f)
    
    inicio_t = time.time()
    container_tempo = st.empty()
    
    try:
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.binary_location = "/usr/bin/chromium"
        
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 45)

        # Definição de Período
        agora = datetime.now()
        hoje_s = agora.strftime("%d/%m/%Y")
        if turno == "DIA":
            d_ini, d_fim = f"{hoje_s} 07:01", f"{hoje_s} 19:00"
        else:
            ontem_s = (agora - timedelta(days=1)).strftime("%d/%m/%Y")
            d_ini, d_fim = f"{ontem_s} 19:00", f"{hoje_s} 07:00"

        # Login
        driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
        wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
        driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]"))
        
        # Navegação e Filtros
        time.sleep(3)
        driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
        
        wait.until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
        driver.execute_script(f"document.getElementById('txtDataInicio').value = '{d_ini}';")
        driver.execute_script(f"document.getElementById('txtDataFim').value = '{d_f_str if 'd_f_str' in locals() else d_fim}';")
        
        driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btnBuscar"))
        
        # Cronômetro de extração
        for _ in range(25):
            time.sleep(1)
            container_tempo.markdown(f'<div class="cronometro">⏱️ Extraindo do SisGeO: {round(time.time()-inicio_t, 1)}s</div>', unsafe_allow_html=True)

        # Trigger de Download
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        driver.execute("send_command", {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}})
        
        btn_xls = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel")))
        driver.execute_script("arguments[0].click();", btn_xls)
        
        # Captura do Arquivo
        time.sleep(6)
        lista_arquivos = glob.glob("*.xlsx")
        if lista_arquivos:
            arq = lista_arquivos[0]
            if tratar_e_filtrar_preservando_nomes(arq, d_ini, d_fim):
                container_tempo.success(f"✅ Finalizado em {round(time.time()-inicio_t, 1)}s")
                with open(arq, "rb") as f:
                    st.download_button(f"📥 Baixar Arquivo {turno}", f, file_name=f"Sisgeo_{turno}_OriginalNames.xlsx")
        else:
            st.warning("Nenhum arquivo gerado pelo sistema.")

    except Exception as e:
        st.error(f"Erro Crítico v6.5: {e}")
    finally:
        if 'driver' in locals(): driver.quit()

# Botões Front-end
c1, c2 = st.columns(2)
with c1:
    if st.button("☀️ Turno DIA"): iniciar_automacao("DIA")
with c2:
    if st.button("🌙 Turno NOITE"): iniciar_automacao("NOITE")
