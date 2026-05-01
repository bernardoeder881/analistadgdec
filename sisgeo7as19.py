import streamlit as st
import os
import time
import glob
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# VERSÃO DO CÓDIGO: v6.0
# Analista de Dados: Python, Sheets & Looker Expert
# ==========================================
VERSAO = "v6.0"

st.set_page_config(page_title=f"SisGeO Extrator {VERSAO} 🚒", page_icon="🚒")

# Estilo para o cronômetro
st.markdown("""
    <style>
    .cronometro {
        font-size: 20px;
        font-weight: bold;
        color: #FF4B4B;
        padding: 10px;
        border: 1px solid #FF4B4B;
        border-radius: 5px;
        text-align: center;
    }
    </style>
    """, unsafe_allow_html=True)

st.title(f"SisGeO Extrator {VERSAO} 🚒")
st.sidebar.markdown(f"**Versão Atual:** {VERSAO}")
st.sidebar.info("Especialista: Python | Google Sheets | Looker")

def tratar_e_clonar_manual(caminho_arquivo, data_ini, data_f):
    try:
        df_dados = pd.read_excel(caminho_arquivo, skiprows=2)
        colunas_data = ['Data Ocorrência', 'Data Despacho', 'Data Deslocamento', 'Data Chegada', 'Data Fechamento']
        for col in colunas_data:
            if col in df_dados.columns:
                df_dados[col] = pd.to_datetime(df_dados[col], dayfirst=True, errors='coerce')
                df_dados[col] = df_dados[col] - pd.Timedelta(hours=3)
                df_dados[col] = df_dados[col].dt.strftime('%d/%m/%Y %H:%M')

        with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
            df_dados.to_excel(writer, index=False, startrow=2, sheet_name='Sheet1')
            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']
            fmt_texto = workbook.add_format({'bold': False, 'align': 'left'})
            worksheet.write(0, 0, "SisGeO - Consulta Ocorrência", fmt_texto)
            texto_periodo = f"Período: {data_ini}:00 a {data_f}:59"
            worksheet.write(1, 0, texto_periodo, fmt_texto)
            for i, col in enumerate(df_dados.columns):
                largura = max(df_dados[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, largura)
        return True
    except Exception as e:
        st.error(f"Erro no processamento {VERSAO}: {e}")
        return False

def executar_extracao(tipo_turno):
    # Limpeza inicial
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    inicio_processo = time.time()
    placeholder_tempo = st.empty() # Espaço para o cronômetro
    
    with st.spinner(f"Iniciando extração {tipo_turno}..."):
        try:
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
            chrome_options.binary_location = "/usr/bin/chromium"
            
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 50) 

            # Datas
            hoje_dt = datetime.now()
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            if tipo_turno == "DIA":
                data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                ontem_str = (hoje_dt - timedelta(days=1)).strftime("%d/%m/%Y")
                data_ini, data_f = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # Login e Navegação
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            
            # Enquanto o Selenium trabalha, atualizamos o tempo
            tempo_decorrido = round(time.time() - inicio_processo, 1)
            placeholder_tempo.markdown(f'<div class="cronometro">⏱️ Tempo: {tempo_decorrido}s</div>', unsafe_allow_html=True)
            
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]"))
            
            time.sleep(5)
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            # Filtros
            tempo_decorrido = round(time.time() - inicio_processo, 1)
            placeholder_tempo.markdown(f'<div class="cronometro">⏱️ Tempo: {tempo_decorrido}s</div>', unsafe_allow_html=True)
            
            wait.until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
            driver.execute_script(f"document.getElementById('txtDataInicio').value = '{data_ini}';")
            driver.execute_script(f"document.getElementById('txtDataFim').value = '{data_f}';")
            
            chk = driver.find_element(By.ID, "chkComEmpenho")
            if not chk.is_selected(): driver.execute_script("arguments[0].click();", chk)

            driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btnBuscar"))
            
            # Espera longa do SisGeO (com atualização do cronômetro)
            for i in range(25):
                time.sleep(1)
                tempo_decorrido = round(time.time() - inicio_processo, 1)
                placeholder_tempo.markdown(f'<div class="cronometro">⏱️ Tempo: {tempo_decorrido}s</div>', unsafe_allow_html=True)

            # Download
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)
            
            btn_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", btn_excel)
            
            arquivo_final = None
            for _ in range(30):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)
                tempo_decorrido = round(time.time() - inicio_processo, 1)
                placeholder_tempo.markdown(f'<div class="cronometro">⏱️ Tempo: {tempo_decorrido}s</div>', unsafe_allow_html=True)

            if arquivo_final:
                if tratar_e_clonar_manual(arquivo_final, data_ini, data_f):
                    tempo_total = round(time.time() - inicio_processo, 2)
                    placeholder_tempo.success(f"✅ Finalizado em {tempo_total} segundos!")
                    with open(arquivo_final, "rb") as f:
                        st.download_button(
                            label=f"💾 BAIXAR EXCEL {tipo_turno}",
                            data=f,
                            file_name=f"SisGeO_Consulta_{tipo_turno}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            else:
                st.error("Arquivo não encontrado.")

        except Exception as e:
            st.error(f"Erro na {VERSAO}: {e}")
        finally:
            if 'driver' in locals(): driver.quit()

# Botões
col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ Turno Dia"): executar_extracao("DIA")
with col2:
    if st.button("🌙 Turno Noite"): executar_extracao("NOITE")
