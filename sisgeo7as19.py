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
from selenium.webdriver.common.keys import Keys

# Configuração da Página
st.set_page_config(page_title="SisGeO Extrator 🚒", page_icon="🚒")
st.title("SisGeO Extrator 🚒")

def tratar_excel_fuso(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo)
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col] - pd.Timedelta(hours=3)
        df.to_excel(caminho_arquivo, index=False)
        return True
    except:
        return False

def executar_extracao(tipo_turno):
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    with st.spinner(f"Processando {tipo_turno}..."):
        try:
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--window-size=1920,1080')
            # Disfarce para o site não bloquear o robô
            chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
            
            chrome_options.binary_location = "/usr/bin/chromium"
            
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            # Aumentamos o Wait para 50 segundos para evitar o TimeoutException
            wait = WebDriverWait(driver, 50) 

            # DATAS
            hoje_dt = datetime.now()
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            if tipo_turno == "DIA":
                data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                ontem_str = (hoje_dt - timedelta(days=1)).strftime("%d/%m/%Y")
                data_ini, data_f = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # LOGIN COM VERIFICAÇÃO ETAPA POR ETAPA
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            
            user_field = wait.until(EC.presence_of_element_located((By.ID, "Usuario")))
            user_field.send_keys("40875")
            
            pass_field = driver.find_element(By.ID, "Senha")
            pass_field.send_keys("Cidadao51@")
            
            btn_entrar = driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]")
            driver.execute_script("arguments[0].click();", btn_entrar)
            
            # ESPERA O LOGIN CONCLUIR
            time.sleep(6)
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            # FILTROS VIA JS (MAIS RÁPIDO E EVITA TIMEOUT)
            wait.until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
            driver.execute_script(f"document.getElementById('txtDataInicio').value = '{data_ini}';")
            driver.execute_script(f"document.getElementById('txtDataFim').value = '{data_f}';")
            
            # MARCAR VIATURAS
            chk = driver.find_element(By.ID, "chkComEmpenho")
            driver.execute_script("arguments[0].click();", chk)

            # BUSCA
            btn_buscar = driver.find_element(By.ID, "btnBuscar")
            driver.execute_script("arguments[0].click();", btn_buscar)
            
            st.info("Aguardando o SisGeO processar a tabela...")
            time.sleep(20) # SisGeO é lento para carregar os resultados

            # DOWNLOAD
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)

            # Localiza o botão de Excel e clica
            btn_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", btn_excel)
            
            # ESPERA O ARQUIVO APARECER
            arquivo_final = None
            for _ in range(40):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)

            if arquivo_final:
                tratar_excel_fuso(arquivo_final)
                with open(arquivo_final, "rb") as f:
                    st.download_button(
                        label=f"💾 BAIXAR EXCEL CORRIGIDO",
                        data=f,
                        file_name=f"Relatorio_{tipo_turno}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success("Sucesso!")
            else:
                st.error("O arquivo não foi gerado a tempo.")

        except Exception as e:
            st.error(f"Erro: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()

col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ Turno Dia"): executar_extracao("DIA")
with col2:
    if st.button("🌙 Turno Noite"): executar_extracao("NOITE")
