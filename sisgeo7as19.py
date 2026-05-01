import streamlit as st
import os
import time
from datetime import datetime
import pytz
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Configuração da Página
st.set_page_config(page_title="SisGeO Extrator", page_icon="🚀")
st.title("🚀 SisGeO Extrator")
st.write("Gatilho para relatório automático (07:01 às 19:00)")

if st.button("Gerar Planilha Agora"):
    with st.spinner("O robô está trabalhando..."):
        try:
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            # Adicionado para evitar o crash (Stacktrace)
            chrome_options.add_argument('--remote-debugging-port=9222')
            
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 25)

            # 1. LOGIN
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()
            time.sleep(5)

            # 2. FILTROS E DATAS
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            time.sleep(3)

            # 3. SELEÇÃO DE TIPOS (Igual ao seu original)
            botao_tipo = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@data-id, 'ddlTipoOcorrencia')]")))
            driver.execute_script("arguments[0].click();", botao_tipo)
            time.sleep(1)

            tipos = ["Corte de Árvore", "DESLIZAMENTO/DESABAMENTO", "Fogo em Vegetação", "Inundação/Alagamento", "Salvamento de Pessoa"]
            for t in tipos:
                try:
                    item = driver.find_element(By.XPATH, f"//span[contains(text(), '{t}')]")
                    driver.execute_script("arguments[0].click();", item)
                except: pass

            # FECHAR O MENU (Para não dar o erro de clique interceptado)
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            time.sleep(2) # Espera o menu sumir
            
            # --- AJUSTE DE FUSO E DATAS ---
            fuso_br = pytz.timezone('America/Sao_Paulo')
            agora_br = datetime.now(fuso_br)
            hoje = agora_br.strftime("%d/%m/%Y")
            data_ini, data_f = f"{hoje} 07:01", f"{hoje} 19:00"

            # 4. PREENCHER DATAS (Usando JavaScript para evitar erros de clique)
            def preencher_js(id_campo, valor):
                campo = driver.find_element(By.ID, id_campo)
                driver.execute_script("arguments[0].value = '';", campo) # Limpa via JS
                campo.send_keys(valor)
                campo.send_keys(Keys.TAB)

            preencher_js("txtDataInicio", data_ini)
            preencher_js("txtDataFim", data_f)
            
            # Com viaturas empenhadas
            driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "chkComEmpenho"))

            # 5. CONSULTA E EXCEL
            driver.find_element(By.ID, "btnBuscar").click()
            st.write(f"🔍 Filtrando: {data_ini} até {data_f}")
            time.sleep(12)

            # Download
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)

            botao_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", botao_excel)
            
            # Aguarda arquivo
            arquivo_final = None
            for _ in range(15):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivo_final = sorted(arquivos, key=os.path.getmtime)[-1]
                    break
                time.sleep(1)

            if arquivo_final:
                with open(arquivo_final, "rb") as f:
                    st.download_button(label="💾 BAIXAR RELATÓRIO", data=f, file_name=arquivo_final)
                st.success("✅ Concluído!")
            else:
                st.error("❌ Arquivo não gerado.")

        except Exception as e:
            st.error(f"❌ Ocorreu um erro: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()
