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
            
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 20)

            # 1. LOGIN
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()
            time.sleep(5)

            # 2. ACESSAR CONSULTA
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            time.sleep(2)

            # 3. SELEÇÃO DE TIPOS
            botao_tipo = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@data-id, 'ddlTipoOcorrencia')]")))
            driver.execute_script("arguments[0].click();", botao_tipo)
            time.sleep(1)

            tipos = ["Corte de Árvore", "DESLIZAMENTO/DESABAMENTO", "Fogo em Vegetação", "Inundação/Alagamento", "Salvamento de Pessoa"]
            for t in tipos:
                try:
                    item = driver.find_element(By.XPATH, f"//span[contains(text(), '{t}')]")
                    driver.execute_script("arguments[0].click();", item)
                except: pass

            # Fecha o menu de tipos com ESC e um clique fora (garantia dupla)
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            time.sleep(1)

            # --- AJUSTE DE FUSO E DATAS ---
            fuso_br = pytz.timezone('America/Sao_Paulo')
            agora_br = datetime.now(fuso_br)
            hoje_str = agora_br.strftime("%d/%m/%Y")
            data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"

            # 4. PREENCHER DATAS (Usando JavaScript para não ser "interceptado")
            def preencher_blindado(id_campo, valor):
                elemento = wait.until(EC.presence_of_element_located((By.ID, id_campo)))
                # Força o clique via JS para ignorar menus sobrepostos
                driver.execute_script("arguments[0].click();", elemento)
                elemento.send_keys(Keys.CONTROL + "a")
                elemento.send_keys(Keys.BACKSPACE)
                elemento.send_keys(valor)
                elemento.send_keys(Keys.TAB)

            preencher_blindado("txtDataInicio", data_ini)
            preencher_blindado("txtDataFim", data_f)
            
            # Clique no checkbox também via JavaScript
            chk = driver.find_element(By.XPATH, "//label[@for='chkComEmpenho']")
            driver.execute_script("arguments[0].click();", chk)

            # 5. BUSCAR
            driver.find_element(By.ID, "btnBuscar").click()
            st.info(f"📅 Horário Brasil: {agora_br.strftime('%H:%M:%S')}")
            time.sleep(12) 

            # 6. DOWNLOAD
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
                    st.download_button(label="💾 BAIXAR EXCEL", data=f, file_name=arquivo_final)
                st.success("✅ Relatório gerado!")
            else:
                st.error("❌ Arquivo não encontrado.")

        except Exception as e:
            st.error(f"❌ Erro: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()
