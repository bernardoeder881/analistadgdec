import streamlit as st
import os
import time
from datetime import datetime, timedelta
import pytz  # Biblioteca para gerenciar fusos horários
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

# Botão de Execução
if st.button("Gerar Planilha Agora"):
    with st.spinner("O robô está trabalhando... Isso pode levar até 40 segundos."):
        try:
            # 1. CONFIGURAÇÃO DO CHROME PARA STREAMLIT
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 20)

            # 2. LOGIN
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()
            time.sleep(5)

            # 3. ACESSAR CONSULTA
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            time.sleep(2)

            # 4. SELEÇÃO DE TIPOS (Primeiro para evitar reset)
            botao_tipo = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@data-id, 'ddlTipoOcorrencia') or contains(@title, 'Selecione')]")))
            driver.execute_script("arguments[0].click();", botao_tipo)
            time.sleep(1)

            tipos = ["Corte de Árvore", "DESLIZAMENTO/DESABAMENTO", "Fogo em Vegetação", "Inundação/Alagamento", "Salvamento de Pessoa"]
            for t in tipos:
                try:
                    item = driver.find_element(By.XPATH, f"//span[contains(text(), '{t}')]")
                    driver.execute_script("arguments[0].click();", item)
                except: pass

            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            time.sleep(1)

            # --- AJUSTE DE FUSO HORÁRIO AQUI ---
            fuso_br = pytz.timezone('America/Sao_Paulo')
            agora_br = datetime.now(fuso_br)
            hoje_str = agora_br.strftime("%d/%m/%Y")
            
            data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"

            # 5. PREENCHER DATAS (Agora com a hora certa do Brasil)
            campo_ini = wait.until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
            campo_ini.click()
            campo_ini.send_keys(Keys.CONTROL + "a")
            campo_ini.send_keys(Keys.BACKSPACE)
            campo_ini.send_keys(data_ini)

            campo_fim = driver.find_element(By.ID, "txtDataFim")
            campo_fim.click()
            campo_fim.send_keys(Keys.CONTROL + "a")
            campo_fim.send_keys(Keys.BACKSPACE)
            campo_fim.send_keys(data_f)
            
            driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//label[@for='chkComEmpenho']"))

            # 6. CONSULTA E EXCEL
            driver.find_element(By.ID, "btnBuscar").click()
            st.info(f"📅 Data/Hora Brasil: {agora_br.strftime('%d/%m/%Y %H:%M:%S')}")
            st.write(f"🔍 Filtrando: {data_ini} até {data_f}")
            time.sleep(12) 

            # Comando de download
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
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)

            # 7. DOWNLOAD NA WEB
            if arquivo_final:
                with open(arquivo_final, "rb") as f:
                    st.download_button(
                        label="💾 BAIXAR RELATÓRIO EXCEL",
                        data=f,
                        file_name=arquivo_final,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success(f"✅ Relatório gerado com sucesso!")
            else:
                st.error("❌ O arquivo não foi gerado.")

        except Exception as e:
            st.error(f"❌ Erro: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()
