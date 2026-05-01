import streamlit as st
import os
import time
import glob
from datetime import datetime
import pytz
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Configuração da Página
st.set_page_config(page_title="SisGeO Extrator 🚒", page_icon="🚒")
st.title("🚒 SisGeO Extrator")
st.write("Relatório Automático: **07:01 às 19:00**")

# Botão de Execução
if st.button("Gerar Planilha Agora"):
    # Limpa arquivos antigos para não baixar o errado
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    with st.spinner("O robô está trabalhando... Isso leva cerca de 40 segundos."):
        try:
            # 1. CONFIGURAÇÃO DO CHROME (Sua base estável)
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            
            # Habilita download no servidor
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 25)

            # 2. LOGIN
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()
            time.sleep(5)

            # 3. ACESSAR CONSULTA
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            time.sleep(2)

            # 4. SELEÇÃO DE TIPOS (Primeiro para o site não resetar as datas depois)
            botao_tipo = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@data-id, 'ddlTipoOcorrencia')]")))
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

            # --- AJUSTE DE FUSO HORÁRIO BRASIL ---
            fuso_br = pytz.timezone('America/Sao_Paulo')
            hoje_br = datetime.now(fuso_br).strftime("%d/%m/%Y")
            data_ini, data_f = f"{hoje_br} 07:01", f"{hoje_br} 19:00"

            # 5. PREENCHER DATAS (Por último para travar o filtro)
            # Viaturas Empenhadas
            driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "chkComEmpenho"))

            for campo_id, valor in [("txtDataInicio", data_ini), ("txtDataFim", data_f)]:
                campo = driver.find_element(By.ID, campo_id)
                campo.click()
                campo.send_keys(Keys.CONTROL + "a")
                campo.send_keys(Keys.BACKSPACE)
                campo.send_keys(valor)
                campo.send_keys(Keys.TAB)
                time.sleep(0.5)

            # 6. BUSCA E DOWNLOAD
            driver.find_element(By.ID, "btnBuscar").click()
            st.write(f"🔍 Filtrando: {data_ini} até {data_f}")
            time.sleep(15) 

            # Autorizar download em headless
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)

            # Clica no botão Amarelo do Excel
            botao_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", botao_excel)
            
            # Aguarda o arquivo aparecer
            arquivo_final = None
            for _ in range(20):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)

            # 7. DOWNLOAD
            if arquivo_final:
                with open(arquivo_final, "rb") as f:
                    st.download_button(
                        label="💾 BAIXAR RELATÓRIO EXCEL",
                        data=f,
                        file_name=f"Relatorio_{hoje_br.replace('/','-')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success(f"✅ Relatório pronto!")
            else:
                st.error("❌ O arquivo não foi gerado.")

        except Exception as e:
            st.error(f"❌ Erro: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()
