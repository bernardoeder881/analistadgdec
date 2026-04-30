import streamlit as st
import os
import time
import glob
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Configuração da Página
st.set_page_config(page_title="SisGeO Extrator", page_icon="🚀")
st.title("🚀 SisGeO Extrator")
st.write("Relatório automático: **07:01 às 19:00**")

# Botão de Execução
if st.button("Gerar Planilha Agora"):
    # Limpa arquivos antigos para evitar confusão no download
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    with st.spinner("O robô está trabalhando... Isso leva cerca de 40 segundos."):
        try:
            # 1. CONFIGURAÇÃO DO CHROME
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

            # 3. FILTROS (Tipos de Ocorrência Primeiro)
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            time.sleep(2)
            
            # Clica para abrir a lista de tipos
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

            # 4. DATAS E VIATURAS (Por último para travar o horário)
            hoje = datetime.now().strftime("%d/%m/%Y")
            data_ini, data_f = f"{hoje} 07:01", f"{hoje} 19:00"

            # Marcar viaturas empenhadas
            driver.find_element(By.XPATH, "//label[@for='chkComEmpenho']").click()

            # Preencher datas com limpeza garantida
            for campo_id, valor in [("txtDataInicio", data_ini), ("txtDataFim", data_f)]:
                campo = driver.find_element(By.ID, campo_id)
                campo.click()
                campo.send_keys(Keys.CONTROL + "a")
                campo.send_keys(Keys.BACKSPACE)
                campo.send_keys(valor)
                campo.send_keys(Keys.TAB)
            
            # 5. CONSULTA E EXCEL
            driver.find_element(By.ID, "btnBuscar").click()
            st.write(f"🔍 Filtrando: {data_ini} até {data_f}")
            time.sleep(12) 

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

            # 6. DOWNLOAD
            if arquivo_final:
                with open(arquivo_final, "rb") as f:
                    st.download_button(
                        label="💾 BAIXAR RELATÓRIO EXCEL",
                        data=f,
                        file_name=f"Relatorio_Sisgeo_{hoje.replace('/','-')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success(f"✅ Relatório de {hoje} pronto!")
            else:
                st.error("❌ O arquivo não foi gerado. Verifique o site.")

        except Exception as e:
            st.error(f"❌ Ocorreu um erro: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()
