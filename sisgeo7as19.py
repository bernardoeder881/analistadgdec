import streamlit as st
import os
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Configuração da Página
st.set_page_config(page_title="SisGeO Extrator Turnos", page_icon="🚒")
st.title("🚒 SisGeO - Extrator de Relatórios")
st.write("Selecione o turno para extração:")

def executar_extracao(tipo_turno):
    with st.spinner(f"Processando Turno {tipo_turno}..."):
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
            wait = WebDriverWait(driver, 25)

            # 2. DEFINIÇÃO DAS DATAS (Lógica de Ontem/Hoje)
            hoje_dt = datetime.now()
            ontem_dt = hoje_dt - timedelta(days=1)
            
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            ontem_str = ontem_dt.strftime("%d/%m/%Y")

            if tipo_turno == "DIA":
                data_ini, data_fim = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:  # Turno NOITE (19h de ontem às 07h de hoje)
                data_ini, data_fim = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # 3. LOGIN
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()
            time.sleep(5)

            # 4. FILTROS
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            input_ini = wait.until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
            input_ini.clear()
            input_ini.send_keys(data_ini)
            
            input_fim = driver.find_element(By.ID, "txtDataFim")
            input_fim.clear()
            input_fim.send_keys(data_fim)
            
            driver.find_element(By.XPATH, "//label[@for='chkComEmpenho']").click()

            # 5. SELEÇÃO DE TIPOS
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
            
            # 6. BUSCA E DOWNLOAD
            driver.find_element(By.ID, "btnBuscar").click()
            st.write(f"⏳ Buscando dados de {data_ini} até {data_fim}...")
            time.sleep(15) 

            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)

            botao_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", botao_excel)
            
            # Verificação do arquivo
            arquivo_final = None
            for _ in range(20):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)

            if arquivo_final:
                with open(arquivo_final, "rb") as f:
                    st.download_button(
                        label=f"💾 BAIXAR RELATÓRIO {tipo_turno}",
                        data=f,
                        file_name=f"Relatorio_{tipo_turno}_{hoje_str.replace('/','-')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success(f"✅ Extração concluída: {data_ini} a {data_fim}")
            else:
                st.error("❌ O SisGeO não gerou o arquivo. Pode não haver ocorrências nesse período.")

        except Exception as e:
            st.error(f"❌ Erro: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()

# Layout dos botões
col1, col2 = st.columns(2)

with col1:
    if st.button("☀️ DIA (07h às 19h Hoje)"):
        executar_extracao("DIA")

with col2:
    if st.button("🌙 NOITE (19h Ontem às 07h Hoje)"):
        executar_extracao("NOITE")
