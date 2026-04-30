import streamlit as st
import os
import time
import glob
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Configuração da Página
st.set_page_config(page_title="SisGeO Extrator 🚒", page_icon="🚒")
st.title("SisGeO Extrator 🚒")
st.write("Selecione o turno para extração automática:")

def iniciar_driver():
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    
    # No Streamlit Cloud, o Chromium geralmente fica neste caminho
    chrome_options.binary_location = "/usr/bin/chromium"

    try:
        # Tenta iniciar usando o ChromeDriver padrão do ambiente Linux
        service = Service("/usr/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        # Fallback caso o caminho acima falhe
        try:
            driver = webdriver.Chrome(options=chrome_options)
            return driver
        except Exception as e2:
            st.error(f"Erro ao iniciar Chrome: {e2}")
            return None

def executar_extracao(tipo_turno):
    # --- LIMPEZA DE SEGURANÇA ---
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    driver = iniciar_driver()
    if not driver:
        return

    wait = WebDriverWait(driver, 30)

    with st.spinner(f"Processando Turno {tipo_turno}..."):
        try:
            # 1. LÓGICA DE DATAS
            hoje_dt = datetime.now()
            ontem_dt = hoje_dt - timedelta(days=1)
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            ontem_str = ontem_dt.strftime("%d/%m/%Y")

            if tipo_turno == "DIA":
                data_ini, data_fim = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                data_ini, data_fim = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # 2. LOGIN
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()
            time.sleep(5)

            # 3. FILTROS
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            def preencher_campo(elemento_id, valor):
                campo = wait.until(EC.presence_of_element_located((By.ID, elemento_id)))
                campo.click()
                campo.send_keys(Keys.CONTROL + "a")
                campo.send_keys(Keys.BACKSPACE)
                time.sleep(0.5)
                campo.send_keys(valor)
                campo.send_keys(Keys.TAB)
                time.sleep(0.5)

            preencher_campo("txtDataInicio", data_ini)
            preencher_campo("txtDataFim", data_fim)
            
            # Clique no checkbox
            empenho = driver.find_element(By.XPATH, "//label[@for='chkComEmpenho']")
            driver.execute_script("arguments[0].click();", empenho)

            # 4. SELEÇÃO DE TIPOS
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
            
            # 5. BUSCA E DOWNLOAD
            driver.find_element(By.ID, "btnBuscar").click()
            st.info(f"⏳ Período: {data_ini} até {data_fim}")
            time.sleep(15) 

            # Habilitar download em modo headless
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)

            botao_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", botao_excel)
            
            # 6. CAPTURA DO ARQUIVO
            arquivo_final = None
            for _ in range(25):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)

            if arquivo_final:
                with open(arquivo_final, "rb") as f:
                    st.download_button(
                        label=f"💾 BAIXAR RELATÓRIO - {tipo_turno}",
                        data=f,
                        file_name=f"Relatorio_{tipo_turno}_{hoje_dt.strftime('%d-%m-%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success(f"✅ Extração {tipo_turno} concluída!")
            else:
                st.error("❌ O Excel não foi gerado pelo sistema.")

        except Exception as e:
            st.error(f"❌ Erro na extração: {e}")
        finally:
            driver.quit()

# Interface de botões
col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ DIA (07:01 às 19:00)"):
        executar_extracao("DIA")
with col2:
    if st.button("🌙 NOITE (19:00 Ontem às 07:00 Hoje)"):
        executar_extracao("NOITE")
