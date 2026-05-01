import streamlit as st
import os
import time
import glob
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
st.write("Escolha o turno para gerar o relatório automático:")

def executar_extracao(tipo_turno):
    # Limpeza de arquivos antigos para não baixar o errado
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    with st.spinner(f"O robô está trabalhando no turno {tipo_turno}..."):
        try:
            # 1. CONFIGURAÇÃO DO CHROME PARA STREAMLIT
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.binary_location = "/usr/bin/chromium"
            
            # Habilita download no servidor
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 25)

            # 2. DEFINIÇÃO DAS DATAS
            hoje_dt = datetime.now()
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            
            if tipo_turno == "DIA":
                data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                # Noite: 19:00 de ontem até 07:00 de hoje
                ontem_str = (hoje_dt - timedelta(days=1)).strftime("%d/%m/%Y")
                data_ini, data_f = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # 3. LOGIN
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()
            time.sleep(5)

            # 4. FILTROS E DATAS (COM LIMPEZA TOTAL)
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            def preencher_campo(id_campo, valor):
                campo = wait.until(EC.presence_of_element_located((By.ID, id_campo)))
                campo.click()
                campo.send_keys(Keys.CONTROL + "a")
                campo.send_keys(Keys.BACKSPACE)
                time.sleep(0.5)
                campo.send_keys(valor)
                campo.send_keys(Keys.TAB)

            preencher_campo("txtDataInicio", data_ini)
            preencher_campo("txtDataFim", data_f)
            
            # Com viaturas empenhadas
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
            
            # 6. CONSULTA E EXCEL
            driver.find_element(By.ID, "btnBuscar").click()
            st.write(f"🔍 Buscando período: {data_ini} até {data_f}")
            time.sleep(12)

            # Comando para autorizar download em headless
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)

            # Clica no botão Amarelo do Excel
            botao_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", botao_excel)
            
            # Aguarda o arquivo aparecer na pasta
            arquivo_final = None
            for _ in range(15):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)

            # 7. DISPONIBILIZAR DOWNLOAD
            if arquivo_final:
                with open(arquivo_final, "rb") as f:
                    st.download_button(
                        label=f"💾 BAIXAR EXCEL - {tipo_turno}",
                        data=f,
                        file_name=f"Relatorio_{tipo_turno}_{hoje_str.replace('/','-')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success(f"✅ Relatório {tipo_turno} gerado!")
            else:
                st.error("❌ O arquivo não foi gerado.")

        except Exception as e:
            st.error(f"❌ Erro: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()

# --- INTERFACE DE BOTÕES ---
col1, col2 = st.columns(2)

with col1:
    if st.button("☀️ Turno Dia (07:01 - 19:00)"):
        executar_extracao("DIA")

with col2:
    if st.button("🌙 Turno Noite (19:00 - 07:00)"):
        executar_extracao("NOITE")
