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
    """Lê o Excel e subtrai 3 horas das colunas de data"""
    try:
        df = pd.read_excel(caminho_arquivo)
        # Identifica colunas que possuem data/hora e subtrai 3 horas
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col] - pd.Timedelta(hours=3)
            elif "Data" in col or "Hora" in col:
                # Tenta converter colunas de texto que parecem data
                df[col] = pd.to_datetime(df[col], errors='ignore')
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col] - pd.Timedelta(hours=3)
        
        df.to_excel(caminho_arquivo, index=False)
        return True
    except Exception as e:
        st.warning(f"Aviso no tratamento de dados: {e}")
        return False

def executar_extracao(tipo_turno):
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    with st.spinner(f"Processando turno {tipo_turno}..."):
        try:
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.binary_location = "/usr/bin/chromium"
            
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 30)

            # --- 1. DATAS (Padrão para o Servidor) ---
            hoje_dt = datetime.now()
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            
            if tipo_turno == "DIA":
                data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                ontem_str = (hoje_dt - timedelta(days=1)).strftime("%d/%m/%Y")
                data_ini, data_f = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # --- 2. LOGIN ---
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()
            time.sleep(5)

            # --- 3. FILTROS ---
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            def preencher_campo(id_campo, valor):
                elem = wait.until(EC.element_to_be_clickable((By.ID, id_campo)))
                driver.execute_script(f"document.getElementById('{id_campo}').value = '{valor}';")
                elem.send_keys(Keys.TAB)

            preencher_campo("txtDataInicio", data_ini)
            preencher_campo("txtDataFim", data_f)
            
            # Viaturas empenhadas
            driver.find_element(By.XPATH, "//label[@for='chkComEmpenho']").click()

            # Tipos de Ocorrência
            botao_tipo = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@data-id, 'ddlTipoOcorrencia')]")))
            driver.execute_script("arguments[0].click();", botao_tipo)
            time.sleep(2)

            tipos = ["Corte de Árvore", "DESLIZAMENTO/DESABAMENTO", "Fogo em Vegetação", "Inundação/Alagamento", "Salvamento de Pessoa"]
            for t in tipos:
                try:
                    item = driver.find_element(By.XPATH, f"//span[contains(text(), '{t}')]")
                    driver.execute_script("arguments[0].click();", item)
                except: pass

            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            
            # --- 4. BUSCA E DOWNLOAD ---
            driver.find_element(By.ID, "btnBuscar").click()
            time.sleep(15)

            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)

            botao_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", botao_excel)
            
            arquivo_final = None
            for _ in range(30):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)

            if arquivo_final:
                # --- 5. TRATAMENTO DE DADOS (ANALISTA) ---
                with st.status("Limpando e corrigindo horários..."):
                    tratar_excel_fuso(arquivo_final)
                
                with open(arquivo_final, "rb") as f:
                    st.download_button(
                        label=f"💾 BAIXAR RELATÓRIO CORRIGIDO ({tipo_turno})",
                        data=f,
                        file_name=f"SisGeO_{tipo_turno}_{hoje_str.replace('/','-')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success("Tudo pronto! Horários ajustados para Brasília.")
            else:
                st.error("O arquivo não foi gerado.")

        except Exception as e:
            st.error(f"Erro na automação: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()

# Interface
col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ Turno Dia"): executar_extracao("DIA")
with col2:
    if st.button("🌙 Turno Noite"): executar_extracao("NOITE")
