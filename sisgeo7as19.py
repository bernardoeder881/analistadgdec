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
st.set_page_config(page_title="SisGeO Extrator 2.0 🚒", page_icon="🚒")
st.title("SisGeO Extrator 🚒")
st.markdown("---")

def tratar_excel_fuso(caminho_arquivo):
    """Lógica de Analista: Se o site não muda o fuso, o Python muda no arquivo."""
    try:
        df = pd.read_excel(caminho_arquivo)
        for col in df.columns:
            # Se for coluna de data, subtrai 3 horas
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col] - pd.Timedelta(hours=3)
        df.to_excel(caminho_arquivo, index=False)
        return True
    except:
        return False

def executar_extracao(tipo_turno):
    # Limpa arquivos .xlsx da pasta para evitar pegar arquivo velho
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    with st.spinner(f"Iniciando Robô para o Turno: {tipo_turno}..."):
        try:
            # 1. SETUP DO NAVEGADOR
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.binary_location = "/usr/bin/chromium"
            
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 35) # Aumentado para 35 segundos

            # 2. DEFINIÇÃO DE DATAS
            hoje_dt = datetime.now()
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            
            if tipo_turno == "DIA":
                data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                ontem_str = (hoje_dt - timedelta(days=1)).strftime("%d/%m/%Y")
                data_ini, data_f = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # 3. PROCESSO DE LOGIN
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            
            btn_entrar = driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]")
            driver.execute_script("arguments[0].click();", btn_entrar)
            
            # 4. NAVEGAÇÃO E FILTROS
            time.sleep(4)
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            # Preenchimento via JS para não falhar
            wait.until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
            driver.execute_script(f"document.getElementById('txtDataInicio').value = '{data_ini}';")
            driver.execute_script(f"document.getElementById('txtDataFim').value = '{data_f}';")
            
            # Viaturas empenhadas
            chk = wait.until(EC.presence_of_element_located((By.ID, "chkComEmpenho")))
            driver.execute_script("arguments[0].click();", chk)

            # Seleção de Tipos
            btn_tipo = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@data-id, 'ddlTipoOcorrencia')]")))
            driver.execute_script("arguments[0].click();", btn_tipo)
            time.sleep(2)

            tipos = ["Corte de Árvore", "DESLIZAMENTO/DESABAMENTO", "Fogo em Vegetação", "Inundação/Alagamento", "Salvamento de Pessoa"]
            for t in tipos:
                try:
                    item = driver.find_element(By.XPATH, f"//span[contains(text(), '{t}')]")
                    driver.execute_script("arguments[0].click();", item)
                except: pass

            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            
            # 5. BUSCA E DOWNLOAD
            btn_buscar = driver.find_element(By.ID, "btnBuscar")
            driver.execute_script("arguments[0].click();", btn_buscar)
            
            st.info(f"⏳ SisGeO processando dados de {data_ini}...")
            time.sleep(18) # Tempo para o SisGeO carregar a tabela

            # Ativa download no Headless
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)

            btn_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", btn_excel)
            
            # Aguarda o download
            arquivo_final = None
            for i in range(30):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)

            if arquivo_final:
                # TRATAMENTO DE ANALISTA: Corrige as 3 horas de fuso
                tratar_excel_fuso(arquivo_final)
                
                with open(arquivo_final, "rb") as f:
                    st.download_button(
                        label=f"💾 BAIXAR EXCEL CORRIGIDO - {tipo_turno}",
                        data=f,
                        file_name=f"Relatorio_{tipo_turno}_{hoje_str.replace('/','-')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success("✅ Relatório gerado e horários ajustados!")
            else:
                st.error("❌ O site não entregou o arquivo. Tente novamente.")

        except Exception as e:
            # Se der erro, mostra o nome amigável do erro
            st.error(f"⚠️ Erro técnico: {type(e).__name__}")
            st.warning("Dica: Verifique se o login/senha no código estão corretos.")
        finally:
            if 'driver' in locals():
                driver.quit()

# BOTÕES
col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ Turno Dia (07:01 - 19:00)"):
        executar_extracao("DIA")
with col2:
    if st.button("🌙 Turno Noite (19:00 - 07:00)"):
        executar_extracao("NOITE")
