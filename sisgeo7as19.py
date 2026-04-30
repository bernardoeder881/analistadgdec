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

def iniciar_driver():
    chrome_options = Options()
    chrome_options.add_argument('--headless=new') # Versão nova do headless para maior estabilidade
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--remote-debugging-port=9222')
    chrome_options.binary_location = "/usr/bin/chromium"
    
    # Forçar o diretório de download para a pasta atual
    prefs = {
        "download.default_directory": os.getcwd(),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    try:
        # No Streamlit Cloud, o ChromeDriver já está no PATH se o packages.txt estiver correto
        driver = webdriver.Chrome(options=chrome_options)
        return driver
    except Exception as e:
        st.error(f"Erro ao inicializar o Chrome: {e}")
        return None

def executar_extracao(tipo_turno):
    # Limpa arquivos Excel antigos para não baixar o errado
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    driver = iniciar_driver()
    if not driver:
        return

    wait = WebDriverWait(driver, 30) # Aumentado para 30s por segurança

    with st.spinner(f"Processando Turno {tipo_turno}..."):
        try:
            # 1. Lógica de Datas
            hoje_dt = datetime.now()
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            
            if tipo_turno == "DIA":
                data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                ontem_str = (hoje_dt - timedelta(days=1)).strftime("%d/%m/%Y")
                data_ini, data_f = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # 2. Login
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()
            time.sleep(5)

            # 3. Consulta (Página de Filtros)
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            time.sleep(3)
            
            # 4. Seleção de Tipos (Primeiro para evitar refresh indesejado)
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

            # 5. Viaturas Empenhadas e Datas (O segredo para não vazar horário)
            driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "chkComEmpenho"))

            for campo_id, valor in [("txtDataInicio", data_ini), ("txtDataFim", data_f)]:
                campo = driver.find_element(By.ID, campo_id)
                campo.click()
                campo.send_keys(Keys.CONTROL + "a")
                campo.send_keys(Keys.BACKSPACE)
                time.sleep(0.5)
                campo.send_keys(valor)
                campo.send_keys(Keys.TAB)

            # 6. Busca
            st.info(f"Buscando período: {data_ini} até {data_f}")
            driver.find_element(By.ID, "btnBuscar").click()
            time.sleep(15) # Tempo para a tabela carregar no DOM

            # 7. Download Headless
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)

            botao_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", botao_excel)
            
            # 8. Capturar Arquivo
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
                        label=f"💾 BAIXAR PLANILHA {tipo_turno}",
                        data=f,
                        file_name=f"Relatorio_Sisgeo_{tipo_turno}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success(f"✅ Extração {tipo_turno} concluída!")
            else:
                st.error("❌ O arquivo não foi gerado a tempo.")

        except Exception as e:
            st.error(f"❌ Erro na execução: {e}")
        finally:
            driver.quit()

# Interface de botões
col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ DIA (07:01 - 19:00)"):
        executar_extracao("DIA")
with col2:
    if st.button("🌙 NOITE (19:00 - 07:00)"):
        executar_extracao("NOITE")
