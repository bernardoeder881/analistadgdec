import streamlit as st
import os
import time
import glob
import pytz
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

def executar_extracao(tipo_turno):
    # Limpar arquivos antigos
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
            wait = WebDriverWait(driver, 25)

            # 1. AJUSTE DE FUSO (BRASÍLIA)
            fuso_br = pytz.timezone('America/Sao_Paulo')
            agora_br = datetime.now(fuso_br)
            hoje_str = agora_br.strftime("%d/%m/%Y")
            
            if tipo_turno == "DIA":
                data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                ontem_br = agora_br - timedelta(days=1)
                ontem_str = ontem_br.strftime("%d/%m/%Y")
                data_ini, data_f = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # 2. LOGIN
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()
            time.sleep(5)

            # 3. FILTROS
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            def preencher_estilo_humano(id_campo, valor):
                campo = wait.until(EC.element_to_be_clickable((By.ID, id_campo)))
                campo.click()
                time.sleep(0.5)
                # Limpa usando atalhos de teclado para garantir
                campo.send_keys(Keys.CONTROL + "a")
                campo.send_keys(Keys.BACKSPACE)
                time.sleep(0.5)
                # Digita o valor
                campo.send_keys(valor)
                time.sleep(0.5)
                # Confirma com TAB ou ENTER
                campo.send_keys(Keys.TAB)
                
                # Verificação (log para o Streamlit)
                valor_final = campo.get_attribute("value")
                st.write(f"Campo {id_campo} preenchido com: `{valor_final}`")

            # Executa o preenchimento
            preencher_estilo_humano("txtDataInicio", data_ini)
            preencher_estilo_humano("txtDataFim", data_f)
            
            # Marcar "Com viaturas empenhadas"
            wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@for='chkComEmpenho']"))).click()

            # 4. SELEÇÃO DE TIPOS (MULTISELEÇÃO)
            botao_tipo = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@data-id, 'ddlTipoOcorrencia')]")))
            driver.execute_script("arguments[0].click();", botao_tipo)
            time.sleep(2)

            tipos = ["Corte de Árvore", "DESLIZAMENTO/DESABAMENTO", "Fogo em Vegetação", "Inundação/Alagamento", "Salvamento de Pessoa"]
            for t in tipos:
                try:
                    # Tenta clicar no item da lista
                    item = driver.find_element(By.XPATH, f"//span[contains(text(), '{t}')]")
                    driver.execute_script("arguments[0].click();", item)
                except:
                    pass

            # Fecha o menu de tipos
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            time.sleep(1)
            
            # 5. BUSCAR
            driver.find_element(By.ID, "btnBuscar").click()
            st.info(f"Filtro aplicado no SisGeO: {data_ini} até {data_f}")
            time.sleep(12)

            # 6. DOWNLOAD DO EXCEL
            # Configuração necessária para downloads em modo headless
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)

            botao_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", botao_excel)
            
            # Esperar o arquivo ser criado na pasta
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
                        label=f"💾 BAIXAR EXCEL - {tipo_turno}",
                        data=f,
                        file_name=f"Relatorio_{tipo_turno}_{hoje_str.replace('/','-')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success(f"Relatório de {tipo_turno} gerado com sucesso!")
            else:
                st.error("Não foi possível baixar o arquivo. Verifique se existem dados para esse período.")

        except Exception as e:
            st.error(f"Erro na automação: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()

# Layout de Botões
col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ Turno Dia"):
        executar_extracao("DIA")
with col2:
    if st.button("🌙 Turno Noite"):
        executar_extracao("NOITE")
