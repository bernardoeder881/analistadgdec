import streamlit as st
import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Interface do Streamlit
st.title("🚀 SisGeO Extrator")
st.write("Clique no botão abaixo para gerar o relatório de hoje (07:01 - 19:00)")

if st.button("Gerar Planilha Agora"):
    with st.spinner("O robô está acessando o SisGeO..."):
        try:
            # --- CONFIGURAÇÃO DO CHROME ---
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 20)

            # --- FLUXO DO SISGEO (LOGIN E FILTROS) ---
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            # (Aqui entra todo o seu código de login e seleção de tipos que já funciona)
            
            # --- EXTRAÇÃO ---
            # Após clicar no botão Excel, o arquivo é gerado no servidor
            time.sleep(10)
            
            # Localizar o arquivo mais recente
            arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
            arquivos.sort(key=os.path.getmtime)
            
            if arquivos:
                with open(arquivos[-1], "rb") as f:
                    st.download_button(
                        label="💾 Clique aqui para baixar o Excel",
                        data=f,
                        file_name=arquivos[-1],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success("Relatório pronto!")
            else:
                st.error("Nenhum arquivo gerado.")

        except Exception as e:
            st.error(f"Erro: {e}")
        finally:
            driver.quit()
