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

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(page_title="SisGeO Extrator 🚒", page_icon="🚒")
st.title("SisGeO Extrator 🚒")

def tratar_e_clonar_manual(caminho_arquivo, data_ini, data_f):
    """
    Transforma a extração do robô em uma cópia idêntica à extração manual.
    """
    try:
        # 1. Lê a planilha bruta ignorando o lixo inicial para processar os dados
        df_dados = pd.read_excel(caminho_arquivo, skiprows=2)
        
        # 2. Correção de Fuso Horário (-3h) nas colunas de data
        colunas_data = ['Data Ocorrência', 'Data Despacho', 'Data Deslocamento', 'Data Chegada', 'Data Fechamento']
        for col in colunas_data:
            if col in df_dados.columns:
                df_dados[col] = pd.to_datetime(df_dados[col], dayfirst=True, errors='coerce')
                df_dados[col] = df_dados[col] - pd.Timedelta(hours=3)
                # Formata exatamente como o SisGeO manual (DD/MM/YYYY HH:MM)
                df_dados[col] = df_dados[col].dt.strftime('%d/%m/%Y %H:%M')

        # 3. RECONSTRUÇÃO DA ESTRUTURA MANUAL (O "PULO DO GATO")
        # Criamos as duas linhas de topo que o SisGeO gera manualmente
        linha_1 = ["SisGeO - Consulta Ocorrência"] + [None] * (len(df_dados.columns) - 1)
        linha_2 = [f"Período: {data_ini}:00 a {data_f}:59"] + [None] * (len(df_dados.columns) - 1)
        
        # Criar um novo DataFrame com essas linhas
        df_topo = pd.DataFrame([linha_1, linha_2], columns=df_dados.columns)
        
        # Concatenar: Topo + Nomes das Colunas (como linha) + Dados
        # Primeiro, transformamos os nomes das colunas em uma linha de dados
        df_colunas = pd.DataFrame([df_dados.columns.tolist()], columns=df_dados.columns)
        
        # Juntamos tudo
        df_final = pd.concat([df_topo, df_colunas, df_dados], ignore_index=True)

        # 4. Salva sem o cabeçalho do Pandas (pois já inserimos manualmente)
        df_final.to_excel(caminho_arquivo, index=False, header=False)
        return True
    except Exception as e:
        st.error(f"Erro na formatação: {e}")
        return False

def executar_extracao(tipo_turno):
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    with st.spinner(f"Extraindo {tipo_turno}..."):
        try:
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
            chrome_options.binary_location = "/usr/bin/chromium"
            
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 50) 

            # DATAS
            hoje_dt = datetime.now()
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            if tipo_turno == "DIA":
                data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                ontem_str = (hoje_dt - timedelta(days=1)).strftime("%d/%m/%Y")
                data_ini, data_f = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # LOGIN E NAVEGAÇÃO
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]"))
            
            time.sleep(5)
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            # FILTROS
            wait.until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
            driver.execute_script(f"document.getElementById('txtDataInicio').value = '{data_ini}';")
            driver.execute_script(f"document.getElementById('txtDataFim').value = '{data_f}';")
            
            chk = driver.find_element(By.ID, "chkComEmpenho")
            if not chk.is_selected(): driver.execute_script("arguments[0].click();", chk)

            driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btnBuscar"))
            st.info("Processando no SisGeO...")
            time.sleep(20)

            # DOWNLOAD
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)
            
            btn_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", btn_excel)
            
            arquivo_final = None
            for _ in range(45):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)

            if arquivo_final:
                # FORMATAÇÃO PARA FICAR IGUAL AO MANUAL
                if tratar_e_clonar_manual(arquivo_final, data_ini, data_f):
                    with open(arquivo_final, "rb") as f:
                        st.download_button(
                            label="💾 BAIXAR EXTRATO IDÊNTICO AO MANUAL",
                            data=f,
                            file_name=f"SisGeO_Consulta_{tipo_turno}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    st.success("Pronto! Arquivo idêntico ao extraído manualmente.")
            else:
                st.error("Erro no download.")

        except Exception as e:
            st.error(f"Erro: {e}")
        finally:
            if 'driver' in locals(): driver.quit()

# BOTAO
col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ Turno Dia"): executar_extracao("DIA")
with col2:
    if st.button("🌙 Turno Noite"): executar_extracao("NOITE")
