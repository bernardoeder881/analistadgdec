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

def tratar_excel_fuso(caminho_arquivo):
    """
    Limpa a 'bagunça' do SisGeO:
    1. Pula linhas inúteis de cabeçalho.
    2. Converte textos em datas reais.
    3. Ajusta o fuso horário de Brasília (-3h).
    """
    try:
        # Lê o arquivo pulando as 2 primeiras linhas (onde ficam os títulos do SisGeO)
        # O cabeçalho real passa a ser a linha 3
        df = pd.read_excel(caminho_arquivo, skiprows=2)
        
        # Remove colunas que estejam totalmente vazias e linhas em branco
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)

        # Lista de colunas de data que precisam de correção de fuso
        colunas_data = [
            'Data Ocorrência', 'Data Despacho', 'Data Deslocamento', 
            'Data Chegada', 'Data Fechamento'
        ]
        
        for col in colunas_data:
            if col in df.columns:
                # Converte para formato de data (datetime)
                df[col] = pd.to_datetime(df[col], errors='coerce')
                
                # Subtrai 3 horas (Correção do fuso do servidor)
                # Nota: Mantemos como datetime para o Looker reconhecer como data
                df[col] = df[col] - pd.Timedelta(hours=3)
                
                # Formata para padrão brasileiro para visualização, mas mantém tipo data
                df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M:%S')

        # Salva o arquivo limpo e organizado
        df.to_excel(caminho_arquivo, index=False)
        return True
    except Exception as e:
        st.error(f"Erro ao tratar dados: {e}")
        return False

def executar_extracao(tipo_turno):
    # Limpa arquivos antigos na pasta para não baixar o arquivo errado
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    with st.spinner(f"Processando {tipo_turno}..."):
        try:
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
            
            # Caminho do Chromium no ambiente Streamlit Cloud
            chrome_options.binary_location = "/usr/bin/chromium"
            
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 50) 

            # LÓGICA DE DATAS
            hoje_dt = datetime.now()
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            if tipo_turno == "DIA":
                data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                ontem_str = (hoje_dt - timedelta(days=1)).strftime("%d/%m/%Y")
                data_ini, data_f = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # LOGIN
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            
            user_field = wait.until(EC.presence_of_element_located((By.ID, "Usuario")))
            user_field.send_keys("40875")
            
            pass_field = driver.find_element(By.ID, "Senha")
            pass_field.send_keys("Cidadao51@")
            
            btn_entrar = driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]")
            driver.execute_script("arguments[0].click();", btn_entrar)
            
            time.sleep(6)
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            # FILTROS
            wait.until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
            driver.execute_script(f"document.getElementById('txtDataInicio').value = '{data_ini}';")
            driver.execute_script(f"document.getElementById('txtDataFim').value = '{data_f}';")
            
            # MARCAR VIATURAS
            chk = driver.find_element(By.ID, "chkComEmpenho")
            if not chk.is_selected():
                driver.execute_script("arguments[0].click();", chk)

            # BUSCAR
            btn_buscar = driver.find_element(By.ID, "btnBuscar")
            driver.execute_script("arguments[0].click();", btn_buscar)
            
            st.info("Aguardando processamento do SisGeO (20s)...")
            time.sleep(20)

            # CONFIGURAÇÃO DE DOWNLOAD HEADLESS
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)

            # CLIQUE NO BOTÃO EXCEL
            btn_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", btn_excel)
            
            # ESPERA O DOWNLOAD CONCLUIR
            arquivo_final = None
            for _ in range(40):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)

            if arquivo_final:
                # APLICA A LIMPEZA DE ANALISTA DE DADOS
                sucesso = tratar_excel_fuso(arquivo_final)
                
                if sucesso:
                    with open(arquivo_final, "rb") as f:
                        st.download_button(
                            label=f"💾 BAIXAR RELATÓRIO LIMPO ({tipo_turno})",
                            data=f,
                            file_name=f"Relatorio_{tipo_turno}_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    st.success("Planilha organizada e fuso horário corrigido!")
            else:
                st.error("O sistema demorou muito para gerar o arquivo. Tente novamente.")

        except Exception as e:
            st.error(f"Ocorreu um erro na extração: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()

# INTERFACE
col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ Turno Dia"):
        executar_extracao("DIA")
with col2:
    if st.button("🌙 Turno Noite"):
        executar_extracao("NOITE")
