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
# VERSÃO DO CÓDIGO: v5.1
# Analista de Dados: Especialista em Python & Excel
# ==========================================

st.set_page_config(page_title="SisGeO Extrator v5.1 🚒", page_icon="🚒")
st.title("SisGeO Extrator 🚒")

def tratar_e_clonar_manual(caminho_arquivo, data_ini, data_f):
    """
    v5.1: Processa dados, corrige fuso e reconstrói a estética manual do SisGeO.
    """
    try:
        # 1. Lê os dados brutos (pula as 2 linhas de título do SisGeO)
        df_dados = pd.read_excel(caminho_arquivo, skiprows=2)
        
        # 2. Correção de Fuso Horário (-3h)
        colunas_data = ['Data Ocorrência', 'Data Despacho', 'Data Deslocamento', 'Data Chegada', 'Data Fechamento']
        for col in colunas_data:
            if col in df_dados.columns:
                # Converte para datetime garantindo formato brasileiro
                df_dados[col] = pd.to_datetime(df_dados[col], dayfirst=True, errors='coerce')
                # Subtrai 3 horas
                df_dados[col] = df_dados[col] - pd.Timedelta(hours=3)
                # Formata como string para o Excel não bagunçar
                df_dados[col] = df_dados[col].dt.strftime('%d/%m/%Y %H:%M')

        # 3. Reconstrução do arquivo com XlsxWriter
        with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
            # Começa a escrever os dados na linha 3 (índice 2)
            df_dados.to_excel(writer, index=False, startrow=2, sheet_name='Sheet1')
            
            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']

            # Formato de texto simples (igual ao manual)
            fmt_texto = workbook.add_format({'bold': False, 'align': 'left'})
            
            # Escreve os metadados do topo
            worksheet.write(0, 0, "SisGeO - Consulta Ocorrência", fmt_texto)
            texto_periodo = f"Período: {data_ini}:00 a {data_f}:59"
            worksheet.write(1, 0, texto_periodo, fmt_texto)

            # Ajuste de largura das colunas para ficar "apresentável"
            for i, col in enumerate(df_dados.columns):
                largura = max(df_dados[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, largura)

        return True
    except Exception as e:
        st.error(f"Erro no processamento v5.1: {e}")
        return False

def executar_extracao(tipo_turno):
    # Limpa a pasta de arquivos antigos
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    with st.spinner(f"Iniciando Extração v5.1 ({tipo_turno})..."):
        try:
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
            
            # Caminho específico para o Streamlit Cloud
            chrome_options.binary_location = "/usr/bin/chromium"
            
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 50) 

            # Configuração de datas do turno
            hoje_dt = datetime.now()
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            if tipo_turno == "DIA":
                data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                ontem_str = (hoje_dt - timedelta(days=1)).strftime("%d/%m/%Y")
                data_ini, data_f = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # Automação do Site
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]"))
            
            time.sleep(5)
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            # Filtros de data e viaturas
            wait.until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
            driver.execute_script(f"document.getElementById('txtDataInicio').value = '{data_ini}';")
            driver.execute_script(f"document.getElementById('txtDataFim').value = '{data_f}';")
            
            chk = driver.find_element(By.ID, "chkComEmpenho")
            if not chk.is_selected(): driver.execute_script("arguments[0].click();", chk)

            driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btnBuscar"))
            st.info("Aguardando o SisGeO processar os resultados...")
            time.sleep(25) # Tempo de segurança para o SisGeO carregar

            # Download Headless
            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
            params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
            driver.execute("send_command", params)
            
            btn_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", btn_excel)
            
            # Espera o arquivo aparecer na pasta
            arquivo_final = None
            for _ in range(50):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)

            if arquivo_final:
                # Aplica a "mágica" de Analista de Dados para clonar o manual
                if tratar_e_clonar_manual(arquivo_final, data_ini, data_f):
                    with open(arquivo_final, "rb") as f:
                        st.download_button(
                            label=f"💾 BAIXAR RELATÓRIO {tipo_turno} (v5.1)",
                            data=f,
                            file_name=f"SisGeO_Consulta_{tipo_turno}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    st.success("✅ Tudo pronto! Arquivo idêntico ao manual.")
            else:
                st.error("❌ O arquivo não foi gerado. Verifique se há ocorrências no período.")

        except Exception as e:
            st.error(f"Erro na extração: {e}")
        finally:
            if 'driver' in locals(): driver.quit()

# Interface
col1, col2 = st.columns(2)
with col1:
    if st.button("☀️ Turno Dia"): executar_extracao("DIA")
with col2:
    if st.button("🌙 Turno Noite"): executar_extracao("NOITE")
