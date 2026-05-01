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
# VERSÃO DO CÓDIGO: v6.2
# Foco: Correção de erro de tipo (float len)
# ==========================================
VERSAO = "v6.2"

st.set_page_config(page_title=f"SisGeO Extrator {VERSAO}", page_icon="🚒")

# Estilo visual do cronômetro
st.markdown("""
    <style>
    .cronometro {
        font-size: 22px;
        font-weight: bold;
        color: #FFFFFF;
        background-color: #2E7D32; /* Verde para indicar operação */
        padding: 15px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 20px;
        border: 2px solid #1B5E20;
    }
    </style>
    """, unsafe_allow_html=True)

st.title(f"SisGeO Extrator {VERSAO} 🚒")
st.sidebar.write(f"**Analista de Dados:** Operacional")
st.sidebar.info(f"Versão: {VERSAO}")

def tratar_e_clonar_manual(caminho_arquivo, data_ini, data_f):
    try:
        # Lê os dados
        df_dados = pd.read_excel(caminho_arquivo, skiprows=2)
        
        # Correção de fuso e formatação
        colunas_data = ['Data Ocorrência', 'Data Despacho', 'Data Deslocamento', 'Data Chegada', 'Data Fechamento']
        for col in colunas_data:
            if col in df_dados.columns:
                df_dados[col] = pd.to_datetime(df_dados[col], dayfirst=True, errors='coerce')
                df_dados[col] = df_dados[col] - pd.Timedelta(hours=3)
                df_dados[col] = df_dados[col].dt.strftime('%d/%m/%Y %H:%M')

        # Escrita com XlsxWriter (Correção do erro float len aqui)
        try:
            with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
                df_dados.to_excel(writer, index=False, startrow=2, sheet_name='Sheet1')
                workbook  = writer.book
                worksheet = writer.sheets['Sheet1']
                fmt_texto = workbook.add_format({'bold': False, 'align': 'left'})
                
                worksheet.write(0, 0, "SisGeO - Consulta Ocorrência", fmt_texto)
                worksheet.write(1, 0, f"Período: {data_ini}:00 a {data_f}:59", fmt_texto)
                
                # CORREÇÃO v6.2: map(len) agora trata nulos e converte tudo para string primeiro
                for i, col in enumerate(df_dados.columns):
                    # Transforma a coluna em string, substitui nulos por vazio e mede o maior tamanho
                    max_len = df_dados[col].astype(str).replace('nan', '').map(len).max()
                    largura = max(max_len, len(col)) + 2
                    worksheet.set_column(i, i, largura)
        except Exception as e:
            st.warning(f"Aviso: Falha na formatação visual, mas os dados foram salvos. ({e})")
            df_dados.to_excel(caminho_arquivo, index=False)
            
        return True
    except Exception as e:
        st.error(f"Erro no tratamento de dados {VERSAO}: {e}")
        return False

def executar_extracao(tipo_turno):
    # Limpa arquivos temporários
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    inicio_processo = time.time()
    placeholder_tempo = st.empty() 
    
    with st.spinner(f"Processando Turno {tipo_turno}..."):
        try:
            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.binary_location = "/usr/bin/chromium"
            
            prefs = {"download.default_directory": os.getcwd()}
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 50) 

            # Lógica de Horários
            hoje_dt = datetime.now()
            hoje_str = hoje_dt.strftime("%d/%m/%Y")
            if tipo_turno == "DIA":
                data_ini, data_f = f"{hoje_str} 07:01", f"{hoje_str} 19:00"
            else:
                ontem_str = (hoje_dt - timedelta(days=1)).strftime("%d/%m/%Y")
                data_ini, data_f = f"{ontem_str} 19:00", f"{hoje_str} 07:00"

            # Início da navegação
            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
            
            # Login
            wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
            driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
            driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]"))
            
            # Atualiza cronômetro enquanto navega
            for _ in range(3):
                time.sleep(1)
                placeholder_tempo.markdown(f'<div class="cronometro">⏱️ Navegando... {round(time.time() - inicio_processo, 1)}s</div>', unsafe_allow_html=True)

            driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
            
            # Preenchimento de filtros
            wait.until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
            driver.execute_script(f"document.getElementById('txtDataInicio').value = '{data_ini}';")
            driver.execute_script(f"document.getElementById('txtDataFim').value = '{data_f}';")
            
            chk = driver.find_element(By.ID, "chkComEmpenho")
            if not chk.is_selected(): driver.execute_script("arguments[0].click();", chk)

            driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btnBuscar"))
            
            # Espera o processamento do SisGeO (ajustado para 25s)
            for i in range(25):
                time.sleep(1)
                placeholder_tempo.markdown(f'<div class="cronometro">⏱️ SisGeO Processando... {round(time.time() - inicio_processo, 1)}s</div>', unsafe_allow_html=True)

            # Download
            btn_excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.buttons-excel.btn-warning")))
            driver.execute_script("arguments[0].click();", btn_excel)
            
            # Monitora chegada do arquivo
            arquivo_final = None
            for _ in range(20):
                arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx')]
                if arquivos:
                    arquivos.sort(key=os.path.getmtime)
                    arquivo_final = arquivos[-1]
                    break
                time.sleep(1)
                placeholder_tempo.markdown(f'<div class="cronometro">⏱️ Finalizando Download... {round(time.time() - inicio_processo, 1)}s</div>', unsafe_allow_html=True)

            if arquivo_final:
                if tratar_e_clonar_manual(arquivo_final, data_ini, data_f):
                    tempo_total = round(time.time() - inicio_processo, 2)
                    placeholder_tempo.success(f"✅ Concluído com Sucesso! Tempo total: {tempo_total}s")
                    with open(arquivo_final, "rb") as f:
                        st.download_button(
                            label=f"💾 BAIXAR PLANILHA {tipo_turno}",
                            data=f,
                            file_name=f"SisGeO_{tipo_turno}_{hoje_str.replace('/','-')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            else:
                st.error("Arquivo não encontrado. O sistema pode estar lento ou sem dados no período.")

        except Exception as e:
            st.error(f"Erro Crítico na v6.2: {e}")
        finally:
            if 'driver' in locals(): driver.quit()

# Botões de Ação
c1, c2 = st.columns(2)
with c1:
    if st.button("☀️ Turno DIA"): executar_extracao("DIA")
with c2:
    if st.button("🌙 Turno NOITE"): executar_extracao("NOITE")
