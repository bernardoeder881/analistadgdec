import streamlit as st
import os, time, glob
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# VERSÃO DO CÓDIGO: v6.9 (ESPELHO DA MANUAL)
# Foco: Fidelidade total à planilha base.
# ==========================================
VERSAO = "v6.9"

st.set_page_config(page_title=f"SisGeO Extrator {VERSAO}", page_icon="🚒")

def tratar_e_filtrar_fiel_a_base(caminho, d_ini_str, d_fim_str):
    try:
        # 1. Localiza o início dos dados (Preservando a estrutura SisGeO)
        df_temp = pd.read_excel(caminho)
        linha_cabecalho = 0
        for i, row in df_temp.iterrows():
            if "Data Ocorrência" in str(row.values) or "Protocolo" in str(row.values):
                linha_cabecalho = i + 1
                break
        
        # 2. Carrega mantendo NOMES DE COLUNAS ORIGINAIS
        df = pd.read_excel(caminho, skiprows=linha_cabecalho)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')] # Remove colunas vazias
        
        # ==========================================================
        # NATUREZAS DEFINIDAS PELO ANALISTA (BASE CORRETA)
        # ==========================================================
        naturezas_corretas = [
            'Corte de Árvore', 
            'Salvamento de Pessoa', 
            'DESLIZAMENTO / DESABAMENTO', 
            'ENCHENTE / INUNDAÇÃO', 
            'FOGO EM VEGETAÇÃO'
        ]
        
        # Localiza a coluna correta (pode variar entre 'Tipo Ocorrência' ou 'Natureza')
        col_alvo = next((c for c in df.columns if c in ['Tipo Ocorrência', 'Natureza']), None)
        
        if col_alvo:
            # FILTRO RADICAL: Mantém apenas o que está na lista
            df = df[df[col_alvo].isin(naturezas_corretas)]
        else:
            # Fallback caso o sistema mude o nome da coluna de novo
            for col in df.columns:
                if df[col].astype(str).isin(naturezas_corretas).any():
                    df = df[df[col].isin(naturezas_corretas)]
                    break

        # 3. Tratamento de Datas e Fuso (-3h) - Sem mudar nomes das colunas
        cols_data = ['Data Ocorrência', 'Data Despacho', 'Data Deslocamento', 'Data Chegada', 'Data Fechamento']
        
        if 'Data Ocorrência' in df.columns:
            df['Data Ocorrência'] = pd.to_datetime(df['Data Ocorrência'], dayfirst=True, errors='coerce') - pd.Timedelta(hours=3)
            # Filtro de data rigoroso conforme o turno solicitado
            limite_ini = pd.to_datetime(d_ini_str, dayfirst=True)
            limite_fim = pd.to_datetime(d_fim_str, dayfirst=True)
            df = df[(df['Data Ocorrência'] >= limite_ini) & (df['Data Ocorrência'] <= limite_fim)]

        for col in cols_data:
            if col in df.columns and col != 'Data Ocorrência':
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce') - pd.Timedelta(hours=3)
            if col in df.columns:
                df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M')

        # 4. Exportação idêntica à planilha manual (Start na linha 3)
        with pd.ExcelWriter(caminho, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, startrow=2, sheet_name='Sheet1')
            ws = writer.sheets['Sheet1']
            # Mantém o cabeçalho informativo do SisGeO
            ws.write(0, 0, "SisGeO - Consulta Ocorrência (Filtro Especializado)")
            ws.write(1, 0, f"Período: {d_ini_str} a {d_fim_str}")
            
        return True
    except Exception as e:
        st.error(f"Erro na limpeza v6.9: {e}")
        return False

# Automação de extração
def iniciar_automacao(turno):
    for f in glob.glob("*.xlsx"): os.remove(f)
    inicio_t = time.time()
    container = st.empty()
    
    try:
        opts = Options()
        opts.add_argument('--headless')
        opts.add_argument('--no-sandbox')
        opts.binary_location = "/usr/bin/chromium"
        driver = webdriver.Chrome(options=opts)
        
        # Períodos
        agora = datetime.now()
        hoje = agora.strftime("%d/%m/%Y")
        if turno == "DIA":
            d_ini, d_fim = f"{hoje} 07:01", f"{hoje} 19:00"
        else:
            ontem = (agora - timedelta(days=1)).strftime("%d/%m/%Y")
            d_ini, d_fim = f"{ontem} 19:00", f"{hoje} 07:00"

        driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/Entrar")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("40875")
        driver.find_element(By.ID, "Senha").send_keys("Cidadao51@")
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]"))
        
        time.sleep(3)
        driver.get("https://sisgeo.cbmerj.rj.gov.br/Sisgeo/ConsultaOcorrencia")
        
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "txtDataInicio")))
        driver.execute_script(f"document.getElementById('txtDataInicio').value = '{d_ini}';")
        driver.execute_script(f"document.getElementById('txtDataFim').value = '{d_fim}';")
        driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btnBuscar"))
        
        time.sleep(10) # Aguarda processamento do SisGeO

        # Configura download
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        driver.execute("send_command", {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}})
        
        btn = driver.find_element(By.CSS_SELECTOR, "button.buttons-excel")
        driver.execute_script("arguments[0].click();", btn)
        
        time.sleep(8)
        arq = glob.glob("*.xlsx")[0]
        
        if tratar_e_filtrar_fiel_a_base(arq, d_ini, d_fim):
            container.success(f"✅ Arquivo gerado seguindo a base correta!")
            with open(arq, "rb") as f:
                st.download_button(f"📥 Baixar Base para Looker ({turno})", f, file_name=f"Sisgeo_{turno}_Final.xlsx")
                
    except Exception as e:
        st.error(f"Erro: {e}")
    finally:
        if 'driver' in locals(): driver.quit()

st.title(f"Extrator SisGeO {VERSAO}")
c1, c2 = st.columns(2)
with c1:
    if st.button("☀️ Puxar Turno DIA"): iniciar_automacao("DIA")
with c2:
    if st.button("🌙 Puxar Turno NOITE"): iniciar_automacao("NOITE")
