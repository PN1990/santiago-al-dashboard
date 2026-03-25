#!/usr/bin/env python3
"""
Bot para descarregar reservas da Ynnov e importar para o Supabase.
Corre automaticamente via GitHub Actions todos os dias às 7h.
"""

import os
import time
import random
import json
import tempfile
from datetime import datetime

# ── Dependências ──────────────────────────────────────────────────────────────
import xlrd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import openpyxl
from supabase import create_client

# ── Configuração (via GitHub Secrets) ─────────────────────────────────────────
YNNOV_EMAIL    = os.environ["YNNOV_EMAIL"]
YNNOV_PASSWORD = os.environ["YNNOV_PASSWORD"]
SUPABASE_URL   = os.environ["SUPABASE_URL"]
SUPABASE_KEY   = os.environ["SUPABASE_KEY"]

# ── Helpers ───────────────────────────────────────────────────────────────────
def esperar(min_s=1.5, max_s=3.5):
    """Pausa aleatória para imitar comportamento humano."""
    time.sleep(random.uniform(min_s, max_s))

def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

# ── Selenium ──────────────────────────────────────────────────────────────────
def criar_driver():
    opts = Options()
    opts.add_argument("--headless")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1280,900")
    opts.add_argument("--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    # Pasta temporária para downloads
    download_dir = tempfile.mkdtemp()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    opts.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=opts)
    return driver, download_dir

def fazer_login(driver):
    log("A abrir página de login da Ynnov...")
    driver.get("https://app.ynnov.pt/login")
    esperar(2, 4)

    wait = WebDriverWait(driver, 15)

    # Preencher email
    log("A preencher email...")
    campo_email = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email'], input[name='email'], input[placeholder*='email' i]")))
    esperar(0.5, 1.5)
    campo_email.clear()
    for char in YNNOV_EMAIL:
        campo_email.send_keys(char)
        time.sleep(random.uniform(0.05, 0.15))
    esperar(0.5, 1.5)

    # Preencher password
    log("A preencher password...")
    campo_pass = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
    campo_pass.clear()
    for char in YNNOV_PASSWORD:
        campo_pass.send_keys(char)
        time.sleep(random.uniform(0.05, 0.15))
    esperar(0.8, 2)

    # Tirar screenshot antes de submeter (para debug)
    driver.save_screenshot("/tmp/ynnov_before_login.png")
    log(f"Screenshot tirado. URL atual: {driver.current_url}")
    log(f"Título da página: {driver.title}")

    # Submeter — tentar vários seletores
    log("A submeter login...")
    btn_login = None
    seletores_btn = [
        "button[type='submit']",
        "input[type='submit']",
        "button.btn-primary",
        "button.login-btn",
        "button.btn-login",
        "button[class*='login']",
        "button[class*='submit']",
        "form button",
        "button",
    ]
    for seletor in seletores_btn:
        try:
            elementos = driver.find_elements(By.CSS_SELECTOR, seletor)
            if elementos:
                btn_login = elementos[-1]  # último botão da página
                log(f"Botão encontrado com: {seletor} — texto: '{btn_login.text}'")
                break
        except:
            continue

    if not btn_login:
        # Tentar submeter com Enter no campo password
        log("Botão não encontrado, a tentar Enter no campo password...")
        from selenium.webdriver.common.keys import Keys
        campo_pass.send_keys(Keys.RETURN)
    else:
        esperar(0.5, 1)
        btn_login.click()

    esperar(4, 6)
    driver.save_screenshot("/tmp/ynnov_after_login.png")
    log(f"URL após login: {driver.current_url}")

    # Verificar login
    if "login" in driver.current_url.lower():
        # Mostrar HTML da página para debug
        log("Login pode ter falhado. HTML da página:")
        log(driver.page_source[:500])
        raise Exception("Login falhou! Ver screenshots de debug.")
    log(f"Login OK — URL: {driver.current_url}")

def clicar_texto(driver, texto, timeout=15):
    """Encontrar e clicar num elemento pelo texto visível."""
    wait = WebDriverWait(driver, timeout)
    from selenium.webdriver.common.by import By
    import selenium.webdriver.support.expected_conditions as EC
    el = wait.until(EC.element_to_be_clickable(
        (By.XPATH, f"//*[normalize-space(text())='{texto}' or @title='{texto}']")
    ))
    esperar(0.3, 0.8)
    el.click()
    return el

def descarregar_excel(driver, download_dir):
    import glob
    wait = WebDriverWait(driver, 20)

    # 1. Clicar em Reservas no menu lateral
    log("A clicar em Reservas...")
    clicar_texto(driver, "Reservas")
    esperar(2, 3)
    driver.save_screenshot("/tmp/ynnov_reservas.png")

    # 2. Clicar em Lista
    log("A clicar em Lista...")
    clicar_texto(driver, "Lista")
    esperar(1, 2)
    driver.save_screenshot("/tmp/ynnov_lista.png")

    # 3. Clicar em Filtros
    log("A clicar em Filtros...")
    clicar_texto(driver, "Filtros")
    esperar(1, 2)
    driver.save_screenshot("/tmp/ynnov_filtros.png")

    # 4. Limpar todos os filtros de estado primeiro (clicar em "Clear" ou "Limpar")
    log("A limpar filtros de estado...")
    try:
        btn_clear = driver.find_element(By.XPATH, "//button[normalize-space(text())='Clear' or normalize-space(text())='Limpar']")
        btn_clear.click()
        esperar(0.5, 1)
    except:
        log("Botão clear não encontrado, a tentar desselecionar todos...")
        # Tentar clicar em "All" para desselecionar tudo
        try:
            btn_all = driver.find_element(By.XPATH, "//button[normalize-space(text())='All']")
            btn_all.click()
            esperar(0.5, 1)
            btn_all.click()  # dois cliques para desselecionar
            esperar(0.5, 1)
        except:
            pass

    # 5. Selecionar apenas Confirmado, Check-in e Check-out
    for estado in ["Confirmado", "Check-in", "Check-out"]:
        log(f"A selecionar estado: {estado}...")
        try:
            els = driver.find_elements(By.XPATH, f"//span[normalize-space(text())='{estado}'] | //li[normalize-space(text())='{estado}'] | //div[normalize-space(text())='{estado}']")
            if els:
                # Usar JavaScript click para evitar elementos sobrepostos
                driver.execute_script("arguments[0].click();", els[0])
                esperar(0.3, 0.8)
                log(f"Estado '{estado}' selecionado.")
            else:
                log(f"Aviso: elemento '{estado}' não encontrado.")
        except Exception as e:
            log(f"Aviso: erro ao selecionar '{estado}': {e}")

    driver.save_screenshot("/tmp/ynnov_filtros_selecionados.png")

    # 6. Clicar em Aplicar
    log("A clicar em Aplicar...")
    clicar_texto(driver, "Aplicar")
    esperar(2, 3)
    driver.save_screenshot("/tmp/ynnov_apos_filtros.png")

    # 7. Clicar no botão XLS para download
    log("A clicar no botão XLS...")
    seletores_xls = [
        "//button[normalize-space(text())='xls' or normalize-space(text())='XLS']",
        "//a[normalize-space(text())='xls' or normalize-space(text())='XLS']",
        "//*[contains(@title,'xls') or contains(@title,'XLS')]",
        "//*[contains(@class,'xls')]",
    ]
    btn_xls = None
    for xpath in seletores_xls:
        try:
            elementos = driver.find_elements(By.XPATH, xpath)
            if elementos:
                btn_xls = elementos[0]
                log(f"Botão XLS encontrado: {xpath}")
                break
        except:
            continue

    if not btn_xls:
        driver.save_screenshot("/tmp/ynnov_debug.png")
        raise Exception("Botão XLS não encontrado. Ver screenshots de debug.")

    esperar(0.5, 1)
    btn_xls.click()
    log("A aguardar download do Excel...")
    esperar(5, 8)

    # 8. Encontrar o ficheiro descarregado
    ficheiros = glob.glob(os.path.join(download_dir, "*.xlsx")) + \
                glob.glob(os.path.join(download_dir, "*.xls"))

    if not ficheiros:
        driver.save_screenshot("/tmp/ynnov_debug.png")
        raise Exception("Ficheiro Excel não encontrado após download.")

    ficheiro = sorted(ficheiros, key=os.path.getmtime)[-1]
    log(f"Excel descarregado: {ficheiro}")
    return ficheiro

# ── Processar Excel ───────────────────────────────────────────────────────────
def parse_data(valor):
    """Converter datas do Excel para string YYYY-MM-DD."""
    import pandas as pd
    if valor is None or (isinstance(valor, float) and str(valor) == 'nan'):
        return None
    # pandas Timestamp
    if hasattr(valor, 'strftime'):
        return valor.strftime("%Y-%m-%d")
    if isinstance(valor, datetime):
        return valor.strftime("%Y-%m-%d")
    s = str(valor).strip()
    if s in ('', 'nan', 'NaT', 'None'):
        return None
    if s.startswith("20") and len(s) >= 10:
        return s[:10]
    return None

def converter_xls_para_csv(ficheiro):
    """Converter XLS para CSV usando LibreOffice (disponível no Ubuntu)."""
    import subprocess, glob, tempfile, shutil
    out_dir = tempfile.mkdtemp()
    log(f"A converter XLS para CSV com LibreOffice...")
    result = subprocess.run([
        'libreoffice', '--headless', '--convert-to', 'csv',
        '--outdir', out_dir, ficheiro
    ], capture_output=True, text=True, timeout=60)
    log(f"LibreOffice stdout: {result.stdout}")
    log(f"LibreOffice stderr: {result.stderr}")
    csvs = glob.glob(os.path.join(out_dir, "*.csv"))
    if not csvs:
        raise Exception("LibreOffice não gerou CSV.")
    return csvs[0]

def processar_excel(ficheiro):
    log(f"A processar Excel: {ficheiro}")
    import pandas as pd

    df = None

    # Tentar ler directamente com pandas
    for engine in ['xlrd', 'openpyxl']:
        try:
            df = pd.read_excel(ficheiro, engine=engine)
            log(f"Excel lido com engine: {engine}")
            break
        except Exception as e:
            log(f"Engine {engine} falhou: {e}")

    # Se falhou, converter com LibreOffice e ler CSV
    if df is None:
        try:
            csv_file = converter_xls_para_csv(ficheiro)
            df = pd.read_csv(csv_file, encoding='utf-8', sep=',')
            log(f"Ficheiro lido via CSV: {csv_file}")
        except Exception as e:
            log(f"Conversão LibreOffice falhou: {e}")

    if df is None:
        raise Exception("Não foi possível ler o ficheiro Excel com nenhum método.")

    headers = [str(c).strip() for c in df.columns]
    log(f"Colunas encontradas: {headers}")

    reservas = []
    for _, row in df.iterrows():
        r = {str(k).strip(): v for k, v in row.items()}

        import pandas as pd
        def val(key, default=""):
            v = r.get(key, default)
            if v is None or (isinstance(v, float) and str(v) == 'nan'):
                return default
            return v

        id_reserva = str(val("ID")).strip().replace(".0", "")
        hospede    = str(val("Hóspede", val("Hospede"))).strip()

        if not id_reserva or id_reserva in ('', 'nan') or not hospede:
            continue

        reservas.append({
            "id":                id_reserva,
            "hospede":           hospede,
            "checkin":           parse_data(val("Data de check-in")),
            "hora_checkin":      str(val("Hora de check-in", "") or ""),
            "checkout":          parse_data(val("Data de check-out")),
            "hora_checkout":     str(val("Hora de check-out", "") or ""),
            "noites":            int(r.get("N.º de noites", val("N de noites", 0)) or 0),
            "adultos":           int(val("Adultos", 0) or 0),
            "criancas":          int(r.get("Crianças", val("Criancas", 0)) or 0),
            "bebes":             int(r.get("Bebés", val("Bebes", 0)) or 0),
            "telefone":          str(val("Telefone", "") or ""),
            "email":             str(r.get("Email (pessoal)", val("Email (canal)", "")) or ""),
            "pais":              str(r.get("País", val("Pais", "")) or ""),
            "codigo_pais":       str(r.get("Código do país", val("Codigo do pais", "")) or ""),
            "alojamento":        str(val("Alojamento", "") or ""),
            "tmt":               float(val("TMT", 0) or 0),
            "total":             float(val("Total da reserva", 0) or 0),
            "estado":            str(val("Estado da reserva", "") or ""),
            "estado_pagamento":  str(val("Estado do pagamento", "") or ""),
            "canal":             str(val("Canal", "") or ""),
            "comissao":          float(r.get("Comissão do canal", val("Comissao do canal", 0)) or 0),
            "comissao_pct":      float(val("Comissão do canal (%)", 0) or 0),
            "id_canal":          str(val("ID da reserva (canal)", "") or ""),
            "data_criacao":      parse_data(r.get("Data de criação", val("Data de criacao"))),
            "antecedencia":      int(r.get("Antecedência da reserva (dias)", val("Antecedencia da reserva (dias)", 0)) or 0),
            "checkin_efetuado":  str(val("Check-in efetuado a...", "") or ""),
            "checkout_efetuado": str(val("Check-out efetuado a...", "") or ""),
            "notas_canal":       str(val("Notas do canal", "") or ""),
            "fatura":            str(val("Fatura(s)", "") or ""),
            "estado_aima":       str(val("Estado da AIMA", "") or ""),
        })

    log(f"{len(reservas)} reservas processadas.")
    return reservas

# ── Importar para Supabase ────────────────────────────────────────────────────
def importar_supabase(reservas):
    log("A ligar ao Supabase...")
    db = create_client(SUPABASE_URL, SUPABASE_KEY)

    # Guardar campos manuais antes de apagar
    log("A guardar campos manuais...")
    result = db.from_("reservas").select(
        "id,hora_checkin,hora_checkout,hora_checkin_manual,hora_checkout_manual,"
        "caucao_necessaria,caucao_cobrada,caucao_valor,pessoas_extra,"
        "custo_pessoa_extra,notas_internas,dados_pessoais_ok"
    ).execute()

    manuais = {}
    for r in (result.data or []):
        manuais[r["id"]] = r

    # Apagar tudo
    log("A limpar base de dados...")
    db.from_("reservas").delete().neq("id", "__never__").execute()

    # Reinserir com campos manuais preservados
    log("A inserir reservas...")
    reservas_final = []
    for r in reservas:
        m = manuais.get(r["id"], {})
        # Horas: preservar se marcadas como manuais, senão usar Excel
        if m.get("hora_checkin_manual"):
            r["hora_checkin"] = m.get("hora_checkin") or r["hora_checkin"]
        if m.get("hora_checkout_manual"):
            r["hora_checkout"] = m.get("hora_checkout") or r["hora_checkout"]
        # Campos sempre manuais
        r["hora_checkin_manual"]  = m.get("hora_checkin_manual", False)
        r["hora_checkout_manual"] = m.get("hora_checkout_manual", False)
        r["caucao_necessaria"]    = m.get("caucao_necessaria")
        r["caucao_cobrada"]       = m.get("caucao_cobrada", False)
        r["caucao_valor"]         = m.get("caucao_valor", 0)
        r["pessoas_extra"]        = m.get("pessoas_extra", 0)
        r["custo_pessoa_extra"]   = m.get("custo_pessoa_extra", 0)
        r["notas_internas"]       = m.get("notas_internas", "")
        r["dados_pessoais_ok"]    = m.get("dados_pessoais_ok", False)
        reservas_final.append(r)

    BATCH = 100
    for i in range(0, len(reservas_final), BATCH):
        lote = reservas_final[i:i+BATCH]
        db.from_("reservas").insert(lote).execute()
        log(f"Inseridas {min(i+BATCH, len(reservas_final))}/{len(reservas_final)} reservas...")

    log("✅ Importação concluída!")

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    log("🤖 Bot Ynnov → Supabase iniciado")
    driver, download_dir = criar_driver()
    try:
        fazer_login(driver)
        ficheiro = descarregar_excel(driver, download_dir)
        reservas = processar_excel(ficheiro)
        importar_supabase(reservas)
    finally:
        driver.quit()
        log("Browser fechado.")

if __name__ == "__main__":
    main()
