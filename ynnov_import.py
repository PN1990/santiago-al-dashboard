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

    log("A preencher email...")
    campo_email = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email'], input[name='email'], input[placeholder*='email' i]")))
    esperar(0.5, 1.5)
    campo_email.clear()
    for char in YNNOV_EMAIL:
        campo_email.send_keys(char)
        time.sleep(random.uniform(0.05, 0.15))
    esperar(0.5, 1.5)

    log("A preencher password...")
    campo_pass = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
    campo_pass.clear()
    for char in YNNOV_PASSWORD:
        campo_pass.send_keys(char)
        time.sleep(random.uniform(0.05, 0.15))
    esperar(0.8, 2)

    driver.save_screenshot("/tmp/ynnov_before_login.png")
    log(f"Screenshot tirado. URL atual: {driver.current_url}")
    log(f"Título da página: {driver.title}")

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
                btn_login = elementos[-1]
                log(f"Botão encontrado com: {seletor} -- texto: '{btn_login.text}'")
                break
        except:
            continue

    if not btn_login:
        log("Botão não encontrado, a tentar Enter no campo password...")
        from selenium.webdriver.common.keys import Keys
        campo_pass.send_keys(Keys.RETURN)
    else:
        esperar(0.5, 1)
        btn_login.click()

    esperar(4, 6)
    driver.save_screenshot("/tmp/ynnov_after_login.png")
    log(f"URL após login: {driver.current_url}")

    if "login" in driver.current_url.lower():
        log("Login pode ter falhado. HTML da página:")
        log(driver.page_source[:500])
        raise Exception("Login falhou! Ver screenshots de debug.")
    log(f"Login OK -- URL: {driver.current_url}")

def clicar_elemento_por_texto(driver, texto, timeout=15):
    """Tenta clicar num elemento pelo texto, com múltiplos métodos."""
    xpaths = [
        f"//*[normalize-space(text())='{texto}']",
        f"//*[contains(text(),'{texto}')]",
        f"//button[normalize-space(.)='{texto}']",
        f"//button[contains(.,'{texto}')]",
        f"//*[@title='{texto}']",
    ]
    # Método 1: WebDriverWait
    for xpath in xpaths:
        try:
            wait = WebDriverWait(driver, timeout // len(xpaths))
            el = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            esperar(0.3, 0.8)
            try:
                el.click()
            except:
                driver.execute_script("arguments[0].click();", el)
            log(f"Clicado '{texto}' via xpath: {xpath}")
            return el
        except:
            continue

    # Método 2: JavaScript direto — procurar por texto em todos os elementos
    log(f"A tentar clicar '{texto}' via JavaScript...")
    resultado = driver.execute_script(f"""
        var elements = document.querySelectorAll('button, a, div, span');
        for (var i = 0; i < elements.length; i++) {{
            if (elements[i].textContent.trim() === '{texto}') {{
                elements[i].click();
                return true;
            }}
        }}
        return false;
    """)
    if resultado:
        log(f"Clicado '{texto}' via JavaScript.")
        return True

    raise Exception(f"Não foi possível clicar em '{texto}'")

def clicar_texto(driver, texto, timeout=15):
    """Wrapper para compatibilidade."""
    return clicar_elemento_por_texto(driver, texto, timeout)

def descarregar_excel(driver, download_dir):
    import glob
    wait = WebDriverWait(driver, 20)

    # 1. Clicar em Reservas no menu lateral
    log("A clicar em Reservas...")
    clicar_elemento_por_texto(driver, "Reservas")
    esperar(2, 3)
    driver.save_screenshot("/tmp/ynnov_reservas.png")

    # 2. Clicar em Lista
    log("A clicar em Lista...")
    clicar_elemento_por_texto(driver, "Lista")
    esperar(1, 2)
    driver.save_screenshot("/tmp/ynnov_lista.png")

    # 3. Clicar em Filtros
    log("A clicar em Filtros...")
    clicar_elemento_por_texto(driver, "Filtros")
    esperar(1, 2)
    driver.save_screenshot("/tmp/ynnov_filtros.png")

    # 4. Limpar filtros de estado
    log("A limpar filtros de estado...")
    try:
        btn_clear = driver.find_element(By.XPATH, "//button[normalize-space(text())='Clear' or normalize-space(text())='Limpar']")
        btn_clear.click()
        esperar(0.5, 1)
    except:
        log("Botão clear não encontrado, a tentar desselecionar todos...")
        try:
            btn_all = driver.find_element(By.XPATH, "//button[normalize-space(text())='All']")
            btn_all.click()
            esperar(0.5, 1)
            btn_all.click()
            esperar(0.5, 1)
        except:
            pass

    # 5. Selecionar Confirmado, Check-in e Check-out
    for estado in ["Confirmado", "Check-in", "Check-out"]:
        log(f"A selecionar estado: {estado}...")
        try:
            els = driver.find_elements(By.XPATH, f"//span[normalize-space(text())='{estado}'] | //li[normalize-space(text())='{estado}'] | //div[normalize-space(text())='{estado}']")
            if els:
                driver.execute_script("arguments[0].click();", els[0])
                esperar(0.3, 0.8)
                log(f"Estado '{estado}' selecionado.")
            else:
                log(f"Aviso: elemento '{estado}' não encontrado.")
        except Exception as e:
            log(f"Aviso: erro ao selecionar '{estado}': {e}")

    driver.save_screenshot("/tmp/ynnov_filtros_selecionados.png")

    # 6. Clicar em Aplicar — método robusto com múltiplos fallbacks
    log("A clicar em Aplicar...")
    aplicar_ok = False
    aplicar_xpaths = [
        "//button[normalize-space(text())='Aplicar']",
        "//button[contains(text(),'Aplicar')]",
        "//button[normalize-space(.)='Aplicar']",
        "//*[normalize-space(text())='Aplicar']",
        "//*[contains(@class,'apply')]",
        "//button[contains(@class,'btn')][last()]",
    ]
    for xpath in aplicar_xpaths:
        try:
            els = driver.find_elements(By.XPATH, xpath)
            if els:
                driver.execute_script("arguments[0].click();", els[-1])
                log(f"Aplicar clicado via: {xpath}")
                aplicar_ok = True
                break
        except:
            continue

    if not aplicar_ok:
        # Fallback: JavaScript direto
        resultado = driver.execute_script("""
            var btns = document.querySelectorAll('button');
            for (var i = 0; i < btns.length; i++) {
                if (btns[i].textContent.trim() === 'Aplicar') {
                    btns[i].click();
                    return true;
                }
            }
            return false;
        """)
        if resultado:
            log("Aplicar clicado via JavaScript fallback.")
            aplicar_ok = True

    if not aplicar_ok:
        log("AVISO: Botão Aplicar não encontrado, a continuar sem filtros...")

    esperar(2, 3)
    driver.save_screenshot("/tmp/ynnov_apos_filtros.png")

    # 7. Clicar no botão XLS
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
    import pandas as pd
    if valor is None or (isinstance(valor, float) and str(valor) == 'nan'):
        return None
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

def parse_hora(valor):
    if valor is None:
        return ""
    import math
    if isinstance(valor, float):
        if math.isnan(valor):
            return ""
        total_min = round(valor * 24 * 60)
        h = total_min // 60
        m = total_min % 60
        return f"{h:02d}:{m:02d}"
    if hasattr(valor, 'strftime'):
        return valor.strftime("%H:%M")
    s = str(valor).strip()
    if s in ('', 'nan', 'NaT', 'None', '-'):
        return ""
    if ':' in s:
        return s[:5]
    return ""

def converter_xls_para_xlsx(ficheiro):
    import subprocess, glob, tempfile
    out_dir = tempfile.mkdtemp()
    log(f"A converter XLS para XLSX com LibreOffice...")
    result = subprocess.run([
        'libreoffice', '--headless', '--convert-to', 'xlsx',
        '--outdir', out_dir, ficheiro
    ], capture_output=True, text=True, timeout=60)
    log(f"LibreOffice stdout: {result.stdout}")
    log(f"LibreOffice stderr: {result.stderr}")
    xlsx_files = glob.glob(os.path.join(out_dir, "*.xlsx"))
    if not xlsx_files:
        raise Exception("LibreOffice não gerou XLSX.")
    return xlsx_files[0]

def processar_excel(ficheiro):
    log(f"A processar Excel: {ficheiro}")
    import pandas as pd

    df = None

    for engine in ['xlrd', 'openpyxl']:
        try:
            df = pd.read_excel(ficheiro, engine=engine)
            log(f"Excel lido com engine: {engine}")
            break
        except Exception as e:
            log(f"Engine {engine} falhou: {e}")

    if df is None:
        try:
            xlsx_file = converter_xls_para_xlsx(ficheiro)
            log(f"A ler XLSX convertido: {xlsx_file}")
            df = pd.read_excel(xlsx_file, engine='openpyxl')
            log(f"XLSX lido com sucesso: {len(df.columns)} colunas, {len(df)} linhas")
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

        def safe_int(v, default=0):
            try:
                if v is None or str(v).strip() in ('nan', 'NaN', '', 'None'):
                    return default
                return int(float(v))
            except:
                return default

        def safe_float(v, default=0.0):
            try:
                if v is None:
                    return default
                import math
                if isinstance(v, float) and math.isnan(v):
                    return default
                s = str(v).strip()
                if s in ('nan', 'NaN', '', 'None', '-', 'N/A'):
                    return default
                if ',' in s and '.' not in s:
                    s = s.replace(',', '.')
                elif ',' in s and '.' in s:
                    s = s.replace('.', '').replace(',', '.')
                return float(s)
            except:
                return default

        def val(key, default=""):
            v = r.get(key, default)
            if v is None or (isinstance(v, float) and str(v) == 'nan'):
                return default
            return v

        if len(reservas) < 3:
            campos_monetarios = ["Total da reserva", "Pagamento", "Em dívida", "Comissão do canal", "TMT"]
            log(f"=== LINHA {len(reservas)+1} ===")
            for c in campos_monetarios:
                log(f"  '{c}': {repr(r.get(c))}")

        id_reserva = str(val("ID")).strip().replace(".0", "")
        hospede    = str(val("Hóspede", val("Hospede"))).strip()

        if not id_reserva or id_reserva in ('', 'nan') or not hospede:
            continue

        reservas.append({
            "id":                id_reserva,
            "hospede":           hospede,
            "checkin":           parse_data(val("Data de check-in")),
            "hora_checkin":      parse_hora(val("Hora de check-in", "")),
            "checkout":          parse_data(val("Data de check-out")),
            "hora_checkout":     parse_hora(val("Hora de check-out", "")),
            "noites":            safe_int(r.get("N.º de noites", val("N de noites", 0))),
            "adultos":           safe_int(val("Adultos", 0)),
            "criancas":          safe_int(r.get("Crianças", val("Criancas", 0))),
            "bebes":             safe_int(r.get("Bebés", val("Bebes", 0))),
            "telefone":          str(val("Telefone", "") or ""),
            "email":             str(r.get("Email (pessoal)", val("Email (canal)", "")) or ""),
            "pais":              str(r.get("País", val("Pais", "")) or ""),
            "codigo_pais":       str(r.get("Código do país", val("Codigo do pais", "")) or ""),
            "alojamento":        str(val("Alojamento", "") or ""),
            "tmt":               safe_float(val("TMT", 0)),
            "total":             (lambda t, c, pct: t if t > 0 else (round(c / (pct if pct > 0 else 15) * 100, 2) if c > 0 else 0))(
                                     safe_float(val("Total da reserva", val("Pagamento", 0))),
                                     safe_float(r.get("Comissão do canal", val("Comissao do canal", 0))),
                                     safe_float(val("Comissão do canal (%)", 0))
                                 ),
            "estado":            str(val("Estado da reserva", "") or ""),
            "estado_pagamento":  str(val("Estado do pagamento", "") or ""),
            "canal":             str(val("Canal", "") or ""),
            "comissao":          safe_float(r.get("Comissão do canal", val("Comissao do canal", 0))),
            "comissao_pct":      safe_float(val("Comissão do canal (%)", 0)),
            "id_canal":          str(val("ID da reserva (canal)", "") or ""),
            "data_criacao":      parse_data(r.get("Data de criação", val("Data de criacao"))),
            "antecedencia":      safe_int(r.get("Antecedência da reserva (dias)", val("Antecedencia da reserva (dias)", 0))),
            "checkin_efetuado":  str(val("Check-in efetuado a...", "") or ""),
            "checkout_efetuado": str(val("Check-out efetuado a...", "") or ""),
            "notas_canal":       str(val("Notas do canal", "") or ""),
            "fatura":            str(val("Fatura(s)", "") or ""),
            "estado_aima":       str(val("Estado da AIMA", "") or ""),
        })

    from datetime import date
    hoje = date.today().strftime("%Y-%m-%d")
    estados_excluir = {"Cancelado", "Cancelada", "Cancelled", "No show"}
    reservas_filtradas = [
        r for r in reservas
        if r.get("estado", "") not in estados_excluir
        and r.get("id")
    ]
    for r in reservas_filtradas[:3]:
        log(f"Debug reserva {r['id']}: estado={r.get('estado')}, total={r.get('total')}, comissao={r.get('comissao')}")
    log(f"{len(reservas)} reservas lidas, {len(reservas_filtradas)} válidas (excluídos cancelados).")
    return reservas_filtradas

# ── Importar para Supabase ────────────────────────────────────────────────────
def importar_supabase(reservas):
    log("A ligar ao Supabase...")
    db = create_client(SUPABASE_URL, SUPABASE_KEY)

    log("A guardar campos manuais...")
    result = db.from_("reservas").select(
        "id,hora_checkin,hora_checkout,hora_checkin_manual,hora_checkout_manual,"
        "caucao_necessaria,caucao_cobrada,caucao_valor,pessoas_extra,"
        "custo_pessoa_extra,notas_internas,dados_pessoais_ok"
    ).execute()

    manuais = {}
    for r in (result.data or []):
        manuais[r["id"]] = r

    log("A limpar base de dados...")
    db.from_("reservas").delete().neq("id", "__never__").execute()

    log("A inserir reservas...")
    reservas_final = []
    for r in reservas:
        m = manuais.get(r["id"], {})
        if m.get("hora_checkin_manual"):
            r["hora_checkin"] = m.get("hora_checkin") or r["hora_checkin"]
        if m.get("hora_checkout_manual"):
            r["hora_checkout"] = m.get("hora_checkout") or r["hora_checkout"]
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
