import os
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"

import time
from concurrent.futures import ThreadPoolExecutor

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright

# -------- CONFIG --------
MAX_WORKERS = 4
MAX_SHEETS = 3

SHEETS_URLS = [
    "https://docs.google.com/spreadsheets/d/1tIDH2n4SqyHKpSoHGxZ2pXx1R3XeACLR6nemdLBtqlo",
    "https://docs.google.com/spreadsheets/d/1Kf3uQUU8VH0ZSpmmk8RCs1Zdp3DNCQHEZkQWMNbpYAs",
    "https://docs.google.com/spreadsheets/d/122nbZPlzm4hujZtZn9YU5qvDlw2sEc1TwJD2hxyNFv4"
]

# -------- LOG --------
def log(msg):
    print(msg, flush=True)

# -------- GOOGLE (ARCHIVO LOCAL) --------
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_name(
    "credenciales.json",
    scope
)

client = gspread.authorize(creds)

# -------- FORMATO --------
def limpiar_fila(sheet, fila):
    sheet.update(values=[[""]], range_name=f"F{fila}")
    sheet.format(f"A{fila}:Z{fila}", {
        "backgroundColor": {"red": 1, "green": 1, "blue": 1}
    })

def pintar_verde(sheet, fila):
    sheet.format(f"F{fila}", {
        "backgroundColor": {"red": 0, "green": 0.6, "blue": 0}
    })

def escribir_si(sheet, fila):
    sheet.update(values=[["SI"]], range_name=f"F{fila}")

def escribir_no(sheet, fila):
    sheet.update(values=[["NO"]], range_name=f"F{fila}")

def corregir_expediente(sheet, fila, nuevo_exp):
    sheet.update(values=[[nuevo_exp]], range_name=f"C{fila}")

# -------- VARIANTES --------
def variantes_expediente(exp):
    exp = exp.strip()
    if "-" in exp:
        return [exp]
    return [f"{exp}-0", exp]

# -------- BUSCAR ACTUACION --------
def buscar_actuacion(page, act):
    act_num = act.split("/")[0]

    for _ in range(10):
        if page.locator(f"text={act_num}").count() > 0:
            return True
        page.mouse.wheel(0, 1000)
        time.sleep(0.4)

    return False

# -------- WORKER --------
def procesar_expediente(sheet, i, caratula, exp, act):
    try:
        if not act.strip():
            return

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()

            page.goto("https://eje.juscaba.gob.ar/iol-ui/p/inicio")

            input_box = page.locator("#inputSearch")
            encontrado = False

            for variante in variantes_expediente(exp):
                log(f"🔎 {variante}")

                input_box.fill("")
                input_box.fill(variante)
                page.keyboard.press("Enter")

                time.sleep(1)

                try:
                    page.locator(f"text={caratula[:30]}").first.click()
                    encontrado = True

                    if variante != exp:
                        log(f"🔧 {exp} → {variante}")
                        corregir_expediente(sheet, i, variante)

                    break
                except:
                    try:
                        page.locator("div[role='row']").nth(1).click()
                        encontrado = True
                        break
                    except:
                        continue

            if not encontrado:
                escribir_no(sheet, i)
                log(f"❌ {exp}")
                browser.close()
                return

            page.locator("text=Actuaciones").first.click()
            time.sleep(1)

            if buscar_actuacion(page, act):
                limpiar_fila(sheet, i)
                escribir_si(sheet, i)
                pintar_verde(sheet, i)
                log(f"✔ {exp}")
            else:
                escribir_no(sheet, i)
                log(f"✖ {exp}")

            browser.close()

    except Exception as e:
        log(f"Error {exp}: {e}")

# -------- PROCESAR SHEET --------
def procesar_sheet(url):
    log(f"📄 Procesando: {url}")

    spreadsheet = client.open_by_url(url)
    hojas = spreadsheet.worksheets()[:10]

    for sheet in hojas:
        filas = sheet.get_all_values()

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            tareas = []

            for i, fila in enumerate(filas[1:], start=2):
                try:
                    caratula = fila[1]
                    exp = fila[2]
                    act = fila[3]
                    firmado = fila[5]
                except:
                    continue

                if firmado.strip().upper() == "SI":
                    continue

                if not act.strip():
                    continue

                tareas.append(
                    executor.submit(
                        procesar_expediente,
                        sheet, i, caratula, exp, act
                    )
                )

            for t in tareas:
                t.result()

# -------- MAIN --------
def procesar():
    log("=== INICIO ===")

    with ThreadPoolExecutor(max_workers=MAX_SHEETS) as executor:
        futures = [executor.submit(procesar_sheet, url) for url in SHEETS_URLS]

        for f in futures:
            f.result()

    log("=== FIN ===")

# -------- ENTRY --------
if __name__ == "__main__":
    procesar()
