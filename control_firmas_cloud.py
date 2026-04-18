import time
import os
import json
from concurrent.futures import ThreadPoolExecutor

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright

MAX_WORKERS = 4
MAX_SHEETS = 3

SHEETS_URLS = [
    "https://docs.google.com/spreadsheets/d/1tIDH2n4SqyHKpSoHGxZ2pXx1R3XeACLR6nemdLBtqlo",
    "https://docs.google.com/spreadsheets/d/1Kf3uQUU8VH0ZSpmmk8RCs1Zdp3DNCQHEZkQWMNbpYAs",
    "https://docs.google.com/spreadsheets/d/122nbZPlzm4hujZtZn9YU5qvDlw2sEc1TwJD2hxyNFv4"
]

def log(msg):
    print(msg, flush=True)

# -------- GOOGLE --------
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

creds_dict = json.loads(os.environ["GOOGLE_CREDS"])
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# -------- SHEET --------
def escribir_si(sheet, fila):
    sheet.update(values=[["SI"]], range_name=f"F{fila}")

def escribir_no(sheet, fila):
    sheet.update(values=[["NO"]], range_name=f"F{fila}")

# -------- CAPTURAR ACTUACIONES --------
def obtener_actuaciones(page):
    actuaciones = []

    def handle_response(response):
        try:
            if "actuaciones" in response.url.lower():
                data = response.json()
                actuaciones.append(data)
        except:
            pass

    page.on("response", handle_response)

    return actuaciones

# -------- WORKER --------
def procesar_expediente(sheet, i, caratula, exp, act):
    try:
        if not act.strip():
            return

        act_num = act.split("/")[0]

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()

            actuaciones_data = obtener_actuaciones(page)

            page.goto("https://eje.juscaba.gob.ar/iol-ui/p/inicio")

            input_box = page.locator("#inputSearch")
            input_box.fill(exp)
            page.keyboard.press("Enter")

            page.wait_for_timeout(2000)

            try:
                page.locator("div[role='row']").nth(1).click()
            except:
                escribir_no(sheet, i)
                log(f"❌ {exp}")
                browser.close()
                return

            page.locator("text=Actuaciones").first.click()
            page.wait_for_timeout(3000)

            encontrado = False

            # 🔥 BUSCAR EN RESPUESTAS API
            for bloque in actuaciones_data:
                texto = json.dumps(bloque)

                if act_num in texto:
                    encontrado = True
                    break

            if encontrado:
                escribir_si(sheet, i)
                log(f"✔ {exp}")
            else:
                escribir_no(sheet, i)
                log(f"✖ {exp}")

            browser.close()

    except Exception as e:
        log(f"Error {exp}: {e}")

# -------- SHEET --------
def procesar_sheet(url):
    log(f"📄 {url}")

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

if __name__ == "__main__":
    procesar()
