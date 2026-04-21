import os
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"

import time
from concurrent.futures import ThreadPoolExecutor
import json

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

# -------- LOG --------
def log(msg):
    print(msg, flush=True)

# -------- GOOGLE --------
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

cred_json = r'''
{
  "type": "service_account",
  "project_id": "control-firmas-juzgado",
  "private_key_id": "47b9080170a17c755bff9750f04d85b66ae5e310",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC5UZYlj6aLQ5u8\nxRzRLKdQOKqGSu7sKCAT0jBE/58B94M8tpENN/72N1ew41yssb26eoyCpAMaeFvX\n1ec1Zih7i2KdbIscNeN47npM5WnyJEhRficM/gY9UFaV3OJRfvpwb/rvcIi7GOiw\n7JViMpJb4ODXJT18wRWIwsOa46tWPXpWoZLElZGL8MVUM9KeWPDD6flwBDb5RPji\nplk4SZv9x2bcaCMsMCBeKLvYq7UWyP2oINy5BfWc6tDaqKbTNDzKFVvyK3zeDXPT\n1pdXX8VCY84i7xUxcPZWP4L0wzm7yi1KVlEdhPu9yfEgnemwzI552qHyLPY3sjQ5\n+BKXg/vHAgMBAAECggEAAgg8Vppwqm5KCD6RG6obBkIuJzLYfovMw5WLz0s4+dM2\nWTVspHHMwJ5yvmR/4P+XTl2G/5e/gPbOQFupdr3GtolyF2UtUViLRLQ4xccY66gs\n3YKTTbfWEa4OhQOFvScLUAL/rsh+d1lv6SDEXUNOCjvju02T6F5QV3jrSYjdmXXR\nAB2dyEYmPyNv5E97oUGNratZ0gQ+Q11DfCt79L+k5UFQP+CMFTrlWR/MqYllLAwm\n5nUXEbTxaC+mB8x/q5W0yTL8duDkkO3MaA0+nJx6kClI6r487JpvTekHBsxntlHZ\nzkuHnTzjpoeBn6MLDzVJkGmeKMpLLnvwP/8SEXQsoQKBgQDr+I1NLGvUktKP5L29\n0S5vmNjAjsy1rxuvfhI+kVkXaAYHlOmAxtxkdwl8LYqT26fPSNeV667GYIvl6hm8\nZQ3CBPvGeUwZH9k5aZe2rRVEUUwjYmQ6D4kfyjaIEPByWisyk6xFN0jUQ4qCqz8z\nG+Q3sbXTIxUBn+ozjQNJhidCxQKBgQDJDGepnHZB+fFu/qUfdk2ORNK1x8i9v6S1\npYoV0dPo9G8k6SuIHX92E9V4L9vFhqaLAjLiil4mSNGKuJAOBaUo/izweVbNAjrg\ndvTJNNOVnJrYVnQtPLs5sH0CjRv6Lj6E34JrkQn3tt4C8v7DFBJVs4mIIgyRDO7s\nwDJE2As9GwKBgFuwS29WOFvz5N9GmTd9ZVa1hFtl4UMjVFWfXgVzwrNmlxkxEn4Y\nRyC+ZDAdHgCP1Cel/Sbi2hl5AEMI8JEUjwD5oL8g+KG2j1hQoEO6A051bGk/XQR2\nbuisUP4T3uoAAVL4sHKApcrcp6BYXAlG6Cl/4s+0jQABnCYFv+Y8u1qlAoGBAK7K\nvC14PFBsD33ioojR/+ea6l7kjSB7R6Ytf/osbUJxkVfT6Ob0TmbII6XUZgw7Xvwo\nMzlF90jtslAa2hN20Prs0QFZXR/rumiAw51S8kl22CESOPtDe7tSN71KFXLTVcOL\n1wXSGYpmUGrb/KZ6At7DsuTKRYauaeMnzgyQkGTVAoGBAOI9jcK0IQI1lE5q2fT6\nDzzb6Ae/rV2XzJ/DFTpgZGVMy2GxoMD7wUhGgjyo4yNP4d4+IsCJjwWI4ALGVD7s\nyWb32Amf4qApmoGUpbYelc/KGwpK7e4CaSN9Tugg9OKwc65aiTLbrE/vipbUBpKm\nqQnKuXt5f49pyov98pTSO1MA\n-----END PRIVATE KEY-----\n",
  "client_email": "bot-juzgado@control-firmas-juzgado.iam.gserviceaccount.com",
  "client_id": "111990996372963001151",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/bot-juzgado%40control-firmas-juzgado.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}
'''

creds_dict = json.loads(cred_json)
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
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

                        if variante != exp:
                            log(f"🔧 {exp} → {variante}")
                            corregir_expediente(sheet, i, variante)

                        break
                    except:
                        continue

            if not encontrado:
                escribir_no(sheet, i)
                log(f"❌ {exp}")
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