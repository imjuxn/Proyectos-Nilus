import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import *
from datetime import datetime, timedelta
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.options import Options
from oauth2client.service_account import ServiceAccountCredentials
from difflib import SequenceMatcher
from selenium.webdriver.common.keys import Keys
import gspread
import time
import requests
import re
import sys
import os
from dotenv import load_dotenv

load_dotenv()

def click_button(driver, selector, by=By.CSS_SELECTOR, wait_time=10):
    try:
        button = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((by, selector))
        )
        button.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en el selector '{selector}': {e}")
        return False

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

SHEET_ID = os.getenv("SHEET_ID")

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

json_filename = os.getenv("GOOGLE_JSON")
RUTA_CREDENCIALES = os.path.join(base_path, json_filename)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciales = ServiceAccountCredentials.from_json_keyfile_name(RUTA_CREDENCIALES, scope)
cliente = gspread.authorize(credenciales)

sheet = cliente.open_by_key(SHEET_ID)
worksheet = sheet.sheet1
worksheet = sheet.worksheet("check_nueva_info_desvios")
valores = worksheet.get_all_values()

df = pd.DataFrame(valores[1:], columns=valores[0])

hoy = datetime.now().strftime("%d/%m/%Y")

df_hoy = df[(df["Fecha"] == hoy) & (df.iloc[:, 1].str.lower() == "ar")]

driver.get(os.getenv("NILUS_LOGIN_URL"))
print("Presiona ENTER para CONTINUAR...")
input()
search_box = driver.find_element(By.CSS_SELECTOR, "input.o_searchview_input")

pedidos = df_hoy.iloc[:, 8].tolist()
print(f"📌 Pedidos de hoy ({hoy}): {pedidos}")

for pedido in pedidos:
    search_box.send_keys(pedido)

    WebDriverWait(driver, 1200).until(
        EC.invisibility_of_element((By.CLASS_NAME, "o_blockUI"))
    )

    enlace_doc_origen = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, f'//a[.//em[contains(text(), "Documento origen")] and .//strong[contains(text(), "{pedido}")]]')
        )
    )
    enlace_doc_origen.click()

    search_box.send_keys(Keys.ENTER)

    time.sleep(10)

print("Presiona ENTER para FINALIZAR...")
input()
