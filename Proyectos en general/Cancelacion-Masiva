import os
from selenium.webdriver.chrome.options import Options
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import *
from datetime import datetime, timedelta
from slack_sdk import WebClient
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from selenium.webdriver.common.keys import Keys
import time
import re
import pytz
import ssl
import json
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

def guardar_cambios(driver, wait_time=10):
    try:
        modal = WebDriverWait(driver, wait_time).until(
            EC.visibility_of_element_located((By.XPATH, "//h2[contains(text(), 'Cambiar el estado del pedido')]/ancestor::div[@role='dialog']"))
        )
        guardar_btn = modal.find_element(By.XPATH, ".//button[normalize-space(text())='Guardar cambios']")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", guardar_btn)
        driver.execute_script("arguments[0].click();", guardar_btn)
        print("✅ 'Guardar cambios' se clickeó correctamente.")
        return True
    except Exception as e:
        print(f"❌ No se pudo hacer clic en el botón correcto: {e}")
        return False

def extraer_id(texto):
    match = re.search(r"[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}", str(texto))
    return match.group(0) if match else None

SHEET_ID = os.getenv("SHEET_ID")

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

json_filename = os.getenv("GOOGLE_JSON")
RUTA_CREDENCIALES = os.path.join(base_path, json_filename)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciales = ServiceAccountCredentials.from_json_keyfile_name(RUTA_CREDENCIALES, scope)
cliente = gspread.authorize(credenciales)

sheet = cliente.open_by_key(SHEET_ID)
worksheet = sheet.sheet1
valores = worksheet.get_all_values()

df = pd.DataFrame(valores[1:], columns=valores[0])

hoy = datetime.now().strftime("%d/%m/%Y")
df_hoy = df[df.iloc[:, 0] == hoy]

print(f"📌 Pedidos de hoy ({hoy}):")
print(df_hoy)

options = Options()
options.add_argument("--headless=new")
options.add_argument("--window-size=1920,1080")
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

driver.get("https://backoffice.nilus.co/es-AR/login")
time.sleep(5)
driver.find_element(By.ID, "email").send_keys(os.getenv("NILUS_EMAIL"))
driver.find_element(By.ID, "password").send_keys(os.getenv("NILUS_PASSWORD"))
click_button(driver, "//button[text()='INGRESAR']", By.XPATH)
time.sleep(5)

for i, row in df_hoy.iterrows():
    pedido_raw = row.iloc[1]
    motivo = row.iloc[2]
    pedido_id = extraer_id(pedido_raw)

    if not pedido_id:
        print(f"❌ No se encontró un ID válido en: {pedido_raw}")
        worksheet.update_cell(i+2, 4, "❌ ID inválido")
        continue

    print(f"🔄 Procesando pedido {pedido_id} con motivo '{motivo}'...")
    try:
        driver.get(f"https://backoffice.nilus.co/es-AR/orders/{pedido_id}")
    except Exception as e:
        print(f"⚠️ Error al abrir el pedido {pedido_id}: {e}")
        worksheet.update_cell(i+2, 4, "❌")
        continue

    try:
        try:
            cambiar_estado = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='email']"))
            )
            cambiar_estado.click()
        except TimeoutException:
            try:
                cambiar_estado = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='status']"))
                )
                cambiar_estado.click()
            except TimeoutException:
                print("❌ No se pudo hacer clic en cambiar estado intento 2. Continuando...")
                continue

        cancelado_opcion = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Cancelado')]"))
        )
        cancelado_opcion.click()
        time.sleep(1)

        motivo_combo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "reason_of_canceled"))
        )
        motivo_combo.click()
        time.sleep(2)

        try:
            MOTIVOS = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[@role='option' and contains(text(), '{motivo}')]"))
            )
            MOTIVOS.click()
        except Exception as e:
            print(f"❌ No se pudo seleccionar el motivo '{motivo}'. Error: {e}")

        guardar_cambios(driver)
        print(f"✅ Pedido {pedido_id} cancelado con motivo '{motivo}'")
        worksheet.update_cell(i+2, 4, "✅ Hecho")
    except Exception as e:
        print(f"❌ Error procesando pedido {pedido_id}: {e}")
        worksheet.update_cell(i+2, 4, "❌")
        continue
