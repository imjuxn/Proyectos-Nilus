import os
from dotenv import load_dotenv
load_dotenv()

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

SLACK_TOKEN = os.getenv("SLACK_TOKEN_HD")
CHANNEL_ID = os.getenv("CHANNEL_ID_HD")

def obtener_ruta_certificado():
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, "certificados", "cacert.pem")

ssl_context = ssl.create_default_context(cafile=obtener_ruta_certificado())
client = WebClient(token=SLACK_TOKEN, ssl=ssl_context)

CHANNEL_ID_NOTIFICACIONES = os.getenv("CHANNEL_ID_NOTIF_HD")

def enviar_notificacion_slack(mensaje):
    try:       
        response = client.chat_postMessage(channel=CHANNEL_ID_NOTIFICACIONES, text=mensaje) 
        if not response["ok"]:
            print(f"❌ Error enviando mensaje a Slack: {response['error']}")
    except Exception as e:
        print(f"❌ Excepción enviando mensaje a Slack: {e}")

enviar_notificacion_slack(mensaje=f"El script de HOMEDELIVERY PARA ARG Y MX ha comenzado 🚀")

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
    

SHEET_ID = os.getenv("GOOGLE_SHEET_ID_HD")

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

json_filename = os.getenv("GOOGLE_CREDENTIALS_FILE")
RUTA_CREDENCIALES = os.path.join(base_path, json_filename)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciales = ServiceAccountCredentials.from_json_keyfile_name(RUTA_CREDENCIALES, scope)
cliente = gspread.authorize(credenciales)

sheet = cliente.open_by_key(SHEET_ID)
worksheet = sheet.worksheet("Order_ids")
valores = worksheet.get_all_values()

df = pd.DataFrame(valores[1:], columns=valores[0])

hoy = datetime.now().strftime("%d/%m/%Y")

print(f"📌 Pedidos de hoy ({hoy}):")

options = Options()
options.add_argument("--headless=new")
options.add_argument("--window-size=1920,1080")
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

driver.get("https://backoffice.nilus.co/es-AR/login")
time.sleep(5)
driver.find_element(By.ID, "email").send_keys(os.getenv("NILUS_LOGIN_HD"))
driver.find_element(By.ID, "password").send_keys(os.getenv("NILUS_PASSWORD_HD"))
click_button(driver, "//button[text()='INGRESAR']", By.XPATH)
time.sleep(10)

for i, row in df.iterrows():
    pedido_id = str(row.iloc[1]).strip()

    if not pedido_id:
        mensaje = f"❌ No se encontró un ID válido en: {pedido_id}"
        print(mensaje)
        enviar_notificacion_slack(mensaje)
        continue

    print(f"🔄 Procesando pedido {pedido_id} con motivo CLiente prófugo")
    try:
        driver.get(f"https://backoffice.nilus.co/es-AR/orders/{pedido_id}")
    except Exception as e:
        mensaje = f"⚠️ Error al abrir el pedido {pedido_id}: {e}"
        print(mensaje)
        enviar_notificacion_slack(mensaje)
        continue
    
    try:
        try:
            cambiar_estado = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='email']"))
            )
            cambiar_estado.click()
        except TimeoutException:
            mensaje = "⚠️ No se pudo hacer clic en cambiar estado intento 1'. Intentando con id='status'..."
            print(mensaje)    

            try:
                cambiar_estado = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='status']"))
                )
                cambiar_estado.click()                
            except TimeoutException:
                mensaje = "❌ No se pudo hacer clic en cambiar estado intento 2. Continuando con el siguiente pedido..."
                print(mensaje)
                enviar_notificacion_slack(mensaje)
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
            Entrega_fallida = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Cliente prófugo')]"))
            )
            Entrega_fallida.click()
        except Exception as e2:
            mensaje = f"❌ Tampoco se pudo seleccionar el motivo por defecto para el pedido {pedido_id}"
            print(mensaje)
            enviar_notificacion_slack(mensaje)

        guardar_cambios(driver)
        print(f"✅ Pedido {pedido_id} cancelado con motivo 'Cliente prófugo'")
    except Exception as e:
        mensaje = f"❌ Error procesando pedido {pedido_id}: TAL VEZ YA ESTA ANULADO"
        print(mensaje)
        enviar_notificacion_slack(mensaje)
        continue

print("✅ Script finalizado correctamente.")
driver.quit()
sys.exit()
