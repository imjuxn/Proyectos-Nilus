import os
import sys
import ssl
import time
import pandas as pd
import gspread

from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import *
from selenium.webdriver.common.keys import Keys

from slack_sdk import WebClient
from oauth2client.service_account import ServiceAccountCredentials

from dotenv import load_dotenv
load_dotenv()

SLACK_TOKEN = os.getenv("SLACK_TOKEN")
CHANNEL_ID = os.getenv("CHANNEL_ID")

EMAIL_LOGIN = os.getenv("EMAIL_LOGIN")
PASSWORD_LOGIN = os.getenv("PASSWORD_LOGIN")

SHEET_ID = os.getenv("SHEET_ID_CANCELAR")
GOOGLE_JSON = os.getenv("GOOGLE_JSON")
SHEET_NAME = os.getenv("SHEET_NAME_CANCELAR")


def obtener_ruta_certificado():
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, "certificados", "cacert.pem")


ssl_context = ssl.create_default_context(cafile=obtener_ruta_certificado())
client = WebClient(token=SLACK_TOKEN, ssl=ssl_context)


def enviar_notificacion_slack(mensaje):
    try:
        response = client.chat_postMessage(channel=CHANNEL_ID, text=mensaje)
    except:
        pass


enviar_notificacion_slack("PREVENCIÓN ARG y MX iniciado.")


def click_button(driver, selector, by=By.CSS_SELECTOR, wait_time=10):
    try:
        button = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((by, selector))
        )
        button.click()
        return True
    except:
        return False


def guardar_cambios(driver, wait_time=10):
    try:
        modal = WebDriverWait(driver, wait_time).until(
            EC.visibility_of_element_located((By.XPATH,
                "//h2[contains(text(), 'Cambiar el estado del pedido')]/ancestor::div[@role='dialog']")
            )
        )
        guardar_btn = modal.find_element(By.XPATH, ".//button[text()='Guardar cambios']")
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", guardar_btn)
        driver.execute_script("arguments[0].click();", guardar_btn)
        return True
    except:
        return False


if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

RUTA_CREDENCIALES = os.path.join(base_path, GOOGLE_JSON)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciales = ServiceAccountCredentials.from_json_keyfile_name(RUTA_CREDENCIALES, scope)
cliente = gspread.authorize(credenciales)

sheet = cliente.open_by_key(SHEET_ID)
worksheet = sheet.worksheet(SHEET_NAME)
valores = worksheet.get_all_values()

df = pd.DataFrame(valores[1:], columns=valores[0])

options = Options()
options.add_argument("--headless=new")
options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(
    service=ChromeService(ChromeDriverManager().install()),
    options=options
)

driver.get("https://backoffice.nilus.co/es-AR/login")
time.sleep(5)

driver.find_element(By.ID, "email").send_keys(EMAIL_LOGIN)
driver.find_element(By.ID, "password").send_keys(PASSWORD_LOGIN)
click_button(driver, "//button[text()='INGRESAR']", By.XPATH)

time.sleep(5)

for i, row in df.iterrows():
    pedido_id = str(row.iloc[1]).strip()
    if not pedido_id:
        enviar_notificacion_slack(f"ID inválido en fila {i}")
        continue

    try:
        driver.get(f"https://backoffice.nilus.co/es-AR/orders/{pedido_id}")
    except:
        enviar_notificacion_slack(f"No se pudo abrir pedido {pedido_id}")
        continue

    try:
        try:
            combo = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@id='email']"))
            )
        except:
            combo = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@id='status']"))
            )

        combo.click()

        opcion = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Cancelado')]"))
        )
        opcion.click()

        motivo_combo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "reason_of_canceled"))
        )
        motivo_combo.click()

        try:
            motivo = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Pago anticipado')]"))
            )
            motivo.click()
        except:
            enviar_notificacion_slack(f"Motivo no encontrado {pedido_id}")

        guardar_cambios(driver)

    except:
        enviar_notificacion_slack(f"Error procesando {pedido_id}")
        continue

driver.quit()
sys.exit()
