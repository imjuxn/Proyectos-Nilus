import os
import sys
import shutil
import time
import json
from datetime import datetime, timedelta


#   CARGA DE VARIABLES .ENV

from dotenv import load_dotenv
load_dotenv()

EMAIL_LOGIN = os.getenv("EMAIL_LOGIN")
PASSWORD_LOGIN = os.getenv("PASSWORD_LOGIN")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME")
GOOGLE_CREDS_FILE = os.getenv("GOOGLE_CREDS")
SLACK_TOKEN = os.getenv("SLACK_TOKEN")

if not all([EMAIL_LOGIN, PASSWORD_LOGIN, SPREADSHEET_ID, SHEET_NAME, GOOGLE_CREDS_FILE]):
    print("❌ ERROR: Faltan variables en tu archivo .env")
    sys.exit()


#   IMPORTS SELENIUM / API

import pandas as pd
import pytz
import re
import ssl

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import *

from slack_sdk import WebClient
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from webdriver_manager.chrome import ChromeDriverManager

#   CONFIGURACIÓN BASE PATH

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

# Carpeta temporal para credenciales
temp_folder = os.path.join(base_path, "temp_creds")
os.makedirs(temp_folder, exist_ok=True)

# Copiar el JSON a la carpeta temporal si no existe
json_original = os.path.join(base_path, GOOGLE_CREDS_FILE)
json_temp = os.path.join(temp_folder, GOOGLE_CREDS_FILE)

if not os.path.exists(json_temp):
    shutil.copy(json_original, json_temp)

creds_json_path = json_temp
print("🔐 Usando credenciales temporales en:", creds_json_path)

#   FUNCIONES AUXILIARES

def click_button(driver, selector, by=By.CSS_SELECTOR, wait_time=10):
    try:
        button = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((by, selector))
        )
        button.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en '{selector}': {e}")
        return False


def normalizar_fecha_para_selenium(fecha_str):
    s = str(fecha_str).strip()
    if not s:
        return s

    try:
        dt = datetime.strptime(s, "%d/%m/%Y")
        return dt.strftime("%m/%d/%Y")
    except:
        pass

    formatos_entrada = [
        ("%d-%m-%Y", "%m/%d/%Y"),
        ("%Y-%m-%d", "%m/%d/%Y"),
        ("%m/%d/%Y", "%m/%d/%Y"),
    ]

    for f_in, f_out in formatos_entrada:
        try:
            dt = datetime.strptime(s, f_in)
            return dt.strftime(f_out)
        except:
            continue

    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%m/%d/%Y")
    except:
        pass

    print(f"⚠️ No se pudo convertir fecha '{s}'")
    return s


def validar_y_completar_fecha(fecha_str):
    fecha_str = str(fecha_str).strip()
    if not fecha_str:
        return None

    partes = fecha_str.split("/")
    if len(partes) == 2:
        dia, mes = partes
        return f"{dia}/{mes}/{datetime.now().year}"

    if len(partes) == 3:
        d, m, a = partes
        if len(a) == 2:
            return f"{d}/{m}/20{a}"
        return fecha_str

    return None


def escribir_fecha_entrega(driver, fecha_str, intentos):
    fecha_validada = validar_y_completar_fecha(fecha_str)
    if not fecha_validada:
        print(f"❌ Fecha inválida: {fecha_str}")
        return False

    try:
        input_fecha = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "(//input[@placeholder='MM/DD/YYYY'])[3]"))
        )
        input_fecha.clear()
        fecha_normalizada = normalizar_fecha_para_selenium(fecha_validada)
        input_fecha.send_keys(fecha_normalizada)
        return True

    except Exception as e:
        print(f"❌ Error escribiendo fecha: {e}")
        return False


def escribir_punto_de_entrega(driver, texto, intentos):
    try:
        input_campo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "deliveryPoint"))
        )
        input_campo.clear()
        input_campo.send_keys(texto)
        input_campo.send_keys(Keys.ENTER)
        opcion = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//li[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{texto.lower()}')]"))
        )
        opcion.click()
        return True
    except:
        return False


def obtener_pedidos_planilla(spreadsheet_id, sheet, creds_json_path,
                             country_col=0, fecha_col=1, punto_col=2, motivo_col=3,
                             header_rows=1):

    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_json_path, scope)
    client = gspread.authorize(creds)

    ws = client.open_by_key(spreadsheet_id).worksheet(sheet)

    data = ws.get_all_values()[header_rows:]

    pedidos = []

    for row in data:
        if len(row) < 3:
            continue

        pais = row[country_col].strip()
        fecha = row[fecha_col].strip()
        punto = row[punto_col].strip()
        motivo = row[motivo_col].strip() if len(row) > motivo_col else ""

        if not pais or not fecha or not punto:
            continue

        pedidos.append((pais, fecha, punto, motivo))

    return pedidos


def guardar_cambios(driver):
    try:
        btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Guardar cambios')]"))
        )
        btn.click()
        return True
    except:
        return False

#   INICIO DEL SCRIPT

options = Options()
options.add_argument("--headless=new")
options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(
    service=ChromeService(ChromeDriverManager().install()),
    options=options
)

# Login
driver.get("https://backoffice.nilus.co/es-AR/login")
time.sleep(3)

driver.find_element(By.ID, "email").send_keys(EMAIL_LOGIN)
driver.find_element(By.ID, "password").send_keys(PASSWORD_LOGIN)
click_button(driver, "//button[text()='INGRESAR']", By.XPATH)

time.sleep(5)

# === Leer pedidos ===
pedidos = obtener_pedidos_planilla(
    SPREADSHEET_ID,
    SHEET_NAME,
    creds_json_path
)

country_map = {"AR": "1", "MX": "2"}

start_time = time.time()
MAX_DURATION = 60 * 60  # 1 hora

for pais, fecha, punto, motivo in pedidos:

    if time.time() - start_time > MAX_DURATION:
        print("⏰ Fin por timeout.")
        break

    codigo = country_map.get(pais.upper(), "1")
    url = f"https://backoffice.nilus.co/es-AR/orders?country={codigo}&status=confirmed"
    driver.get(url)

    if not escribir_punto_de_entrega(driver, punto, 0):
        continue
    if not escribir_fecha_entrega(driver, fecha, 0):
        continue

    time.sleep(2)

    while True:
        botones = driver.find_elements(By.XPATH, "//button[contains(., 'Detalle')]")
        if not botones:
            print("✔ No quedan pedidos.")
            break

        boton = botones[0]
        boton.click()
        time.sleep(2)

        # Cambiar estado → Cancelado
        try:
            combo = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "status"))
            )
            combo.click()

            opt_cancelado = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//li[contains(., 'Cancelado')]"))
            )
            opt_cancelado.click()

            combo_motivo = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "reason_of_canceled"))
            )
            combo_motivo.click()

            opt_motivo = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{motivo.lower()}')]"))
            )
            opt_motivo.click()

        except:
            print("⚠ No se pudo seleccionar motivo, usando default")
            try:
                df = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//li[contains(., 'Error en la entrega')]"))
                )
                df.click()
            except:
                pass

        guardar_cambios(driver)
        time.sleep(1)
        driver.back()
        time.sleep(2)

print("🏁 Script finalizado")
driver.quit()
