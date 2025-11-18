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
from selenium.webdriver.common.keys import Keys
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
import pytz
import ssl
import shutil
import json
from dotenv import load_dotenv

load_dotenv()

SLACK_TOKEN = os.getenv("SLACK_TOKEN")
CHANNEL_ID_NOTIFICACIONES = os.getenv("CHANNEL_ID_NOTIFICACIONES")
EMAIL_LOGIN = os.getenv("EMAIL_LOGIN")
PASSWORD_LOGIN = os.getenv("PASSWORD_LOGIN")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME")
GOOGLE_JSON = os.getenv("GOOGLE_JSON")

client = WebClient(token=SLACK_TOKEN)

def enviar_notificacion_slack(mensaje):
    try:
        response = client.chat_postMessage(channel=CHANNEL_ID_NOTIFICACIONES, text=mensaje)
    except Exception as e:
        print(e)

enviar_notificacion_slack("El script de entrega fallida de ARG Y MX ha comenzado una vez por día 🚀")

base_path = os.path.dirname(__file__)
temp_folder = os.path.join(base_path, "temp_creds")
os.makedirs(temp_folder, exist_ok=True)

json_original = os.path.join(base_path, GOOGLE_JSON)
json_temp = os.path.join(temp_folder, GOOGLE_JSON)

if not os.path.exists(json_temp):
    shutil.copy(json_original, json_temp)

creds_json_path = os.path.join(base_path, "temp_creds", GOOGLE_JSON)

def click_button(driver, selector, by=By.CSS_SELECTOR, wait_time=10):
    try:
        btn = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((by, selector))
        )
        btn.click()
        return True
    except:
        return False

def normalizar_fecha_para_selenium(fecha_str):
    s = str(fecha_str).strip()
    if not s:
        return s
    formatos = ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y")
    for fmt in formatos:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%m/%d/%Y")
        except:
            continue
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
        if not pd.isna(dt):
            return dt.strftime("%m/%d/%Y")
    except:
        pass
    return s

def escribir_punto_de_entrega(driver, texto, intentos):
    try:
        campo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "deliveryPoint"))
        )
        campo.clear()
        campo.send_keys(texto)
        campo.send_keys(Keys.ENTER)
        opcion = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//li[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{texto.lower()}')]"))
        )
        opcion.click()
        return True
    except:
        if intentos < 2:
            return escribir_punto_de_entrega(driver, texto, intentos + 1)
        return False

def validar_y_completar_fecha(fecha_str):
    fecha_str = str(fecha_str).strip()
    if not fecha_str:
        return None
    partes = fecha_str.split("/")
    if len(partes) == 2:
        d, m = partes
        return f"{d}/{m}/{datetime.now().year}"
    if len(partes) == 3:
        d, m, a = partes
        if len(a) == 2:
            return f"{d}/{m}/20{a}"
        return fecha_str
    return None

def escribir_fecha_entrega(driver, fecha_str, intentos):
    fecha_validada = validar_y_completar_fecha(fecha_str)
    if not fecha_validada:
        return False
    try:
        inp = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "(//input[@placeholder='MM/DD/YYYY'])[3]"))
        )
        inp.clear()
        fecha_final = normalizar_fecha_para_selenium(fecha_validada)
        inp.send_keys(fecha_final)
        return True
    except:
        if intentos < 2:
            return escribir_fecha_entrega(driver, fecha_validada, intentos + 1)
        return False

def guardar_cambios(driver, wait_time=10):
    try:
        modal = WebDriverWait(driver, wait_time).until(
            EC.visibility_of_element_located((By.XPATH, "//h2[contains(text(), 'Cambiar el estado del pedido')]/ancestor::div[@role='dialog']"))
        )
        boton = modal.find_element(By.XPATH, ".//button[normalize-space(text())='Guardar cambios']")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton)
        driver.execute_script("arguments[0].click();", boton)
        return True
    except:
        return False

def obtener_pedidos_planilla(spreadsheet_id, sheet_name, creds_json_path,
                             country_col=0, fecha_col=1, punto_col=2, header_rows=1):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_json_path, scope)
    client = gspread.authorize(creds)
    sh = client.open_by_key(spreadsheet_id)
    ws = sh.worksheet(sheet_name)
    data = ws.get_all_values()[header_rows:]
    pedidos = []
    for row in data:
        if len(row) <= max(country_col, fecha_col, punto_col):
            continue
        pais = row[country_col].strip()
        fecha = row[fecha_col].strip()
        punto = row[punto_col].strip()
        if not pais or not fecha or not punto:
            continue
        pedidos.append((pais, fecha, punto))
    return pedidos

options = Options()
options.add_argument("--headless=new")
options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

driver.get("https://backoffice.nilus.co/es-AR/login")
time.sleep(5)

driver.find_element(By.ID, "email").send_keys(EMAIL_LOGIN)
driver.find_element(By.ID, "password").send_keys(PASSWORD_LOGIN)
click_button(driver, "//button[text()='INGRESAR']", By.XPATH)
time.sleep(5)

pedidos_planilla = obtener_pedidos_planilla(
    SPREADSHEET_ID,
    SHEET_NAME,
    creds_json_path,
    country_col=0,
    fecha_col=1,
    punto_col=2,
    header_rows=1
)

country_map = {"AR": "1", "MX": "2"}

start_time = time.time()
MAX_DURATION = 60 * 60
MOTIVO_FIJO = "Error en la entrega"

for pais, fecha, punto in pedidos_planilla:
    codigo = country_map.get(pais.upper(), "1")
    url = f"https://backoffice.nilus.co/es-AR/orders?country={codigo}&status=confirmed"
    driver.get(url)
    time.sleep(2)

    if not escribir_punto_de_entrega(driver, punto, 0):
        continue

    fecha_norm = normalizar_fecha_para_selenium(fecha)
    if not escribir_fecha_entrega(driver, fecha_norm, 0):
        continue

    time.sleep(2)

    while True:
        if time.time() - start_time > MAX_DURATION:
            driver.quit()
            sys.exit()

        botones = driver.find_elements(By.XPATH, "//button[contains(., 'Detalle')]")
        if not botones:
            break

        boton = botones[0]
        boton.click()
        time.sleep(2)

        try:
            try:
                combo = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='email']"))
                )
                combo.click()
            except:
                combo = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='status']"))
                )
                combo.click()

            cancelado = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Cancelado')]"))
            )
            cancelado.click()

            motivo_combo = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "reason_of_canceled"))
            )
            motivo_combo.click()

            motivo_elem = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[@role='option' and contains(text(), '{MOTIVO_FIJO}')]"))
            )
            motivo_elem.click()

            guardar_cambios(driver)
            time.sleep(2)
            driver.back()
            time.sleep(2)

        except Exception:
            try:
                driver.back()
                time.sleep(2)
            except:
                pass
            continue

print("Script finalizado correctamente.")
driver.quit()
sys.exit()
