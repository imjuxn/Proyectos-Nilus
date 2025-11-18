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

# Cargar .env
load_dotenv()

SLACK_WEBHOOK_URL = os.getenv("SLACK_WEBHOOK_URL")
EMAIL_LOGIN = os.getenv("EMAIL_LOGIN")
PASSWORD_LOGIN = os.getenv("PASSWORD_LOGIN")

SHEET_ID = os.getenv("SHEET_ID")
GOOGLE_JSON = os.getenv("GOOGLE_JSON")
SHEET_NAME = os.getenv("SHEET_NAME")

def enviar_notificacion_slack(mensaje):
    payload = {"text": mensaje}
    try:
        response = requests.post(SLACK_WEBHOOK_URL, json=payload)
        if response.status_code != 200:
            print(f"❌ Error enviando Slack: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"❌ Excepción Slack: {e}")

# Notificar inicio
enviar_notificacion_slack("🇦🇷 OPERACIONES ejecuto los desvios en ARGENTINA 🚀.")

def normalizar(texto):
    texto = re.sub(r"[^a-zA-Z0-9áéíóúñü\s]", "", texto.lower())
    return set(texto.split())

def coincidencia_parcial(nombre_excel, nombre_web, umbral=0.5):
    palabras_excel = normalizar(nombre_excel)
    palabras_web = normalizar(nombre_web)
    if not palabras_excel:
        return False
    comunes = palabras_excel & palabras_web
    return len(comunes) / len(palabras_excel) >= umbral

def click_button(driver, selector, by=By.CSS_SELECTOR, wait_time=10):
    try:
        button = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((by, selector))
        )
        button.click()
        return True
    except Exception as e:
        print(f"❌ Error click '{selector}': {e}")
        return False

def sincronizar_pedido(driver):
    try:
        try:
            boton_sync = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "syncOrdder"))
            )
        except TimeoutException:
            boton_sync = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Sincronizar pedido']"))
            )

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton_sync)
        boton_sync.click()
        print("✅ Pedido sincronizado.")
    except Exception as e:
        print(f"❌ Error sincronizar: {e}")

def procesar_pedido(driver, datos_pedido, producto_buscado, cantidad_deseada,
                    cantidad_original, tipo_desvio, producto_nuevo=None):

    def similitud(a, b):
        return SequenceMatcher(None, a.lower(), b.lower()).ratio()

    pedido_url = f"https://backoffice.nilus.co/es-AR/orders/{datos_pedido}?odoo_references"
    print(f"\n🔄 Procesando pedido: {datos_pedido}")
    driver.get(pedido_url)
    time.sleep(5)

    try:
        productos = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "css-1es96wk"))
        )

        if tipo_desvio == "faltante" and len(productos) == 1:
            cancelar_pedido(driver)
            return

        mejor_match = None
        mayor_similitud = 0
        
        for producto in productos:
            nombre = producto.find_element(By.TAG_NAME, "span").text.strip()
            score = similitud(producto_buscado, nombre)
            if score > mayor_similitud:
                mejor_match = (producto, nombre)
                mayor_similitud = score

        if mayor_similitud < 0.8:
            mejor_match = None
            for producto in productos:
                nombre = producto.find_element(By.TAG_NAME, "span").text.strip()
                if coincidencia_parcial(producto_buscado, nombre):
                    mejor_match = (producto, nombre)
                    break

        if not mejor_match:
            msg = f"❌ Producto '{producto_buscado}' no encontrado en {datos_pedido}."
            print(msg)
            enviar_notificacion_slack(msg)
            return

        producto, nombre = mejor_match
        boton_svg = producto.find_element(By.CSS_SELECTOR, 'svg[data-testid="DoNotDisturbOnIcon"]')

        if tipo_desvio == "faltante_parcial":
            for _ in range(cantidad_deseada):
                boton_svg.click()
                time.sleep(0.5)

        elif tipo_desvio in ["faltante", "reemplazo"]:
            boton_del = producto.find_element(By.CSS_SELECTOR, 'svg[data-testid="DeleteForeverIcon"]')
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton_del)
            boton_del.click()
            
            boton_eliminar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Eliminar']"))
            )
            boton_eliminar.click()
            time.sleep(0.5)

        if tipo_desvio in ["faltante", "faltante_parcial"]:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Guardar cambios']"))
            ).click()
            time.sleep(2)
            sincronizar_pedido(driver)

        if tipo_desvio == "reemplazo":
            boton_agregar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Agregar Producto")]'))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton_agregar)
            boton_agregar.click()

            input_producto = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, "productSelected"))
            )
            input_producto.click()
            input_producto.clear()
            input_producto.send_keys(producto_nuevo)
            input_producto.send_keys(Keys.SPACE)
            input_producto.send_keys(Keys.BACKSPACE)

            opciones = WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "li[role='option']"))
            )
            elegido = False
            for opcion in opciones:
                if opcion.text.strip().lower().startswith(producto_nuevo.lower()):
                    opcion.click()
                    elegido = True
                    break

            if not elegido:
                msg = f"❌ Producto de reemplazo '{producto_nuevo}' no encontrado ({datos_pedido})."
                print(msg)
                enviar_notificacion_slack(msg)
                return

            boton_modal = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[@role="dialog"]//button[normalize-space()="Agregar Producto"]'))
            )
            boton_modal.click()

            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Guardar cambios']"))
            ).click()
            time.sleep(3)

            productos_actualizados = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, "css-1es96wk"))
            )
            for producto in productos_actualizados:
                nombre = producto.find_element(By.TAG_NAME, "span").text.strip().lower()
                if producto_nuevo.lower() in nombre:
                    boton_sumar = producto.find_element(By.CSS_SELECTOR, 'svg[data-testid="AddCircleIcon"]')
                    for _ in range(max(cantidad_deseada - 1, 0)):
                        boton_sumar.click()
                        WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Guardar cambios']"))
                        ).click()
                        time.sleep(1)
                        sincronizar_pedido(driver)

        print(f"✅ Pedido {datos_pedido} procesado.")

    except Exception as e:
        msg = f"❌ Error procesando pedido {datos_pedido}: {str(e)}"
        print(msg)
        enviar_notificacion_slack(msg)

def cancelar_pedido(driver):
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

        cancelar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Cancelado')]"))
        )
        cancelar.click()

        motivo_combo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "reason_of_canceled"))
        )
        motivo_combo.click()

        motivo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Productos faltantes')]"))
        )
        motivo.click()

        guardar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@role="dialog"]//button[normalize-space()="Guardar cambios"]'))
        )
        guardar.click()
        time.sleep(2)

        sincronizar_pedido(driver)
        print("✅ Pedido cancelado completamente.")

    except Exception as e:
        print(f"❌ Error cancelar pedido: {e}")

# === INICIO ===
options = Options()
options.add_argument("--headless=new")
options.add_argument("--window-size=1920,1080")
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

driver.get("https://backoffice.nilus.co/es-AR/login")
time.sleep(5)

driver.find_element(By.ID, "email").send_keys(EMAIL_LOGIN)
driver.find_element(By.ID, "password").send_keys(PASSWORD
