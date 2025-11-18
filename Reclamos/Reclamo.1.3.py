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
import gspread
import time
import requests
import re
import sys
import os
from dotenv import load_dotenv

load_dotenv()

SLACK_WEBHOOK_URL = os.getenv("SLACK_WEBHOOK_URL_RECLAMOS")
EMAIL_LOGIN = os.getenv("EMAIL_LOGIN")
PASSWORD_LOGIN = os.getenv("PASSWORD_LOGIN")
SHEET_ID = os.getenv("SHEET_ID_RECLAMOS")
GOOGLE_JSON = os.getenv("GOOGLE_JSON")
SHEET_NAME = os.getenv("SHEET_NAME_RECLAMOS")


def enviar_notificacion_slack(mensaje):
    payload = {"text": mensaje}
    try:
        requests.post(SLACK_WEBHOOK_URL, json=payload)
    except:
        pass


enviar_notificacion_slack("🚀 El script de RECLAMOS ha comenzado.")


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
    except:
        return False


def procesar_pedido(driver, url, producto_buscado, cantidad_deseada, estado):
    from difflib import SequenceMatcher

    def similitud(nombre1, nombre2):
        return SequenceMatcher(None, nombre1.lower(), nombre2.lower()).ratio()

    pedido_url = f"https://backoffice.nilus.co/es-AR/orders/{url}"
    driver.get(pedido_url)
    time.sleep(5)

    try:
        productos = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "css-1es96wk"))
        )
        mejor_match = None
        mayor_similitud = 0

        for producto in productos:
            nombre = producto.find_element(By.TAG_NAME, "span").text.strip()
            score = similitud(producto_buscado, nombre)
            if score > mayor_similitud:
                mayor_similitud = score
                mejor_match = (producto, nombre)

        if not (mejor_match and mayor_similitud >= 0.8):
            mejor_match = None
            for producto in productos:
                nombre = producto.find_element(By.TAG_NAME, "span").text.strip()
                if coincidencia_parcial(producto_buscado, nombre):
                    mejor_match = (producto, nombre)
                    break

        if not mejor_match:
            enviar_notificacion_slack(f"❌ Producto '{producto_buscado}' no encontrado en pedido {url}.")
            return

        producto, nombre = mejor_match
        boton_svg = producto.find_element(By.CSS_SELECTOR, 'svg[data-testid="DoNotDisturbOnIcon"]')
        try:
            boton_svg.click()
        except:
            driver.execute_script("arguments[0].click();", boton_svg)

        estado_norm = (str(estado) if estado is not None else "").strip().lower()

        if re.search(r"mal[\s/-]?estado", estado_norm):
            motivo_texto = "Support - DTC - Delivery Point - Product in bad condition"
        elif re.search(r"faltante", estado_norm):
            motivo_texto = "Support - DTC - Delivery Point - Missing Product"
        else:
            motivo_texto = "Support - DTC - Delivery Point - Missing Product"

        motivo_label = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//label[contains(text(), "Selecciona un motivo")]'))
        )
        combo_div = motivo_label.find_element(By.XPATH, './ancestor::div[contains(@class, "MuiFormControl-root")]')
        combo_clickable = combo_div.find_element(By.XPATH, './/div[@role="button" or @role="combobox"]')
        combo_clickable.click()

        motivo_lower = motivo_texto.lower()
        opcion_xpath = f"//li[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), \"{motivo_lower}\")]"
        opcion = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, opcion_xpath)))
        opcion.click()

        input_cantidad = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "quantity")))
        input_cantidad.clear()
        input_cantidad.send_keys(str(cantidad_deseada))

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Solicitar']"))
        ).click()

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Guardar cambios']"))
        ).click()

    except Exception as e:
        enviar_notificacion_slack(f"❌ Error procesando el pedido {url}: {str(e)}")


def procesar_producto_en_pedido(driver, producto_buscado, cantidad_deseada, estado):
    from difflib import SequenceMatcher

    def similitud(nombre1, nombre2):
        return SequenceMatcher(None, nombre1.lower(), nombre2.lower()).ratio()

    try:
        productos = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "css-1es96wk"))
        )
        mejor_match = None
        mayor_similitud = 0

        for producto in productos:
            nombre = producto.find_element(By.TAG_NAME, "span").text.strip()
            score = similitud(producto_buscado, nombre)
            if score > mayor_similitud:
                mayor_similitud = score
                mejor_match = (producto, nombre)

        if not (mejor_match and mayor_similitud >= 0.8):
            mejor_match = None
            for producto in productos:
                nombre = producto.find_element(By.TAG_NAME, "span").text.strip()
                if coincidencia_parcial(producto_buscado, nombre):
                    mejor_match = (producto, nombre)
                    break

        if not mejor_match:
            enviar_notificacion_slack(f"❌ Producto '{producto_buscado}' no encontrado en pedido.")
            return

        producto, nombre = mejor_match
        boton_svg = producto.find_element(By.CSS_SELECTOR, 'svg[data-testid="DoNotDisturbOnIcon"]')
        try:
            boton_svg.click()
        except:
            driver.execute_script("arguments[0].click();", boton_svg)

        estado_norm = (str(estado) if estado is not None else "").strip().lower()

        if re.search(r"mal[\s/-]?estado", estado_norm):
            motivo_texto = "Support - DTC - Delivery Point - Product in bad condition"
        else:
            motivo_texto = "Support - DTC - Delivery Point - Missing Product"

        motivo_label = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//label[contains(text(), "Selecciona un motivo")]'))
        )
        combo_div = motivo_label.find_element(By.XPATH, './ancestor::div[contains(@class, "MuiFormControl-root")]')
        combo_clickable = combo_div.find_element(By.XPATH, './/div[@role="button" or @role="combobox"]')
        combo_clickable.click()

        motivo_lower = motivo_texto.lower()
        opcion_xpath = f"//li[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), \"{motivo_lower}\")]"
        opcion = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, opcion_xpath)))
        opcion.click()

        input_cantidad = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "quantity")))
        input_cantidad.clear()
        input_cantidad.send_keys(str(cantidad_deseada))

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Solicitar']"))
        ).click()

    except Exception as e:
        enviar_notificacion_slack(f"❌ Error procesando producto '{producto_buscado}': {str(e)}")


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

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

RUTA_CREDENCIALES = os.path.join(base_path, GOOGLE_JSON)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciales = ServiceAccountCredentials.from_json_keyfile_name(RUTA_CREDENCIALES, scope)
cliente = gspread.authorize(credenciales)

sheet = cliente.open_by_key(SHEET_ID)
worksheet = sheet.worksheet(SHEET_NAME)
valores = worksheet.get_all_values()

filas = valores[1:]
datos_limpios = []

for fila in filas:
    if len(fila) >= 15 and fila[0].strip():
        datos_limpios.append([
            fila[0].strip(),
            fila[8].strip(),
            fila[15].strip(),
            fila[9].strip(),
            fila[10].strip() if len(fila) > 15 else "0"
        ])

df = pd.DataFrame(datos_limpios, columns=["fecha", "Estado", "url", "Producto_Reclamado", "Cantidad"])

df["Estado"] = df["Estado"].fillna("").astype(str)
df["Estado"] = df["Estado"].str.strip()

df.dropna(subset=["fecha"], inplace=True)

df["fecha"] = pd.to_datetime(df["fecha"], dayfirst=True, errors="coerce")

df_hoy = df[df["fecha"].dt.date == datetime.now().date()]

df_hoy = df_hoy[
    df_hoy["Estado"].str.lower().str.contains("faltante|mal estado", na=False) |
    df_hoy["Estado"].isna() |
    (df_hoy["Estado"].str.strip() == "")
]

grupos = df_hoy.groupby("url")

for url, grupo in grupos:
    try:
        pedido_url = f"https://backoffice.nilus.co/es-AR/orders/{url}"
        driver.get(pedido_url)
        time.sleep(5)

        for _, row in grupo.iterrows():
            producto_reclamado = str(row["Producto_Reclamado"]).strip()
            estado = row["Estado"]
            try:
                modificada = int(row["Cantidad"]) if str(row["Cantidad"]).strip().isdigit() else 0
            except:
                modificada = 0
            cantidad_deseada = modificada

            procesar_producto_en_pedido(driver, producto_reclamado, cantidad_deseada, estado)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Guardar cambios']"))
        ).click()

    except Exception as e:
        pass

time.sleep(5)
driver.quit()
