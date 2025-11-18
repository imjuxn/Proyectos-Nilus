import os
from selenium.webdriver.chrome.options import Options
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
from selenium.webdriver.common.keys import Keys
import time
import re
import pytz
import ssl
import json
from dotenv import load_dotenv

load_dotenv()

SLACK_TOKEN = os.getenv("SLACK_TOKEN")
CHANNEL_ID = os.getenv("CHANNEL_ID")

def obtener_ruta_certificado():
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, "certificados", "cacert.pem")

ssl_context = ssl.create_default_context(cafile=obtener_ruta_certificado())
client = WebClient(token=SLACK_TOKEN, ssl=ssl_context)

CHANNEL_ID_NOTIFICACIONES = os.getenv("CHANNEL_ID_NOTIFICACIONES")

def enviar_notificacion_slack(mensaje):
    try:       
        response = client.chat_postMessage(channel=CHANNEL_ID_NOTIFICACIONES, text=mensaje) 
        if not response["ok"]:
            print(f"❌ Error enviando mensaje a Slack: {response['error']}")
    except Exception as e:
        print(f"❌ Excepción enviando mensaje a Slack: {e}")

enviar_notificacion_slack(mensaje=f"El script de entrega fallida de ARGENTINA ha comenzado 🚀")

try:
    with open("ids_procesados.txt", "r") as f:
        ids_procesados = set(line.strip() for line in f)
except FileNotFoundError:
    ids_procesados = set()

def obtener_ids_desde_archivos():
    ahora = datetime.now(pytz.timezone("America/Argentina/Buenos_Aires"))
    inicio_dia = datetime(ahora.year, ahora.month, ahora.day, tzinfo=ahora.tzinfo)
    ts_inicio = inicio_dia.timestamp()

    response = client.conversations_history(channel=CHANNEL_ID, oldest=ts_inicio)
    mensajes = response["messages"]

    ids_encontrados = []

    for mensaje in mensajes:
        if "files" in mensaje:
            for archivo in mensaje["files"]:
                nombre_archivo = archivo.get("name", "")
                print(f"📎 Nombre de archivo encontrado: {nombre_archivo}")

                matches = re.findall(r"-\s*(\w+)", nombre_archivo)
                if matches:
                    ids_encontrados.append(matches[-1])

    return ids_encontrados

if __name__ == "__main__":
    ids = obtener_ids_desde_archivos()
    print("✅ IDs extraídos:")
    for id_ in ids:
        print(id_)

with open("ids_procesados.txt", "a") as f:
    f.write(id_ + "\n")

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

def escribir_punto_de_entrega(driver, texto, intentos):
    try:
        print(f"📝 Escribiendo '{texto}' en el campo deliveryPoint")

        input_campo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "deliveryPoint"))
        )
        
        input_campo.clear()
        input_campo.send_keys(texto)
        input_campo.send_keys(Keys.ENTER)
        time.sleep(1)

        opcion = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//li[contains(., '{texto}')]"))
        )
        opcion.click()

        time.sleep(2)
        print("✅ Punto de entrega seleccionado correctamente.")
        return True
    except Exception as e:
        print(f"❌ Error seleccionando punto de entrega: {e}")
        if intentos < 2:
            print("🔄 Reintentando...")
            intentos += 1
            time.sleep(2)
            escribir_punto_de_entrega(driver, texto, intentos)
        else:
            print("❌ No se pudo seleccionar el punto de entrega tras varios intentos.")   
            return False

def escribir_fecha_entrega(driver, fecha_str, intentos):
    try:
        input_fecha = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "(//input[@placeholder='MM/DD/YYYY'])[3]"))
        )

        input_fecha.click()
        time.sleep(0.5)
        input_fecha.clear()
        input_fecha.send_keys(Keys.HOME)
        input_fecha.send_keys(fecha_str)
        print(f"✅ Fecha '{fecha_str}' escrita correctamente.")
        return True
    except Exception as e:
        print(f"❌ Error al escribir fecha de entrega: {e}")
        if intentos < 2:
            print("🔄 Reintentando...")
            intentos += 1
            time.sleep(2)
            escribir_fecha_entrega(driver, fecha_str, intentos)
        else:
            print("❌ No se pudo escribir la fecha de entrega tras varios intentos.")
            return False

def guardar_cambios(driver, wait_time=10):
    try:
        modal = WebDriverWait(driver, wait_time).until(
            EC.visibility_of_element_located((By.XPATH, "//h2[contains(text(), 'Cambiar el estado del pedido')]/ancestor::div[@role='dialog']"))
        )

        guardar_btn = modal.find_element(By.XPATH, ".//button[normalize-space(text())='Guardar cambios']")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", guardar_btn)
        driver.execute_script("arguments[0].click();", guardar_btn)

        print("✅ 'Guardar cambios' dentro del modal fue clickeado correctamente.")
        return True

    except Exception as e:
        print(f"❌ No se pudo hacer clic en el botón correcto: {e}")
        return False

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

ids_slack = obtener_ids_desde_archivos()
ids_a_procesar = [id_ for id_ in ids_slack if id_ not in ids_procesados]

start_time = time.time()
MAX_DURATION = 60 * 60

for id_ in ids:
        driver.get("https://backoffice.nilus.co/es-AR/orders?country=1&status=confirmed")
        time.sleep(2)
        exito_entrega = escribir_punto_de_entrega(driver, id_, 0)
        if not exito_entrega:
            print(f"❌ No se pudo colocar el punto de entrega para el ID: {id_}. Continuando con el siguiente ID...")
            continue
        time.sleep(3)
        fecha_hoy = datetime.now().strftime("%m/%d/%Y")
        exito_fecha = escribir_fecha_entrega(driver, fecha_hoy, 0)
        if not exito_fecha:
            print(f"❌ No se pudo escribir la fecha de entrega para el ID: {id_}. Continuando con el siguiente ID...")
            continue
        time.sleep(2)
        
        while True:
                if time.time() - start_time > MAX_DURATION:
                    print("⏰ Tiempo máximo alcanzado. Finalizando script...")
                    enviar_notificacion_slack(mensaje=f"El script ha finalizado por tiempo máximo ARGENTINA.")
                    driver.quit()
                    sys.exit()

                driver.execute_script("document.querySelector('div.MuiDataGrid-virtualScroller').scrollLeft = 10000")
                time.sleep(0.5)

                scroll_container = driver.find_element(By.CSS_SELECTOR, "div.MuiDataGrid-virtualScroller")
                botones_detalle = []
                for y in range(0, 10000, 500):
                    driver.execute_script("arguments[0].scrollTop = arguments[1];", scroll_container, y)
                    time.sleep(0.5)
                    botones_detalle = driver.find_elements(By.XPATH, "//button[contains(., 'Detalle')]")
                    if botones_detalle:
                        break

                if not botones_detalle:
                    print("✅ No quedan más pedidos para procesar.")
                    break

                boton = botones_detalle[0]
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton)
                    time.sleep(0.7)
                    boton.click()
                    print("✅ Click en Detalle realizado")
                    time.sleep(2)

                    try:
                        cambiar_estado = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='email']"))
                    )
                        cambiar_estado.click()
                        print("✅ Click en combobox con id='email'")
                    except Exception:
                        try:
                            cambiar_estado = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='status']"))
                            )
                            cambiar_estado.click()
                            print("✅ Click en combobox con id='status'")
                        except Exception as e:
                            print(f"❌ No se pudo hacer clic en ningún combobox: {e}")
                        
                    cancelado_opcion = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Cancelado')]"))
                    )
                    cancelado_opcion.click()
                    motivo_combo = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "reason_of_canceled"))
                    )
                    motivo_combo.click()
                    Entrega_fallida = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Error en la entrega')]"))
                    )
                    Entrega_fallida.click()
                    guardar_cambios(driver)
                    time.sleep(2)

                    driver.back()
                    time.sleep(2)
                except Exception as e:
                    print(f"❌ No se pudo procesar el pedido: {e}")
                    try:
                        driver.back()
                        time.sleep(2)
                    except Exception as e2:
                        print(f"❌ Error al volver atrás: {e2}")
                    continue
                with open("ids_procesados.txt", "a") as f:
                    f.write(id_ + "\n")
                ids_procesados.add(id_)
                time.sleep(5)
print("✅ Script finalizado correctamente.")
driver.quit()
sys.exit()
