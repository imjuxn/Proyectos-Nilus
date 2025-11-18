import shutil
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
import locale
import time
import requests
import re
import sys
import os
from concurrent.futures import ProcessPoolExecutor, as_completed
from dotenv import load_dotenv

load_dotenv()

base_path = os.path.dirname(__file__)
temp_folder = os.path.join(base_path, "temp_creds")
os.makedirs(temp_folder, exist_ok=True)

json_original = os.path.join(base_path, os.getenv("GOOGLE_CREDENTIALS_FILE"))
json_temp = os.path.join(temp_folder, os.getenv("GOOGLE_CREDENTIALS_FILE"))

if not os.path.exists(json_temp):
    shutil.copy(json_original, json_temp)

creds_json_path = os.path.join(base_path, "temp_creds", os.getenv("GOOGLE_CREDENTIALS_FILE"))
print("Usando credenciales en:", creds_json_path)

def puede_modificar_pedido(fecha_pedido_str):
    try:
        fecha_pedido = datetime.strptime(fecha_pedido_str, "%m/%d/%Y")
        ahora = datetime.now()
        ultimo_dia_modificacion = fecha_pedido.date() - timedelta(days=2)
        hora_limite = datetime.strptime("12:00", "%H:%M").time()
        limite_modificacion = datetime.combine(ultimo_dia_modificacion, hora_limite)
        if ahora > limite_modificacion:
            pass
        if ahora <= limite_modificacion:
            return False
        else:
            diferencia = ahora - limite_modificacion
            if diferencia.days > 0:
                print(f"   ⚠️ Pasó hace {diferencia.days} día(s)")
            else:
                horas = diferencia.seconds // 3600
                minutos = (diferencia.seconds % 3600) // 60
                print(f"   ⚠️ Pasó hace {horas}h {minutos}m")
            return True
    except Exception as e:
        print(f"❌ Error al validar fecha del pedido {fecha_pedido_str}: {e}")
        return False

def puede_modificar_pedido_mexico(fecha_pedido_str):
    try:
        fecha_pedido = datetime.strptime(fecha_pedido_str, "%m/%d/%Y")
        ahora = datetime.now()
        ultimo_dia_modificacion = fecha_pedido.date() - timedelta(days=1)
        hora_limite = datetime.strptime("08:00", "%H:%M").time()
        limite_modificacion = datetime.combine(ultimo_dia_modificacion, hora_limite)
        if ahora <= limite_modificacion:
            return False
        else:
            return True
    except Exception as e:
        print(f"❌ Error al validar fecha del pedido {fecha_pedido_str}: {e}")
        return False

def obtener_fecha_pedido_desde_html(driver, wait_time=10):
    try:
        fecha_element = WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.XPATH, "//h6[contains(@class, 'MuiTypography-h6') and contains(text(), '/')]"))
        )
        fecha_completa = fecha_element.text.strip()
        fecha_solo = fecha_completa.split()[0]
        partes_fecha = fecha_solo.split('/')
        if len(partes_fecha) == 3:
            dia = partes_fecha[0].zfill(2)
            mes = partes_fecha[1].zfill(2)
            año = partes_fecha[2]
            fecha_formateada = f"{dia}/{mes}/{año}"
            return fecha_formateada
        else:
            print(f"❌ Formato de fecha no reconocido: {fecha_solo}")
            return None
    except Exception as e:
        print(f"❌ Error al obtener fecha del pedido desde HTML: {e}")
        return None

def click_button(driver, selector, by=By.CSS_SELECTOR, wait_time=10):
    try:
        button = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((by, selector))
        )
        button.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en el selector '{selector}'")
        return False

def click_edit_icon(driver, wait_time=10):
    try:
        contenedor = WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.XPATH, "//div[h6[text()='Fecha de entrega']]"))
        )
        edit_icon = contenedor.find_element(By.CSS_SELECTOR, "svg.MuiSvgIcon-root.MuiSvgIcon-colorSecondary[data-testid='EditIcon']")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", edit_icon)
        time.sleep(1)
        edit_icon.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en ícono de edición: {e}")
        return False

def escribir_slug_shop(driver, texto, wait_time=10):
    try:
        edit_icon = WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "svg.MuiSvgIcon-root.MuiSvgIcon-colorSecondary[data-testid='EditIcon']"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", edit_icon)
        time.sleep(1)
        edit_icon = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "svg.MuiSvgIcon-root.MuiSvgIcon-colorSecondary[data-testid='EditIcon']"))
        )
        edit_icon.click()
        time.sleep(1)
        input_slug = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.ID, "slugShop"))
        )
        input_slug.clear()
        input_slug.send_keys(texto)
        input_slug.send_keys(Keys.ENTER)
        time.sleep(1)
        try:
            opcion = WebDriverWait(driver, wait_time).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "[id*='slugShop-option'][id$='-0']"))
            )
            opcion.click()
            time.sleep(1)
            return True
        except Exception as e2:
            print(f"❌ Error al seleccionar opción del dropdown: {e2}")
            return False
    except Exception as e:
        print(f"❌ Error al escribir en el input slugShop: {e}")
        return False

def escribir_slug_2(driver, texto, wait_time=10):
    try:
        input_slug = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.ID, "slugShop"))
        )
        input_slug.clear()
        input_slug.send_keys(texto)
        input_slug.send_keys(Keys.ENTER)
        time.sleep(1)
        try:
            opcion = WebDriverWait(driver, wait_time).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "[id*='slugShop-option'][id$='-0']"))
            )
            opcion.click()
            time.sleep(1)
            return True
        except Exception as e2:
            print(f"❌ Error al seleccionar opción del dropdown: {e2}")
            return False
    except Exception as e:
        print(f"❌ Error al escribir en el input slugShop: {e}")
        return False

def seleccionar_tercera_fecha(driver, wait_time=10):
    try:
        dropdown_fecha = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.ID, "selected_delivery_date"))
        )
        dropdown_fecha.click()
        time.sleep(1)
        opciones = WebDriverWait(driver, wait_time).until(
            EC.presence_of_all_elements_located((By.XPATH, "//li[@role='option' and contains(@class, 'MuiMenuItem-root')]"))
        )
        if len(opciones) >= 3:
            tercera_opcion = opciones[2]
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", tercera_opcion)
            time.sleep(1)
            tercera_opcion.click()
            return True
        else:
            print("❌ No hay suficientes opciones para seleccionar la tercera.")
            return False
    except Exception as e:
        print(f"❌ Error al seleccionar la tercera fecha de entrega: {e}")
        return False

def click_edit_entreg(driver, wait_time=10):
    try:
        edit_icon = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "svg.MuiSvgIcon-root.MuiSvgIcon-colorSecondary[data-testid='EditIcon']"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", edit_icon)
        time.sleep(1)
        edit_icon.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en el ícono de edición: {e}")
        return False

def ir_a_detalle_primer_pendiente(driver, wait_time=10):
    try:
        pendiente_span = WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'MuiChip-label') and normalize-space(text())='Pendiente']"))
        )
        status_cell = pendiente_span.find_element(By.XPATH, "./ancestor::div[@data-field='status']")
        fila = status_cell.find_element(By.XPATH, "./ancestor::div[contains(@class, 'MuiDataGrid-row')]")
        orderid_cell = fila.find_element(By.XPATH, ".//div[@data-field='orderId']")
        boton_detalle = orderid_cell.find_element(By.XPATH, ".//button[normalize-space(text())='Detalle']")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", status_cell)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton_detalle)
        time.sleep(1)
        boton_detalle.click()
        return True
    except Exception as e:
        print(f"❌ No se pudo hacer clic en el botón 'Detalle' de la primera fila pendiente: {e}")
        return False

def seleccionar_fecha_entrega(driver, fecha_dd_mm_yyyy, wait_time=10):
    try:
        dropdown_fecha = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.ID, "selected_delivery_date"))
        )
        dropdown_fecha.click()
        time.sleep(1)
        if buscar_y_seleccionar_fecha(driver, fecha_dd_mm_yyyy):
            time.sleep(1)
            return True
        else:
            print("❌ No se pudo encontrar la fecha")
            return False
    except Exception as e:
        print(f"❌ Error al abrir dropdown de fecha: {e}")
        return False

def click_editar_fecha(driver, wait_time=10):
    try:
        boton_editar_fecha = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Editar Fecha']"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton_editar_fecha)
        time.sleep(1)
        boton_editar_fecha.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en el botón 'Editar Fecha': {e}")
        return False

def convertir_fecha_a_texto_espanol(fecha_str):
    try:
        try:
            locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
        except:
            try:
                locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
            except:
                pass
        fecha_obj = datetime.strptime(fecha_str, "%d/%m/%Y")
        meses = {
            1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
            5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
            9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
        }
        dias = {
            0: "lunes", 1: "martes", 2: "miércoles", 3: "jueves",
            4: "viernes", 5: "sábado", 6: "domingo"
        }
        dia_semana = dias[fecha_obj.weekday()]
        dia = fecha_obj.day
        mes = meses[fecha_obj.month]
        año = fecha_obj.year
        fecha_español = f"{dia_semana}, {dia} de {mes} de {año}"
        return fecha_español
    except Exception as e:
        print(f"❌ Error al convertir fecha: {e}")
        return None

def buscar_y_seleccionar_fecha(driver, fecha_dd_mm_yyyy, wait_time=10):
    try:
        fecha_texto = convertir_fecha_a_texto_espanol(fecha_dd_mm_yyyy)
        if not fecha_texto:
            return False
        fecha_element = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '{fecha_texto}')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", fecha_element)
        time.sleep(1)
        fecha_element.click()
        return True
    except Exception as e:
        print(f"❌ Error al buscar fecha")
        return False

def conectar_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_json_path, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(os.getenv("SPREADSHEET_ID"))
        worksheet = sheet.worksheet(os.getenv("SHEET_NAME"))
        valores = worksheet.get_all_values()
        print(f"✅ Conectado a Google Sheets. Filas encontradas: {len(valores)}")
        return valores, worksheet
    except Exception as e:
        print(f"❌ Error al conectar con Google Sheets: {e}")
        return None, None

def click_editar_punto_entrega(driver, wait_time=10):
    try:
        boton_editar = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Editar punto de entrega')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton_editar)
        boton_editar.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en 'Editar punto de entrega': {e}")
        return False

def cambiar_estado_a_confirmado(driver, wait_time=10):
    try:
        combobox_element = None
        try:
            combobox_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//div[@role='combobox' and @id='email']"))
            )
        except Exception:
            try:
                combobox_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@role='combobox' and @id='status']"))
                )
            except Exception as e:
                print(f"❌ No se pudo encontrar el combobox de estado: {e}")
                return False
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", combobox_element)
        try:
            cambiar_estado = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='email']"))
            )
            cambiar_estado.click()
        except Exception:
            try:
                cambiar_estado = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='status']"))
                )
                cambiar_estado.click()
            except Exception as e:
                print(f"❌ No se pudo hacer clic en ningún combobox:")
                return False
        try:
            confirmando_opcion = WebDriverWait(driver, wait_time).until(
                EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Confirmado')]"))
            )
            confirmando_opcion.click()
            time.sleep(0.5)
            try:
                modal = WebDriverWait(driver, wait_time).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@role='dialog' or contains(@class, 'modal') or contains(@class, 'MuiDialog')]"))
                )
                guardar_btn = modal.find_element(By.XPATH, ".//button[normalize-space(text())='Guardar cambios']")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", guardar_btn)
                driver.execute_script("arguments[0].click();", guardar_btn)
                return True
            except Exception as e:
                print(f"❌ No se pudo hacer clic en el botón 'Guardar cambios': {e}")
                return False
        except Exception as e:
            print(f"❌ Error al seleccionar 'Confirmado': {e}")
            return False
    except Exception as e:
        print(f"❌ Error general al cambiar estado: {e}")
        return False

def cambiar_estado_a_confirmado2(driver, wait_time=10):
    try:
        try:
            estado_confirmado = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'MuiChip-label') and normalize-space(text())='Confirmado']"))
            )
            print("✅ El pedido ya está CONFIRMADO. Saltando...")
            return True
        except TimeoutException:
            print("📝 El pedido no está confirmado. Procediendo a cambiar estado...")
            pass
        combobox_element = None
        try:
            combobox_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//div[@role='combobox' and @id='email']"))
            )
        except Exception:
            try:
                combobox_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@role='combobox' and @id='status']"))
                )
            except Exception as e:
                print(f"❌ No se pudo encontrar el combobox de estado: {e}")
                return False
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", combobox_element)
        try:
            cambiar_estado = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='email']"))
            )
            cambiar_estado.click()
        except Exception:
            try:
                cambiar_estado = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='status']"))
                )
                cambiar_estado.click()
            except Exception as e:
                print(f"❌ No se pudo hacer clic en ningún combobox:")
                return False
        try:
            confirmando_opcion = WebDriverWait(driver, wait_time).until(
                EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Confirmado')]"))
            )
            confirmando_opcion.click()
            time.sleep(0.5)
            try:
                modal = WebDriverWait(driver, wait_time).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@role='dialog' or contains(@class, 'modal') or contains(@class, 'MuiDialog')]"))
                )
                guardar_btn = modal.find_element(By.XPATH, ".//button[normalize-space(text())='Guardar cambios']")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", guardar_btn)
                driver.execute_script("arguments[0].click();", guardar_btn)
                print("✅ Estado cambiado a 'Confirmado' exitosamente.")
                return True
            except Exception as e:
                print(f"❌ No se pudo hacer clic en el botón 'Guardar cambios': {e}")
                return False
        except Exception as e:
            print(f"❌ Error al seleccionar 'Confirmado': {e}")
            return False
    except Exception as e:
        print(f"❌ Error general al cambiar estado: {e}")
        return False

def click_visibility_icon(driver, wait_time=10):
    try:
        visibility_icon = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "svg.MuiSvgIcon-root.MuiSvgIcon-colorSecondary[data-testid='VisibilityIcon']"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", visibility_icon)
        time.sleep(1)
        visibility_icon.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en ícono de visibilidad ")
        return False

def procesar_pedido2(driver, pedido_id, fecha_entrega, slug_shop):
    try:
        if click_edit_entreg(driver):
            time.sleep(1)
            if escribir_slug_2(driver, slug_shop):
                time.sleep(1)
                if seleccionar_fecha_entrega(driver, fecha_entrega):
                    time.sleep(1)
                    if click_editar_punto_entrega(driver):
                        time.sleep(1)
                        if cambiar_estado_a_confirmado2(driver):
                            print(f"✅ Pedido {pedido_id} procesado completamente")
                            return True
        print(f"❌ Error al procesar pedido {pedido_id}")
        return False
    except Exception as e:
        print(f"❌ Error procesando pedido {pedido_id}: {e}")
        return False

def procesar_pedido(driver, pedido_id, fecha_entrega, slug_shop):
    try:
        if click_edit_icon(driver):
            time.sleep(0.5)
            if seleccionar_tercera_fecha(driver):
                time.sleep(0.5)
                if click_editar_fecha(driver):
                    time.sleep(0.5)
                    if click_visibility_icon(driver):
                        time.sleep(0.5)
                        if ir_a_detalle_primer_pendiente(driver):
                            time.sleep(0.5)
                            if escribir_slug_shop(driver, slug_shop, wait_time=10):
                                time.sleep(0.5)
                                if seleccionar_fecha_entrega(driver, fecha_entrega):
                                    time.sleep(0.5)
                                    if click_editar_punto_entrega(driver):
                                        time.sleep(0.5)
                                        if cambiar_estado_a_confirmado(driver):
                                            time.sleep(0.5)
                                            print(f"✅ Pedido {pedido_id} procesado completamente")
                                            return True
        print(f"❌ Error al procesar pedido {pedido_id}")
        return False
    except Exception as e:
        print(f"❌ Error procesando pedido {pedido_id}")
        return False

def procesar_pedido_wrapper(args):
    pedido_id, fecha_entrega, slug_shop, pais, fecha_pedido_actual = args
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    driver_local = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    try:
        driver_local.get("https://backoffice.nilus.co/es-AR/login")
        time.sleep(5)
        driver_local.find_element(By.ID, "email").send_keys(os.getenv("NILUS_EMAIL"))
        driver_local.find_element(By.ID, "password").send_keys(os.getenv("NILUS_PASSWORD"))
        click_button(driver_local, "//button[text()='INGRESAR']", By.XPATH)
        time.sleep(8)
        url_pedido = f"https://backoffice.nilus.co/es-AR/orders/{pedido_id}"
        driver_local.get(url_pedido)
        time.sleep(3)
        resultado = procesar_pedido(driver_local, pedido_id, fecha_entrega, slug_shop)
        return (pedido_id, resultado, "Éxito" if resultado else "Falló procesar_pedido")
    except Exception as e:
        return (pedido_id, False, f"Error: {str(e)}")
    finally:
        driver_local.quit()

options = Options()
options.add_argument("--headless=new")
options.add_argument("--window-size=1920,1080")

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

if __name__ == '__main__':
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    driver.get("https://backoffice.nilus.co/es-AR/login")
    time.sleep(5)
    driver.find_element(By.ID, "email").send_keys(os.getenv("NILUS_EMAIL"))
    driver.find_element(By.ID, "password").send_keys(os.getenv("NILUS_PASSWORD"))
    click_button(driver, "//button[text()='INGRESAR']", By.XPATH)
    time.sleep(5)
    try:
        valores, worksheet = conectar_google_sheets()
        if not valores:
            print("❌ No se pudo conectar con Google Sheets")
            exit()
        pedidos_procesados = 0
        pedidos_fallidos = 0
        pedidos_para_paralelizar = []
        pedidos_secuenciales = []
        for i, fila in enumerate(valores[1:], start=2):
            if len(fila) < 5:
                print(f"❌ Fila {i}: Datos insuficientes. Saltando...")
                worksheet.update_cell(i, 5, "❌ Datos faltantes")
                continue
            pais = str(fila[0]).strip().upper() if fila[0] else ""
            pedido_id = str(fila[1]).strip() if fila[1] else ""
            fecha_entrega = str(fila[2]).strip() if fila[2] else ""
            slug_shop = str(fila[3]).strip() if fila[3] else ""
            if not pais or not pedido_id or not fecha_entrega or not slug_shop:
                print(f"❌ Fila {i}: Datos faltantes. Saltando...")
                worksheet.update_cell(i, 5, "❌ Datos faltantes")
                continue
            if pais not in ["AR", "MX"]:
                print(f"❌ Fila {i}: País inválido '{pais}'. Saltando...")
                worksheet.update_cell(i, 5, "❌ País inválido")
                continue
            url_pedido = f"https://backoffice.nilus.co/es-AR/orders/{pedido_id}"
            driver.get(url_pedido)
            time.sleep(5)
            fecha_pedido_actual = obtener_fecha_pedido_desde_html(driver)
            if not fecha_pedido_actual:
                print(f"❌ No se pudo obtener la fecha del pedido {pedido_id}")
                worksheet.update_cell(i, 5, "❌ No se pudo obtener fecha")
                continue
            puede_modificar = False
            if pais == "AR":
                puede_modificar = puede_modificar_pedido(fecha_pedido_actual)
            elif pais == "MX":
                puede_modificar = puede_modificar_pedido_mexico(fecha_pedido_actual)
            if puede_modificar:
                pedidos_para_paralelizar.append((pedido_id, fecha_entrega, slug_shop, pais, fecha_pedido_actual, i))
            else:
                pedidos_secuenciales.append((pedido_id, fecha_entrega, slug_shop, pais, i))
        print(f"📦 Pedidos para paralelizar (procesar_pedido): {len(pedidos_para_paralelizar)}")
        print(f"📦 Pedidos secuenciales (procesar_pedido2): {len(pedidos_secuenciales)}")
        if pedidos_para_paralelizar:
            print("\n🚀 Iniciando procesamiento en PARALELO...")
            max_workers = 3
            with ProcessPoolExecutor(max_workers=max_workers) as executor:
                future_to_info = {}
                for pedido_id, fecha_entrega, slug_shop, pais, fecha_pedido_actual, fila_num in pedidos_para_paralelizar:
                    future = executor.submit(procesar_pedido_wrapper, (pedido_id, fecha_entrega, slug_shop, pais, fecha_pedido_actual))
                    future_to_info[future] = fila_num
                for future in as_completed(future_to_info):
                    fila_num = future_to_info[future]
                    try:
                        pedido_id, exito, mensaje = future.result()
                        if exito:
                            pedidos_procesados += 1
                            print(f"✅ Pedido {pedido_id} (fila {fila_num}) procesado en paralelo")
                            worksheet.update_cell(fila_num, 5, "✅ Hecho")
                        else:
                            pedidos_fallidos += 1
                            print(f"❌ Pedido {pedido_id} (fila {fila_num}) falló: {mensaje}")
                            worksheet.update_cell(fila_num, 5, f"❌ {mensaje}")
                    except Exception as e:
                        pedidos_fallidos += 1
                        print(f"❌ Error en fila {fila_num}: {e}")
                        worksheet.update_cell(fila_num, 5, f"❌ Error")
        if pedidos_secuenciales:
            print(f"\n⏳ Iniciando procesamiento SECUENCIAL de {len(pedidos_secuenciales)} pedidos...")
            for pedido_id, fecha_entrega, slug_shop, pais, fila_num in pedidos_secuenciales:
                print(f"\n🔄 Procesando secuencialmente: {pedido_id}")
                url_pedido = f"https://backoffice.nilus.co/es-AR/orders/{pedido_id}"
                driver.get(url_pedido)
                time.sleep(3)
                if procesar_pedido2(driver, pedido_id, fecha_entrega, slug_shop):
                    pedidos_procesados += 1
                    print(f"✅ Fila {fila_num} procesada correctamente (secuencial)")
                    worksheet.update_cell(fila_num, 5, "✅ Hecho")
                else:
                    pedidos_fallidos += 1
                    print(f"❌ Fila {fila_num} falló (secuencial)")
                    worksheet.update_cell(fila_num, 5, "❌ Fallo")
        print(f"\n📊 RESUMEN FINAL:")
        print(f"✅ Pedidos procesados correctamente: {pedidos_procesados}")
        print(f"❌ Pedidos fallidos: {pedidos_fallidos}")
        print(f"📈 Total procesados: {pedidos_procesados + pedidos_fallidos}")
    except Exception as e:
        print(f"❌ Error general en el script: {e}")
    finally:
        try:
            driver.quit()
        except:
            pass
    print("✅ Script finalizado.")
