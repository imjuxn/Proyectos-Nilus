import os
from dotenv import load_dotenv
load_dotenv()

from selenium.webdriver.chrome.options import Options
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
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import ssl
import time
import gspread
import sys
import tempfile
import json

SLACK_TOKEN = os.getenv("SLACK_TOKEN")
CHANNEL_ID = os.getenv("CHANNEL_ID")

def obtener_ruta_certificado():
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, "certificados", "cacert.pem")

ssl_context = ssl.create_default_context(cafile=obtener_ruta_certificado())
client = WebClient(token=SLACK_TOKEN, ssl=ssl_context)
CHANNEL_ID_NOTIFICACIONES = CHANNEL_ID

def enviar_notificacion_slack(mensaje):
    try:
        response = client.chat_postMessage(channel=CHANNEL_ID_NOTIFICACIONES, text=mensaje)
        if not response["ok"]:
            print(f"❌ Error enviando mensaje a Slack: {response['error']}")
    except Exception as e:
        print(f"❌ Excepción enviando mensaje a Slack: {e}")

enviar_notificacion_slack(mensaje=f"Inicio anulación de SIMPLIROUTE ARG 🇦🇷")

def marcar_estado_pedido(worksheet, fila_numero, estado):
    try:
        columna = 10
        if estado == 'procesado':
            marca = "✅"
        elif estado == 'fallido':
            marca = "❌"
        else:
            marca = "?"
        worksheet.update_cell(fila_numero, columna + 1, marca)
        print(f"   📝 Fila {fila_numero}: Marcado como {marca}")
    except Exception as e:
        print(f"❌ Error al marcar estado en fila {fila_numero}: {e}")

def tiene_check_verde(fila, columna_estado=10):
    if len(fila) > columna_estado and fila[columna_estado].strip():
        estado = fila[columna_estado].strip()
        return estado in ["✅", "☑️", "✓", "√"]
    return False

def obtener_fecha_objetivo():
    from datetime import datetime, timedelta
    ahora = datetime.now()
    hora_actual = ahora.hour
    if 0 <= hora_actual <= 17:
        fecha_objetivo = ahora.date()
        periodo = "MAÑANA"
    else:
        fecha_objetivo = ahora.date() + timedelta(days=1)
        periodo = "TARDE/NOCHE"
    fecha_formateada = fecha_objetivo.strftime("%d/%m/%Y")
    print(f"🕐 Hora actual: {ahora.strftime('%H:%M:%S')} ({periodo})")
    print(f"📅 Fecha objetivo determinada: {fecha_formateada}")
    return fecha_formateada

def filtrar_pedidos_por_fecha_objetivo(valores, fecha_objetivo):
    pedidos_filtrados = []
    fecha_objetivo_normalizada = fecha_objetivo.replace("/", "/")
    print(f"🔍 Filtrando pedidos para la fecha: {fecha_objetivo}")
    pedidos_con_check = 0
    pedidos_sin_fecha = 0
    for i, fila in enumerate(valores[1:], start=2):
        if len(fila) > 2 and fila[2].strip():
            fecha_excel = fila[2].strip()
            fecha_excel_normalizada = fecha_excel.replace("-", "/").replace(".", "/")
            if fecha_excel_normalizada == fecha_objetivo_normalizada:
                if tiene_check_verde(fila):
                    pedidos_con_check += 1
                    print(f"   ⏭️ Fila {i}: Ya procesado (✅), saltando...")
                else:
                    pedidos_filtrados.append((fila, i))
        else:
            pedidos_sin_fecha += 1
    if pedidos_sin_fecha > 0:
        print(f"   ⚠️ Se encontraron {pedidos_sin_fecha} filas sin fecha válida (omitidas)")
    if pedidos_con_check > 0:
        print(f"   ✅ Se encontraron {pedidos_con_check} pedidos ya procesados (saltados)")
    print(f"📊 Pedidos pendientes para {fecha_objetivo}: {len(pedidos_filtrados)}")
    return pedidos_filtrados

def click_button(driver, selector, by=By.CSS_SELECTOR, wait_time=10):
    try:
        button = WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((by, selector)))
        button.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en el selector '{selector}': {e}")
        return False

def click_button_js(driver, selector, by=By.XPATH, wait_time=10):
    try:
        button = WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((by, selector)))
        driver.execute_script("arguments[0].click();", button)
        print("✅ Click realizado con JavaScript")
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic con JS: {e}")
        return False

def write_in_input(driver, selector, text, by=By.ID, wait_time=10):
    try:
        input_element = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((by, selector)))
        input_element.clear()
        input_element.send_keys(text)
        print(f"✅ Texto '{text}' escrito en el input '{selector}'")
        return True
    except Exception as e:
        print(f"❌ Error al escribir en el input '{selector}': {e}")
        return False

def set_fecha(driver, date_text):
    def to_iso(d):
        d = (d or "").strip()
        if "-" in d and len(d.split("-")[0]) == 4:
            return d
        if "/" in d:
            dd, mm, yyyy = d.split("/")
            return f"{yyyy}-{mm.zfill(2)}-{dd.zfill(2)}"
        return d
    iso_date = to_iso(date_text)
    def find_inputs_in_current_frame(wait):
        selectors = [
            "input[data-testid='Widgets::BaseSubfield_input']",
            "input[data-testid*='BaseSubfield_input']",
            "input[data-testid*='DateRangePicker'] input",
            "input[placeholder*='YYYY']",
            "input[type='text'][inputmode='numeric']",
        ]
        for sel in selectors:
            try:
                els = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, sel)))
                if len(els) >= 2:
                    return els[:2]
            except TimeoutException:
                continue
        return None
    driver.switch_to.default_content()
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    print("Cantidad de iframes encontrados:", len(frames))
    inputs = None
    for i in range(len(frames)):
        try:
            driver.switch_to.default_content()
            frames = driver.find_elements(By.TAG_NAME, "iframe")
            driver.switch_to.frame(frames[i])
            print(f"Intentando en iframe {i}...")
            wait = WebDriverWait(driver, 5)
            inputs = find_inputs_in_current_frame(wait)
            if inputs:
                print(f"✅ Campos de fecha encontrados en iframe {i}")
                break
            inner = driver.find_elements(By.TAG_NAME, "iframe")
            for j in range(len(inner)):
                driver.switch_to.frame(inner[j])
                print(f"Intentando en iframe {i}>{j}...")
                wait = WebDriverWait(driver, 5)
                inputs = find_inputs_in_current_frame(wait)
                if inputs:
                    print(f"✅ Campos de fecha encontrados en iframe {i}>{j}")
                    break
                driver.switch_to.parent_frame()
            if inputs:
                break
        except Exception:
            continue
    if not inputs:
        driver.switch_to.default_content()
        raise TimeoutException("No se encontraron los inputs de fecha en ningún iframe.")
    start_date, end_date = inputs[0], inputs[1]
    def fill_input(elem, value):
        elem.click()
        elem.send_keys(Keys.CONTROL, "a")
        elem.send_keys(Keys.DELETE)
        elem.send_keys(value)
    fill_input(start_date, iso_date)
    fill_input(end_date, iso_date)
    end_date.send_keys(Keys.ENTER)
    print(f"✅ Fecha cargada: {iso_date} - {iso_date}")

def escribir_id_pedido(driver, reference_id, wait_time=10):
    try:
        input_element = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.ID, "search_reference--0")))
        input_element.clear()
        input_element.send_keys(reference_id)
        print(f"✅ Referencia '{reference_id}' escrita en el input de búsqueda")
        return True
    except Exception as e:
        print(f"❌ Error al escribir en el input de búsqueda: {e}")
        return False

def click_coordinadora(driver, nombre_cliente, wait_time=10):
    try:
        nombre_lower = nombre_cliente.lower()
        conductor_element = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.XPATH, f"//div[contains(@data-testid, 'DisplayDataCell::Container') and contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{nombre_lower}')]"))
        )
        conductor_element.click()
        print(f"✅ Click realizado en cliente: '{nombre_cliente}'")
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en cliente '{nombre_cliente}'")
        return False

def detectar_y_cambiar_estado_fallido(driver, wait_time=10):
    try:
        input_pendiente = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//input[@value='Pendiente']")))
        if input_pendiente:
            print("✅ Detectado modo 'Pendiente', procediendo a cambiar a Fallido")
            boton_desplegable = WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((By.XPATH, "//button[@data-testid='Trigger::select_status--0']")))
            boton_desplegable.click()
            print("✅ Click en botón desplegable realizado")
            time.sleep(2)
            estrategias = [
                "//span[text()='Fallido']",
                "//div[contains(@class, '_option') and .//span[text()='Fallido']]",
                "//*[contains(text(), 'Fallido')]",
                "//li[contains(., 'Fallido')] | //div[contains(., 'Fallido')]"
            ]
            for selector in estrategias:
                try:
                    opcion_fallido = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, selector)))
                    opcion_fallido.click()
                    print(f"✅ Click en opción 'Fallido' realizado")
                    return True
                except:
                    continue
            print("❌ No se pudo hacer clic en 'Fallido' con ninguna estrategia")
            return False
    except TimeoutException:
        print("⚠️ No se encontró modo 'Pendiente' - probablemente ya está en otro estado")
        return False
    except Exception as e:
        print(f"❌ Error al cambiar estado a Fallido:")
        return False

def click_actualizar_informacion(driver, wait_time=10):
    try:
        boton_actualizar = WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((By.XPATH, "//button[@data-testid='Component::Action-button3--0']")))
        boton_actualizar.click()
        print("✅ Click en 'Actualizar sin fotografía' realizado")
        return True
    except Exception:
        try:
            boton_actualizar = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Actualizar sin fotografía')]")))
            boton_actualizar.click()
            print("✅ Click en 'Actualizar sin fotografía' realizado (selector alternativo)")
            return True
        except Exception as e2:
            print(f"❌ Error con selector alternativo: {e2}")
            return False

def seleccionar_observacion_cancelacion(driver, wait_time=10):
    try:
        input_observacion = WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((By.ID, "select_observation--0")))
        input_observacion.click()
        print("✅ Click en campo de observaciones realizado")
        time.sleep(2)
        try:
            opcion = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='CX (NO USAR)']")))
            opcion.click()
            print("✅ Click en 'CX (NO USAR)' realizado (Estrategia 1)")
            return True
        except:
            try:
                opcion = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'CX (NO USAR)')]")))
                opcion.click()
                print("✅ Click en 'CX (NO USAR)' realizado (Estrategia 2)")
                return True
            except:
                try:
                    opcion = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//li[contains(., 'CX (NO USAR)')] | //div[contains(., 'CX (NO USAR)')]")))
                    opcion.click()
                    print("✅ Click en 'CX (NO USAR)' realizado (Estrategia 3)")
                    return True
                except:
                    opcion = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//div[@data-value='CX (NO USAR)'] | //*[@value='CX (NO USAR)']")))
                    opcion.click()
                    print("✅ Click en 'CX (NO USAR)' realizado (Estrategia 4)")
                    return True
    except Exception:
        print(f"❌ Error al seleccionar observación:")
        return False

def procesar_pedido_individual(driver, idpedido, nombre_cliente, worksheet=None, fila_numero=None):
    try:
        if not escribir_id_pedido(driver, idpedido):
            print(f"❌ No se pudo escribir el ID del pedido: {idpedido}")
            if worksheet and fila_numero:
                marcar_estado_pedido(worksheet, fila_numero, 'fallido')
            return False
        time.sleep(2)
        if not click_coordinadora(driver, nombre_cliente):
            print(f"❌ No se pudo hacer clic en coordinadora: {nombre_cliente}")
            if worksheet and fila_numero:
                marcar_estado_pedido(worksheet, fila_numero, 'fallido')
            return False
        time.sleep(2)
        if detectar_y_cambiar_estado_fallido(driver):
            print("✅ Estado cambiado a Fallido, continuando con el proceso...")
            time.sleep(2)
            if not seleccionar_observacion_cancelacion(driver):
                print("❌ No se pudo seleccionar observación")
                if worksheet and fila_numero:
                    marcar_estado_pedido(worksheet, fila_numero, 'fallido')
                return False
            time.sleep(2)
            if not click_actualizar_informacion(driver):
                print("❌ No se pudo actualizar información")
                if worksheet and fila_numero:
                    marcar_estado_pedido(worksheet, fila_numero, 'fallido')
                return False
            print("🎯 Pedido procesado completamente")
            if worksheet and fila_numero:
                marcar_estado_pedido(worksheet, fila_numero, 'procesado')
            return True
        else:
            print("⏭️ Pedido no está en estado 'Pendiente', saltando al siguiente")
            if worksheet and fila_numero:
                marcar_estado_pedido(worksheet, fila_numero, 'fallido')
            return False
    except Exception as e:
        print(f"❌ Error procesando pedido {idpedido}: {e}")
        if worksheet and fila_numero:
            marcar_estado_pedido(worksheet, fila_numero, 'fallido')
        return False

def procesar_pedidos_desde_excel(driver, pedidos_con_filas, worksheet):
    pedidos_procesados = 0
    pedidos_saltados = 0
    print(f"📋 Total de pedidos a procesar: {len(pedidos_con_filas)}")
    for i, (fila, fila_numero_original) in enumerate(pedidos_con_filas):
        try:
            if len(fila) < 10:
                print(f"⚠️ Fila {fila_numero_original}: Datos insuficientes, saltando...")
                marcar_estado_pedido(worksheet, fila_numero_original, 'fallido')
                pedidos_saltados += 1
                continue
            pedido_id = fila[6].strip() if len(fila) > 6 and fila[6] else ""
            coordinadora = fila[4].strip() if len(fila) > 4 and fila[4] else ""
            if not pedido_id or not coordinadora:
                print(f"⚠️ Fila {fila_numero_original}: Datos faltantes - ID: '{pedido_id}', Coordinadora: '{coordinadora}'. Saltando...")
                marcar_estado_pedido(worksheet, fila_numero_original, 'fallido')
                pedidos_saltados += 1
                continue
            print(f"\n🔄 Procesando pedido {i+1}/{len(pedidos_con_filas)} (Fila Excel: {fila_numero_original})")
            print(f"   🆔 ID: {pedido_id}")
            print(f"   👤 Coordinadora: {coordinadora}")
            if procesar_pedido_individual(driver, pedido_id, coordinadora, worksheet, fila_numero_original):
                pedidos_procesados += 1
                print(f"✅ Pedido {i+1} procesado exitosamente")
                time.sleep(3)
            else:
                pedidos_saltados += 1
                print(f"❌ Pedido {i+1} falló o se saltó")
                time.sleep(1)
            if (i + 1) % 5 == 0:
                print(f"📊 Progreso: {pedidos_procesados} procesados, {pedidos_saltados} saltados de {i+1} total")
        except Exception as e:
            print(f"❌ Error en fila {fila_numero_original}: {e}")
            marcar_estado_pedido(worksheet, fila_numero_original, 'fallido')
            pedidos_saltados += 1
            continue

GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
GOOGLE_CREDENTIALS = json.loads(GOOGLE_CREDENTIALS_JSON)
SHEET_ID = os.getenv("SHEET_ID")

def create_temp_json_file(credentials_dict):
    temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
    json.dump(credentials_dict, temp_file, indent=2)
    temp_file.close()
    return temp_file.name

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
temp_credentials_path = create_temp_json_file(GOOGLE_CREDENTIALS)
credenciales = ServiceAccountCredentials.from_json_keyfile_name(temp_credentials_path, scope)
cliente = gspread.authorize(credenciales)
sheet = cliente.open_by_key(SHEET_ID)
worksheet = sheet.worksheet("Cancelaciones")
valores = worksheet.get_all_values()

options = Options()
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")
options.add_argument("--disable-infobars")
options.add_argument("--disable-extensions")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--start-maximized")
options.add_argument("--window-position=-2000,0")

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

driver.get("https://app2.simpliroute.com/#/login")
time.sleep(5)
write_in_input(driver, "loginUser", os.getenv("SIMPLI_USER"), By.ID)
time.sleep(2)
write_in_input(driver, "loginPass", os.getenv("SIMPLI_PASS"), By.ID)
time.sleep(5)
click_button_js(driver, "//button[contains(@class, 'btn-auth')]", By.XPATH)
time.sleep(10)
driver.get("https://app3.simpliroute.com/extensions")
time.sleep(25)

fecha_objetivo = obtener_fecha_objetivo()
pedidos_del_dia = filtrar_pedidos_por_fecha_objetivo(valores, fecha_objetivo)

if pedidos_del_dia:
    print(f"📅 Configurando fecha en el sistema: {fecha_objetivo}")
    time.sleep(3)
    max_reintentos = 3
    intentos = 0
    while intentos < max_reintentos:
        try:
            set_fecha(driver, fecha_objetivo)
            break
        except Exception as e:
            intentos += 1
            print(f"⚠️ Error al configurar fecha (intento {intentos}): {e}")
            if intentos < max_reintentos:
                print("🔄 Refrescando página y reintentando...")
                driver.get("https://app3.simpliroute.com/extensions")
                time.sleep(25)
            else:
                print("❌ Falló 5 veces al configurar fecha. Abortando script.")
                driver.quit()
                sys.exit()
    time.sleep(3)
    pedidos_seleccionados = []
    omitidos_otro_pais = 0
    omitidos_no_si = 0
    for fila, numero_fila in pedidos_del_dia:
        es_ar = len(fila) > 0 and str(fila[0]).strip().upper() == "AR"
        es_si = len(fila) > 9 and str(fila[9]).strip().upper() == "SI"
        if es_ar and es_si:
            pedidos_seleccionados.append((fila, numero_fila))
        else:
            if not es_ar:
                omitidos_otro_pais += 1
            elif not es_si:
                omitidos_no_si += 1
    if omitidos_otro_pais > 0:
        print(f"🌍 Pedidos de otros países omitidos: {omitidos_otro_pais}")
    if omitidos_no_si > 0:
        print(f"🟡 Pedidos AR con 'NO' en col 9 omitidos: {omitidos_no_si}")
    if pedidos_seleccionados:
        procesar_pedidos_desde_excel(driver, pedidos_seleccionados, worksheet)
    else:
        print("❌ No hay pedidos AR con 'SI' en col 9 para procesar")
        print("🏁 Script finalizado")
        driver.quit()
        sys.exit()
else:
    print(f"❌ No se encontraron pedidos para la fecha {fecha_objetivo}")
    print("🏁 Script finalizado - Sin pedidos para procesar")

driver.quit()
sys.exit()
