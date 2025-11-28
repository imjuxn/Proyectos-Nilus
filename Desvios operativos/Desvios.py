import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome. service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium. webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver. common.by import By
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
import json
import tempfile
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()

def enviar_notificacion_slack(mensaje):
    slack_webhook_url = os.getenv("SLACK_WEBHOOK_URL")
    payload = {"text": mensaje}
    try:
        response = requests. post(slack_webhook_url, json=payload)
        if response.status_code != 200:
            print(f"❌ Error enviando mensaje a Slack: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"❌ Excepción enviando mensaje a Slack: {e}")

# Notifica al iniciar el script
enviar_notificacion_slack("🚀 El script de procesamiento de desvíos ha comenzado EN AMBOS PAÍSES.")

def normalizar(texto):
    """Elimina símbolos, convierte a minúsculas y separa palabras clave."""
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
            EC. element_to_be_clickable((by, selector))
        )
        button.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en el selector '{selector}'")
        return False

def procesar_pedido(driver, datos_pedido, producto_buscado, cantidad_deseada):
    def similitud(nombre1, nombre2):
        return SequenceMatcher(None, nombre1.lower(), nombre2.lower()).ratio()

    backoffice_url = os.getenv("BACKOFFICE_BASE_URL")
    locale = os.getenv("BACKOFFICE_LOCALE", "es-AR")
    pedido_url = f"{backoffice_url}/{locale}/orders/{datos_pedido}"
    print(f"\n🔄 Procesando pedido: {datos_pedido}")
    driver.get(pedido_url)
    time.sleep(5)

    try:
        productos = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "css-1es96wk"))
        )

        # Primera instancia: buscar con similitud >= umbral configurado
        umbral_similitud = float(os.getenv("UMBRAL_SIMILITUD", "0.8"))
        mejor_match = None
        mayor_similitud = 0

        for producto in productos:
            nombre = producto. find_element(By.TAG_NAME, "span").text.strip()
            score = similitud(producto_buscado, nombre)
            if score > mayor_similitud:
                mayor_similitud = score
                mejor_match = (producto, nombre)

        if mejor_match and mayor_similitud >= umbral_similitud:
            producto, nombre = mejor_match
            print(f"🔍 [Terminal] Coincidencia por similitud utilizada ({mayor_similitud:.2f}) → Producto encontrado: {nombre}")
        else:
            # Segunda instancia: usar coincidencia parcial
            mejor_match = None
            umbral_parcial = float(os.getenv("UMBRAL_PARCIAL", "0.5"))
            for producto in productos:
                nombre = producto.find_element(By.TAG_NAME, "span").text.strip()
                if coincidencia_parcial(producto_buscado, nombre, umbral=umbral_parcial):
                    mejor_match = (producto, nombre)
                    print(f"🔍 [Terminal] Coincidencia parcial utilizada → Producto encontrado: {nombre}")
                    break

        if not mejor_match:
            print(f"❌ Producto '{producto_buscado}' no encontrado por ningún método.")
            enviar_notificacion_slack(f"❌ Producto '{producto_buscado}' no encontrado en pedido {datos_pedido}.")
            return

        producto, nombre = mejor_match
        boton_svg = producto.find_element(By.CSS_SELECTOR, 'svg[data-testid="DoNotDisturbOnIcon"]')
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton_svg)
        try:
            boton_svg.click()
        except Exception as e1:
            try:
                driver.execute_script("arguments[0].click();", boton_svg)
            except Exception as e2:
                mensaje_error = f"❌ No se pudo hacer clic en el botón del producto '{nombre}' en pedido {datos_pedido}"
                print(mensaje_error)
                enviar_notificacion_slack(mensaje_error)
                return

        motivo_label = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//label[contains(text(), "Selecciona un motivo")]'))
        )
        combo_div = motivo_label.find_element(By.XPATH, './ancestor::div[contains(@class, "MuiFormControl-root")]')
        combo_clickable = combo_div.find_element(By.XPATH, './/div[@role="button" or @role="combobox"]')
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", combo_clickable)
        combo_clickable.click()
        
        motivo_seleccionado = os.getenv("MOTIVO_DESVIO", "Support - DTC - Operations - Missing Product")
        opcion = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, f'//li[contains(text(), "{motivo_seleccionado}")]'))
        )
        opcion.click()

        input_cantidad = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "quantity"))
        )
        input_cantidad.clear()
        input_cantidad.send_keys(str(cantidad_deseada))

        WebDriverWait(driver, 10).until(
            EC. element_to_be_clickable((By.XPATH, "//button[normalize-space()='Solicitar']"))
        ).click()

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Guardar cambios']"))
        ).click()

        print(f"✅ Pedido {datos_pedido} procesado con éxito.")

    except Exception as e:
        mensaje_error = f"❌ Error procesando el pedido {datos_pedido}: FIJARSE en la planilla con el ID"
        print(mensaje_error)
        enviar_notificacion_slack(mensaje_error)

def get_google_credentials():
    """Construye el diccionario de credenciales desde variables de entorno"""
    return {
        "type": "service_account",
        "project_id": os.getenv("GOOGLE_PROJECT_ID"),
        "private_key_id": os.getenv("GOOGLE_PRIVATE_KEY_ID"),
        "private_key": os.getenv("GOOGLE_PRIVATE_KEY"). replace('\\n', '\n'),
        "client_email": os.getenv("GOOGLE_CLIENT_EMAIL"),
        "client_id": os.getenv("GOOGLE_CLIENT_ID"),
        "auth_uri": "https://accounts. google.com/o/oauth2/auth",
        "token_uri": "https://oauth2. googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": os.getenv("GOOGLE_CLIENT_CERT_URL"),
        "universe_domain": "googleapis.com"
    }

def create_temp_json_file(credentials_dict):
    """Crea un archivo JSON temporal con las credenciales"""
    temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
    json.dump(credentials_dict, temp_file, indent=2)
    temp_file.close()
    return temp_file.name

def conectar_google_sheets():
    """Conecta a Google Sheets usando credenciales de variables de entorno"""
    scope = ["https://spreadsheets. google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    google_credentials = get_google_credentials()
    temp_credentials_path = create_temp_json_file(google_credentials)
    
    credenciales = ServiceAccountCredentials. from_json_keyfile_name(temp_credentials_path, scope)
    cliente = gspread.authorize(credenciales)
    
    # Limpiar archivo temporal
    os.remove(temp_credentials_path)
    
    return cliente

def extraer_id(texto):
    """Extrae el ID del pedido (hash de 32 caracteres)"""
    match = re.search(r"\b[a-f0-9]{32}\b", str(texto))
    return match.group(0) if match else None

def main():
    # === INICIO DEL SCRIPT ===
    options = Options()
    headless_mode = os.getenv("HEADLESS_MODE", "true").lower() == "true"
    if headless_mode:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    # LOGIN
    backoffice_url = os.getenv("BACKOFFICE_BASE_URL")
    locale = os.getenv("BACKOFFICE_LOCALE", "es-AR")
    backoffice_user = os.getenv("BACKOFFICE_USER")
    backoffice_password = os.getenv("BACKOFFICE_PASSWORD")
    
    driver.get(f"{backoffice_url}/{locale}/login")
    time.sleep(5)
    driver.find_element(By.ID, "email").send_keys(backoffice_user)
    driver.find_element(By. ID, "password").send_keys(backoffice_password)
    click_button(driver, "//button[text()='INGRESAR']", By.XPATH)
    time. sleep(10)

    # === CONEXIÓN A GOOGLE SHEETS ===
    cliente = conectar_google_sheets()
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    sheet = cliente.open_by_key(sheet_id)
    
    worksheet_name = os.getenv("GOOGLE_WORKSHEET_NAME", "check_nueva_info_desvios")
    worksheet = sheet.worksheet(worksheet_name)
    valores = worksheet.get_all_values()

    filas = valores[1:]  # Salteamos encabezado

    # === FILTRAR FILAS SEGÚN PAÍSES CONFIGURADOS ===
    paises_permitidos = os.getenv("PAISES_PERMITIDOS", "mx,ar").lower().split(",")
    paises_permitidos = [p.strip() for p in paises_permitidos]
    
    filas_filtradas = [fila for fila in filas if len(fila) > 1 and str(fila[1]).strip().lower() in paises_permitidos]

    df = pd.DataFrame([
        [fila[0], fila[2], fila[7], fila[9], fila[11], fila[12] if len(fila) > 12 else "0"]
        for fila in filas_filtradas if len(fila) >= 12
    ], columns=["fecha", "type_desvio", "datos_pedido", "producto_afectado", "cantidad_original", "cantidad_modificada"])

    print("📄 Pedidos leídos del Excel:")
    print(df)

    df["fecha"] = df["fecha"].astype(str). str.strip()
    df["fecha"] = pd.to_datetime(df["fecha"], dayfirst=True, errors="coerce")
    
    df["datos_pedido"] = df["datos_pedido"].apply(extraer_id)

    # Limpieza de datos nulos
    df. dropna(subset=["fecha"], inplace=True)

    # Obtener la fecha según días de offset configurados
    dias_offset = int(os.getenv("DIAS_OFFSET", "-1"))
    fecha_filtro = (datetime.now() + timedelta(days=dias_offset)).date()
    print("Fecha de filtro:", fecha_filtro)

    # Filtrar por fecha
    df_filtrado = df[df["fecha"]. dt.date == fecha_filtro]

    print(f"📆 Fecha actual: {datetime.now().strftime('%d/%m/%Y')}")
    print(f"📆 Fecha filtrada: {fecha_filtro}")
    print(f"🔎 Filas encontradas: {len(df_filtrado)}")

    # Filtrar solo pedidos con estado 'faltante' o 'faltante_parcial'
    df_filtrado = df_filtrado[df_filtrado["type_desvio"].isin(["faltante", "faltante_parcial"])]

    if df_filtrado.empty:
        print("⚠️ No se encontraron pedidos para la fecha configurada.")
        driver.quit()
        exit()

    # PROCESAR PEDIDOS UNO A UNO
    for i, row in df_filtrado.iterrows():
        try:
            datos_pedido = str(row["datos_pedido"]).strip()
            producto_afectado = str(row["producto_afectado"]).strip()

            # Limpieza segura de columnas numéricas
            try:
                original = int(row. get("cantidad_original", 0) or 0)
            except ValueError:
                original = 0

            try:
                modificada = int(row["cantidad_modificada"]) if str(row["cantidad_modificada"]).strip(). isdigit() else 0
            except:
                modificada = 0

            tipo_desvio = str(row. get("type_desvio", "")).strip().lower()
            if tipo_desvio == "faltante_parcial":
                cantidad_deseada = max(original - modificada, 0)
            else:
                cantidad_deseada = original

            procesar_pedido(driver, datos_pedido, producto_afectado, cantidad_deseada)
        except Exception as e:
            print(f"⚠️ Error en fila {i}: {e}")
    
    time.sleep(10)
    driver.quit()
    print("✅ Proceso finalizado y navegador cerrado.")

if __name__ == "__main__":
    main()
