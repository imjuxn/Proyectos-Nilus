import json
import tempfile
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium. webdriver.common.by import By
from selenium.common.exceptions import *
from datetime import datetime, timedelta
from selenium.webdriver.chrome.options import Options
from oauth2client.service_account import ServiceAccountCredentials
from difflib import SequenceMatcher
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import gspread
import time
from decimal import Decimal, ROUND_HALF_UP
from selenium.webdriver import ActionChains
import requests
import re
import sys
import os
from dotenv import load_dotenv

load_dotenv()

def click_button(driver, selector, by=By.CSS_SELECTOR, wait_time=10):
    try:
        button = WebDriverWait(driver, wait_time).until(
            EC. element_to_be_clickable((by, selector))
        )
        button.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en el selector '{selector}': {e}")
        return False

def buscar_pedido(driver, pedido_id):
    """
    Abre directamente un pedido usando su ID de Odoo en la URL. 
    """
    try:
        odoo_base_url = os.getenv("ODOO_BASE_URL")
        menu_id = os.getenv("ODOO_MENU_ID", "182")
        action_id = os.getenv("ODOO_ACTION_ID", "299")
        url = f"{odoo_base_url}/web#id={pedido_id}&menu_id={menu_id}&cids=1&action={action_id}&model=sale.order&view_type=form"
        driver.get(url)
        time.sleep(10)
        print(f"✅ Pedido ID {pedido_id} abierto correctamente")
        return True
    except Exception as e:
        print(f"❌ Error al abrir pedido ID {pedido_id}: {e}")
        return False

def parse_moneda(texto: str) -> Decimal:
    """
    Convierte strings monetarios a Decimal.
    """
    if texto is None:
        raise ValueError("Texto vacío")

    s = str(texto).strip(). replace('\u00A0', '').replace('\u202F', '')
    sign = '-' if s. lstrip(). startswith('-') else ''
    cleaned = ''.join(ch for ch in s if ch.isdigit() or ch in '.,')
    if not cleaned:
        raise ValueError(f"No se encontraron dígitos en: {texto! r}")

    last_comma = cleaned.rfind(',')
    last_dot = cleaned.rfind('.')

    decimal_sep = None
    if last_comma != -1 and last_dot != -1:
        decimal_sep = ',' if last_comma > last_dot else '.'
    elif last_comma != -1 or last_dot != -1:
        idx = last_comma if last_comma != -1 else last_dot
        if (len(cleaned) - idx - 1) == 2:
            decimal_sep = ',' if last_comma != -1 else '.'

    if decimal_sep:
        thousand_sep = '.' if decimal_sep == ',' else ','
        tmp = cleaned.replace(thousand_sep, '')
        num_str = tmp.replace(decimal_sep, '.')
    else:
        num_str = cleaned. replace(',', ''). replace('.', '')

    if sign:
        num_str = sign + num_str
    return Decimal(num_str)

def formatear_moneda_mx(valor: Decimal) -> str:
    """
    Convierte Decimal -> string con formato MX: '5,600.00'
    (coma miles, punto decimal).  Siempre 2 decimales.
    """
    if not isinstance(valor, Decimal):
        valor = Decimal(str(valor))
    v = valor.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    s = f"{v:. 2f}"
    entero, dec = s.split('.')
    grupos = []
    while len(entero) > 3:
        grupos.append(entero[-3:])
        entero = entero[:-3]
    grupos.append(entero)
    grupos = grupos[::-1]
    entero_fmt = ','.join(grupos)
    return f"{entero_fmt}.{dec}"

def ajustar_price_unit_con_descuento(
    driver,
    total_porcentaje_18: Decimal,
    min_price: Decimal = None,
    locator_xpath: str = "//td[@name='price_unit']",
    click_delay: float = 0.8,
    after_enter_delay: float = 1.2,
    usar_title_prioritario: bool = True
):
    """
    Ajusta el precio unitario en la celda 'price_unit' restando el porcentaje acumulado. 
    """
    if min_price is None:
        min_price = Decimal(os.getenv("MIN_PRICE_UNIT", "39"))
    
    try:
        celda = driver.find_element(By. XPATH, locator_xpath)
    except Exception as e:
        print(f"❌ No se encontró la celda price_unit: {e}")
        return None

    raw_title = celda.get_attribute('title') if usar_title_prioritario else None
    raw_text = celda.text
    base_texto = raw_title if (usar_title_prioritario and raw_title) else raw_text

    try:
        original = parse_moneda(base_texto)
    except Exception:
        original = Decimal('0')

    if not isinstance(total_porcentaje_18, Decimal):
        total_porcentaje_18 = Decimal(str(total_porcentaje_18))

    provisional = original - total_porcentaje_18
    forzado_a_minimo = provisional < min_price
    nuevo_valor = min_price if forzado_a_minimo else provisional
    nuevo_valor = nuevo_valor.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    valor_str = formatear_moneda_mx(nuevo_valor)

    print(f"💡 Precio original: {original} | Total 18% acumulado a restar: {total_porcentaje_18}")
    if forzado_a_minimo:
        print(f"⚠️ La resta bajaba de {min_price}, se fuerza a mínimo.")
    print(f"➡️ Nuevo precio a establecer: {nuevo_valor} (enviado como '{valor_str}')")

    try:
        celda. click()
        time.sleep(click_delay)
        acciones = ActionChains(driver)
        acciones.key_down(Keys.CONTROL). send_keys('a').key_up(Keys.CONTROL).perform()
        acciones = ActionChains(driver)
        acciones.send_keys(Keys. BACK_SPACE).perform()
        acciones = ActionChains(driver)
        acciones.send_keys(valor_str).send_keys(Keys.ENTER).perform()
        time.sleep(after_enter_delay)
    except Exception as e:
        print(f"⚠️ Error enviando nuevo valor a price_unit: {e}")

    return {
        'original': original,
        'descuento_solicitado': total_porcentaje_18,
        'descuento_aplicado': (original - nuevo_valor) if original > nuevo_valor else Decimal('0'),
        'nuevo': nuevo_valor,
        'tope_minimo': min_price,
        'forzado_a_minimo': forzado_a_minimo,
        'valor_enviado': valor_str
    }

def editar_cantidad_por_producto(driver, nombre_producto, nueva_cantidad="0", wait_time=10):
    """
    Pone en nueva_cantidad todos los productos cuyo nombre coincide. 
    """
    xpath_producto = (
        f"//span[@name='name' and contains(translate(text(), "
        f"'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), "
        f"'{nombre_producto. lower()}')]"
    )
    try:
        productos_iniciales = driver.find_elements(By.XPATH, xpath_producto)
        total_productos = len(productos_iniciales)
        
        if total_productos == 0:
            print(f"❌ No se encontró el producto '{nombre_producto}'")
            print("🔍 Productos encontrados en el pedido (primeros 15):")
            todos_productos = driver.find_elements(By.XPATH, "//span[@name='name']")
            for idx, prod in enumerate(todos_productos[:15], 1):
                texto_crudo = prod.text
            return False

        print(f"✅ Se encontraron {total_productos} productos '{nombre_producto}'")
    except Exception as e:
        print(f"❌ Error al buscar producto '{nombre_producto}': {e}")
        return False

    productos_editados = 0
    intentos = 0
    max_intentos = total_productos * 2
    total_porcentaje_18 = Decimal("0. 00")
    porcentaje_descuento = Decimal(os.getenv("PORCENTAJE_DESCUENTO", "0. 18"))

    while productos_editados < total_productos and intentos < max_intentos:
        intentos += 1
        try:
            productos = driver.find_elements(By. XPATH, xpath_producto)
            if not productos:
                print("⚠️ Ya no hay más productos para editar")
                break

            producto_a_editar = None
            fila_a_editar = None
            for producto in productos:
                fila = producto.find_element(By. XPATH, "./ancestor::tr")
                celda_cantidad = fila.find_element(By.XPATH, ". //td[@name='product_uom_qty']")
                cantidad_actual_txt = celda_cantidad.text. strip(). replace(",", ".")
                try:
                    if float(cantidad_actual_txt) != float(nueva_cantidad):
                        producto_a_editar = producto
                        fila_a_editar = fila
                        break
                except:
                    producto_a_editar = producto
                    fila_a_editar = fila
                    break

            if not producto_a_editar:
                print("✅ Todos los productos ya tienen la cantidad objetivo")
                break

            try:
                try:
                    el_subtotal = fila_a_editar.find_element(By.XPATH, ".//span[@name='price_subtotal']")
                    texto_subtotal = el_subtotal.text
                    subtotal = parse_moneda(texto_subtotal)
                except:
                    celda_precio = fila_a_editar.find_element(By. XPATH, ".//td[@name='price_unit']")
                    precio_txt = celda_precio.get_attribute("title") or celda_precio.text
                    celda_qty = fila_a_editar.find_element(By. XPATH, ".//td[@name='product_uom_qty']")
                    qty_txt = celda_qty.get_attribute("title") or celda_qty.text
                    precio = parse_moneda(precio_txt)
                    qty = parse_moneda(qty_txt)
                    subtotal = (precio * qty).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

                porc_18 = (subtotal * porcentaje_descuento).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                total_porcentaje_18 += porc_18
                print(f"💡 Subtotal: {subtotal} | {porcentaje_descuento*100}%: {porc_18} | Acumulado: {total_porcentaje_18}")
            except Exception as e:
                print(f"⚠️ No se pudo calcular el porcentaje de la fila: {e}")

            celda_cantidad = fila_a_editar. find_element(By.XPATH, ".//td[@name='product_uom_qty']")
            wait = WebDriverWait(driver, 15)
            wait.until(EC.element_to_be_clickable(celda_cantidad))

            celda_cantidad.click()
            time.sleep(3)
            ActionChains(driver).send_keys(str(nueva_cantidad)).send_keys(Keys.ENTER). perform()
            
            productos_editados += 1
            print(f"✅ Producto {productos_editados}/{total_productos}: cantidad editada a {nueva_cantidad}")
            time.sleep(3)

            click_button(driver, "//button[@class='btn btn-primary o_form_button_save']", By.XPATH)
            print("✅ Pedido guardado correctamente")
            time. sleep(3)

        except Exception as e:
            print(f"⚠️ Error en iteración: {e}")
            continue

    print(f"🔢 Total {porcentaje_descuento*100}% acumulado: {total_porcentaje_18}")
    time.sleep(2)
    
    if total_porcentaje_18 > 0:
        tasa_servicio_nombre = os.getenv("TASA_SERVICIO_NOMBRE", "tasa de serv")
        xpath_tasa_price_unit = (
            f"//span[@name='name' and contains(translate(text(), "
            f"'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ','abcdefghijklmnopqrstuvwxyzáéíóúüñ'), "
            f"'{tasa_servicio_nombre}')]/ancestor::tr//td[@name='price_unit']"
        )
        time.sleep(2)
        ajuste = ajustar_price_unit_con_descuento(
            driver,
            total_porcentaje_18=total_porcentaje_18,
            locator_xpath=xpath_tasa_price_unit
        )
        time.sleep(2)
        click_button(driver, "//button[@class='btn btn-primary o_form_button_save']", By.XPATH)
        print("Resumen ajuste price_unit:", ajuste)
    else:
        print(f"⚠️ No se acumuló ningún porcentaje para ajustar la tasa.")

    return True

def Faltante_completo(driver, pedido_id, nombre_producto, nueva_cantidad="0"):
    try:
        if not buscar_pedido(driver, pedido_id):
            print(f"❌ No se pudo procesar el pedido {pedido_id}")
            return False
        
        if not editar_cantidad_por_producto(driver, nombre_producto, nueva_cantidad):
            print(f"❌ No se pudo editar el producto '{nombre_producto}'")
            return False
        
        time.sleep(2)
        print(f"✅ Faltante completo procesado para pedido {pedido_id}")
        return True
        
    except Exception as e:
        print(f"❌ Error en Faltante_completo: {e}")
        return False

def editar_parcial(driver, nombre_producto, nueva_cantidad, cantidad_original, wait_time=10):
    xpath_producto = (
        f"//span[@name='name' and contains(translate(text(), "
        f"'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ','abcdefghijklmnopqrstuvwxyzáéíóúüñ'), "
        f"'{nombre_producto.lower()}')]"
    )
    try:
        productos = driver.find_elements(By. XPATH, xpath_producto)
        if not productos:
            print("❌ No se encontró el producto")
            return False
        if len(productos) > 1:
            print(f"⚠️ {len(productos)} coincidencias.  Se omite.")
            return False

        fila = productos[0]. find_element(By.XPATH, "./ancestor::tr")

        celda_precio = fila.find_element(By.XPATH, ".//td[@name='price_unit']")
        precio_raw = (celda_precio.get_attribute("title") or celda_precio.text).strip()
        try:
            precio_unit = parse_moneda(precio_raw)
        except:
            precio_unit = Decimal('0')
        precio_original_str = formatear_moneda_mx(precio_unit)
        print(f"💡 price_unit original: {precio_unit}")

        try:
            unidades_faltantes = Decimal(str(cantidad_original)) - Decimal(str(nueva_cantidad))
        except:
            unidades_faltantes = Decimal('0')

        porcentaje_descuento = Decimal(os.getenv("PORCENTAJE_DESCUENTO", "0.18"))
        
        if unidades_faltantes > 0 and precio_unit > 0:
            valor_faltante = (precio_unit * unidades_faltantes).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            porc_18 = (valor_faltante * porcentaje_descuento).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            print(f"💡 Unidades faltantes: {unidades_faltantes} | Valor: {valor_faltante} | {porcentaje_descuento*100}%: {porc_18}")
        else:
            porc_18 = Decimal('0')
            print("⚠️ Sin faltante o precio = 0.  No ajuste.")

        celda_cantidad = fila.find_element(By. XPATH, ".//td[@name='product_uom_qty']")
        wait = WebDriverWait(driver, 15)
        wait.until(EC.element_to_be_clickable(celda_cantidad))

        celda_cantidad. click()
        time.sleep(3)
        ActionChains(driver).send_keys(str(nueva_cantidad)).send_keys(Keys.ENTER).perform()
        print(f"✅ Cantidad editada a {nueva_cantidad}")
        time.sleep(3)

        wait = WebDriverWait(driver, 15)

        try:
            wait.until(EC.presence_of_element_located((By. XPATH, xpath_producto)))
            productos2 = driver.find_elements(By.XPATH, xpath_producto)
            if not productos2:
                print("⚠️ Producto no encontrado tras editar (DOM refrescado)")
                return True

            fila2 = productos2[0]. find_element(By.XPATH, "./ancestor::tr")
            celda_precio2 = wait.until(
                EC.visibility_of_element_located((By.XPATH, ".//td[@name='price_unit']"))
            )
            celda_precio2 = fila2.find_element(By.XPATH, ".//td[@name='price_unit']")

            wait.until(EC.element_to_be_clickable(celda_precio2))
            celda_precio2.click()
            time.sleep(3)
            ActionChains(driver).send_keys(precio_original_str).send_keys(Keys.ENTER).perform()
            time.sleep(3)

            click_button(driver, "//button[@class='btn btn-primary o_form_button_save']", By.XPATH)
            print("✅ Pedido guardado con price_unit restaurado")
            time.sleep(3)
        except Exception as e:
            print(f"⚠️ Error restaurando price_unit: {e}")

        if porc_18 > 0:
            tasa_servicio_nombre = os.getenv("TASA_SERVICIO_NOMBRE", "tasa de serv")
            xpath_tasa_price_unit = (
                f"//span[@name='name' and contains(translate(text(), "
                f"'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ','abcdefghijklmnopqrstuvwxyzáéíóúüñ'), "
                f"'{tasa_servicio_nombre}')]/ancestor::tr//td[@name='price_unit']"
            )
            time.sleep(3)
            ajuste = ajustar_price_unit_con_descuento(
                driver,
                total_porcentaje_18=porc_18,
                locator_xpath=xpath_tasa_price_unit
            )
            time.sleep(3)
            click_button(driver, "//button[@class='btn btn-primary o_form_button_save']", By.XPATH)
            print("Resumen ajuste price_unit:", ajuste)
        else:
            print(f"⚠️ {porcentaje_descuento*100}% = 0.  No ajuste de tasa.")

        return True

    except Exception as e:
        print(f"❌ Error en editar_parcial '{nombre_producto}': {e}")
        return False
    
def Faltante_parcial(driver, pedido_id, nombre_producto, nueva_cantidad, cantidad_original):
    try:
        if not buscar_pedido(driver, pedido_id):
            print(f"❌ No se pudo procesar el pedido {pedido_id}")
            return False
        
        if not editar_parcial(driver, nombre_producto, nueva_cantidad, cantidad_original):
            print(f"❌ No se pudo editar el producto '{nombre_producto}'")
            return False
        time.sleep(2)
        
        print(f"✅ Faltante parcial procesado para pedido {pedido_id}")
        return True
    except Exception as e:
        print(f"❌ Error en Faltante_parcial: {e}")
        return False

def get_google_credentials():
    return {
        "type": "service_account",
        "project_id": os.getenv("GOOGLE_PROJECT_ID"),
        "private_key_id": os.getenv("GOOGLE_PRIVATE_KEY_ID"),
        "private_key": os.getenv("GOOGLE_PRIVATE_KEY").replace('\\n', '\n'),
        "client_email": os.getenv("GOOGLE_CLIENT_EMAIL"),
        "client_id": os.getenv("GOOGLE_CLIENT_ID"),
        "auth_uri": "https://accounts. google.com/o/oauth2/auth",
        "token_uri": "https://oauth2. googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": os.getenv("GOOGLE_CLIENT_CERT_URL"),
        "universe_domain": "googleapis.com"
    }

def create_temp_json_file(credentials_dict):
    temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
    json.dump(credentials_dict, temp_file, indent=2)
    temp_file.close()
    return temp_file. name

def conectar_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis. com/auth/drive"]
    
    google_credentials = get_google_credentials()
    temp_credentials_path = create_temp_json_file(google_credentials)
    
    credenciales = ServiceAccountCredentials.from_json_keyfile_name(temp_credentials_path, scope)
    cliente = gspread.authorize(credenciales)
    
    os.remove(temp_credentials_path)
    
    return cliente

def main():
    cliente = conectar_google_sheets()
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    sheet = cliente.open_by_key(sheet_id)
    worksheet_name = os.getenv("GOOGLE_WORKSHEET_NAME", "check_nueva_info_desvios")
    worksheet = sheet.worksheet(worksheet_name)
    valores = worksheet.get_all_values()

    options = Options()
    headless_mode = os.getenv("HEADLESS_MODE", "true").lower() == "true"
    if headless_mode:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager(). install()), options=options)
    driver.implicitly_wait(10)
    wait = WebDriverWait(driver, 20)

    odoo_url = os.getenv("ODOO_LOGIN_URL")
    odoo_user = os.getenv("ODOO_USER")
    odoo_password = os.getenv("ODOO_PASSWORD")
    
    driver.get(odoo_url)

    campo_login = wait.until(EC.visibility_of_element_located((By.ID, "login")))
    campo_login. send_keys(odoo_user)

    campo_password = wait.until(EC.visibility_of_element_located((By.ID, "password")))
    campo_password.send_keys(odoo_password)
    time.sleep(3)
    
    boton_login = wait.until(EC.element_to_be_clickable((By. XPATH, "//button[text()='Log in']")))
    boton_login.click()
    time. sleep(3)
    
    try:
        boton_mensaje = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "message-btn")))
        time.sleep(2)
        boton_mensaje. click()
        print("✅ Botón de mensaje clickeado")
    except Exception as e:
        print(f"⚠️ No se encontró o no se pudo clickear el botón de mensaje (continuando): {e}")
    
    time.sleep(10)

    dias_offset = int(os.getenv("DIAS_OFFSET", "0"))
    fecha_hoy = (datetime.now() + timedelta(days=dias_offset)).strftime("%d/%m/%Y")
    print(f"📅 Fecha de hoy: {fecha_hoy}")

    pais_filtro = os.getenv("PAIS_FILTRO", "mx"). lower()

    for i, fila in enumerate(valores):
        if i == 0:
            continue
        
        fecha_pedido = fila[0]
        pais = fila[1]
        tipo_faltante = fila[2]
        num_pedido = fila[8]
        nombre_producto = fila[9]
        pedido_id = fila[14]
        
        if fecha_pedido != fecha_hoy:
            continue
        
        if pais. lower() != pais_filtro:
            continue
        
        if not pedido_id or not nombre_producto:
            print(f"⚠️ Fila {i} omitida: datos incompletos")
            worksheet.update_cell(i + 1, 16, "❌")
            continue
        
        resultado = False
        
        try:
            if tipo_faltante. lower() == "faltante":
                print(f"🔄 Procesando faltante completo - Pedido: {num_pedido} (ID: {pedido_id})")
                resultado = Faltante_completo(driver, pedido_id, nombre_producto, nueva_cantidad="0")
                time.sleep(2)
                
            elif tipo_faltante.lower() == "faltante_parcial":
                cantidad_original = (fila[11] or ""). strip() if len(fila) > 11 else "0"
                cantidad_enviada = (fila[12] or "").strip() if len(fila) > 12 else "0"
                
                if not cantidad_original or not cantidad_enviada:
                    print(f"⚠️ Fila {i}: faltan cantidades (col 11/12)")
                    worksheet.update_cell(i + 1, 16, "❌")
                    continue
                    
                print(f"🔄 Faltante parcial | Pedido: {num_pedido} (ID: {pedido_id}) | Original: {cantidad_original} | Enviada: {cantidad_enviada}")
                resultado = Faltante_parcial(driver, pedido_id, nombre_producto, nueva_cantidad=cantidad_enviada, cantidad_original=cantidad_original)
                time.sleep(2)
            else:
                print(f"⚠️ Fila {i}: tipo de faltante no reconocido '{tipo_faltante}'")
                worksheet.update_cell(i + 1, 16, "❌")
                continue
        except Exception as e:
            print(f"❌ Error procesando fila {i}: {e}")
            resultado = False
        
        if resultado:
            worksheet.update_cell(i + 1, 16, "✅")
            print(f"✅ Fila {i}: Marcada como exitosa")
        else:
            worksheet.update_cell(i + 1, 16, "❌")
            print(f"❌ Fila {i}: Marcada como fallida")
        
        time.sleep(1)

    print("✅ Proceso de desvíos operativos completado")
    driver.quit()

if __name__ == "__main__":
    main()
