import os
import sys
import json
import ssl
import time

import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from slack_sdk import WebClient

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException


def get_env(name, required=True, default=None):
    value = os.getenv(name, default)
    if required and (value is None or str(value).strip() == ""):
        raise RuntimeError(f"Falta la variable de entorno requerida: {name}")
    return value


def build_ssl_context():
    cacert_path = os.getenv("CUSTOM_CACERT_PATH", "").strip()
    if cacert_path and os.path.exists(cacert_path):
        try:
            return ssl.create_default_context(cafile=cacert_path)
        except Exception as e:
            print(f"⚠️ No se pudo crear SSLContext con {cacert_path}: {e}")
    return None


def slack_client():
    token = os.getenv("SLACK_BOT_TOKEN", "").strip()
    if not token:
        return None
    return WebClient(token=token, ssl=build_ssl_context())


def enviar_notificacion_slack(mensaje):
    channel_id = os.getenv("SLACK_CHANNEL_ID", "").strip()
    client = slack_client()
    if not client or not channel_id:
        return
    try:
        resp = client.chat_postMessage(channel=channel_id, text=mensaje)
        if not resp.get("ok", False):
            print(f"❌ Error enviando mensaje a Slack: {resp.get('error')}")
    except Exception as e:
        print(f"❌ Excepción enviando mensaje a Slack: {e}")


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


def click_button_js(driver, selector, by=By.XPATH, wait_time=10):
    try:
        button = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((by, selector))
        )
        driver.execute_script("arguments[0].click();", button)
        print("✅ Click realizado con JavaScript")
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic con JS: {e}")
        return False


def write_in_input(driver, selector, text, by=By.ID, wait_time=10):
    try:
        input_element = WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((by, selector))
        )
        input_element.clear()
        input_element.send_keys(text)
        print(f"✅ Texto '{text}' escrito en el input '{selector}'")
        return True
    except Exception as e:
        print(f"❌ Error al escribir en el input '{selector}': {e}")
        return False


def normalizar_fecha_para_selenium(fecha_str):
    s = str(fecha_str).strip()
    if not s:
        return s
    formatos = ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y")
    for fmt in formatos:
        try:
            dt = pd.to_datetime(s, format=fmt)
            return dt.strftime("%m/%d/%Y")
        except Exception:
            continue
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
        if not pd.isna(dt):
            return dt.strftime("%m/%d/%Y")
    except Exception:
        pass
    return s


def validar_y_completar_fecha(fecha_str):
    fecha_str = str(fecha_str).strip()
    if not fecha_str:
        return None
    partes = fecha_str.split('/')
    if len(partes) == 2:
        dia, mes = partes
        año_actual = pd.Timestamp.now().year
        return f"{dia}/{mes}/{año_actual}"
    elif len(partes) == 3:
        dia, mes, año = partes
        if len(año) == 2:
            return f"{dia}/{mes}/20{año}"
        return fecha_str
    else:
        print(f"❌ Formato de fecha no válido: '{fecha_str}'")
        return None


def escribir_fecha_entrega(driver, fecha_str, intentos):
    fecha_validada = validar_y_completar_fecha(fecha_str)
    if not fecha_validada:
        print(f"❌ Fecha inválida o vacía: '{fecha_str}'.")
        return False
    try:
        input_fecha = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "(//input[@placeholder='MM/DD/YYYY'])[3]"))
        )
        input_fecha.click()
        time.sleep(0.5)
        input_fecha.clear()
        input_fecha.send_keys(Keys.HOME)
        fecha_normalizada = normalizar_fecha_para_selenium(fecha_validada)
        input_fecha.send_keys(fecha_normalizada)
        print(f"✅ Fecha '{fecha_validada}' (normalizada: '{fecha_normalizada}') escrita.")
        return True
    except Exception as e:
        print(f"❌ Error al escribir fecha de entrega: {e}")
        if intentos < 2:
            print("🔄 Reintentando...")
            return escribir_fecha_entrega(driver, fecha_validada, intentos + 1)
        print("❌ No se pudo escribir la fecha tras varios intentos.")
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
        input_element = WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.ID, "search_reference--0"))
        )
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
        input_pendiente = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, "//input[@value='Pendiente']"))
        )
        if input_pendiente:
            print("✅ Detectado modo 'Pendiente', procediendo a cambiar a Fallido")
            boton_desplegable = WebDriverWait(driver, wait_time).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@data-testid='Trigger::select_status--0']"))
            )
            boton_desplegable.click()
            print("✅ Click en botón desplegable realizado")
            time.sleep(2)
            estrategias = [
                "//span[text()='Fallido']",
                "//div[contains(@class, '_option') and .//span[text()='Fallido']]",
                "//*[contains(text(), 'Fallido')]",
                "//li[contains(., 'Fallido')] | //div[contains(., 'Fallido')]"
            ]
            for i, selector in enumerate(estrategias, 1):
                try:
                    opcion_fallido = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    opcion_fallido.click()
                    print(f"✅ Click en opción 'Fallido' realizado (Estrategia {i})")
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
        boton_actualizar = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@data-testid='Component::Action-button3--0']"))
        )
        boton_actualizar.click()
        print("✅ Click en 'Actualizar sin fotografía' realizado")
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en 'Actualizar sin fotografía':")
        try:
            boton_actualizar = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Actualizar sin fotografía')]"))
            )
            boton_actualizar.click()
            print("✅ Click en 'Actualizar sin fotografía' realizado (selector alternativo)")
            return True
        except Exception as e2:
            print(f"❌ Error con selector alternativo: {e2}")
            return False


def seleccionar_observacion_cancelacion(driver, wait_time=10):
    try:
        input_observacion = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.ID, "select_observation--0"))
        )
        input_observacion.click()
        print("✅ Click en campo de observaciones realizado")
        time.sleep(2)
        for selector in [
            "//span[text()='CX (NO USAR)']",
            "//*[contains(text(), 'CX (NO USAR)')]",
            "//li[contains(., 'CX (NO USAR)')] | //div[contains(., 'CX (NO USAR)')]",
            "//div[@data-value='CX (NO USAR)'] | //*[@value='CX (NO USAR)']",
        ]:
            try:
                opcion = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                opcion.click()
                print("✅ 'CX (NO USAR)' seleccionado")
                return True
            except:
                continue
        return False
    except Exception as e:
        print(f"❌ Error al seleccionar observación:")
        return False


def marcar_estado_pedido(worksheet, fila_numero, estado):
    try:
        columna = 10  # Columna N (índice 10 desde 0)
        marca = "✅" if estado == 'procesado' else "❌" if estado == 'fallido' else "?"
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
    ahora = pd.Timestamp.now()
    if 0 <= ahora.hour <= 17:
        fecha_objetivo = ahora.date()
        periodo = "MAÑANA"
    else:
        fecha_objetivo = (ahora + pd.Timedelta(days=1)).date()
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
        print(f"   ⚠️ Filas sin fecha válida (omitidas): {pedidos_sin_fecha}")
    if pedidos_con_check > 0:
        print(f"   ✅ Pedidos ya procesados (saltados): {pedidos_con_check}")

    print(f"📊 Pendientes para {fecha_objetivo}: {len(pedidos_filtrados)}")
    return pedidos_filtrados


def get_gspread_client():
    json_content = os.getenv("GSERVICE_CREDENTIALS_JSON_CONTENT", "").strip()
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if not json_content:
        raise RuntimeError("Falta GSERVICE_CREDENTIALS_JSON_CONTENT")
    try:
        creds_info = json.loads(json_content)
    except json.JSONDecodeError:
        raise RuntimeError("GSERVICE_CREDENTIALS_JSON_CONTENT no es un JSON válido")
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
    return gspread.authorize(creds)


def obtener_pedidos_desde_sheet(client, spreadsheet_id, worksheet_name):
    sheet = client.open_by_key(spreadsheet_id)
    ws = sheet.worksheet(worksheet_name)
    return ws, ws.get_all_values()


def build_driver():
    options = Options()
    options.add_argument("--headless=new")  # Headless en Actions
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    chrome_bin = os.getenv("GOOGLE_CHROME_BIN", "").strip()
    if chrome_bin:
        options.binary_location = chrome_bin
    return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)


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
            print("✅ Estado cambiado a Fallido, continuando...")
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
            print("⏭️ Pedido no está en estado 'Pendiente', saltando")
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
                print(f"✅ Pedido {i+1} procesado")
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


def main():
    enviar_notificacion_slack("Inicio anulación de SIMPLIROUTE ARG 🇦🇷")

    # SimpliRoute
    SIMPLIROUTE_LOGIN_URL = os.getenv("SIMPLIROUTE_LOGIN_URL", "https://app2.simpliroute.com/#/login")
    SIMPLIROUTE_EXT_URL = os.getenv("SIMPLIROUTE_EXT_URL", "https://app3.simpliroute.com/extensions")
    SIMPLIROUTE_USER = get_env("SIMPLIROUTE_USER")
    SIMPLIROUTE_PASSWORD = get_env("SIMPLIROUTE_PASSWORD")

    # Google Sheets
    SHEET_ID = get_env("SHEET_ID")
    GSHEET_WORKSHEET = os.getenv("GSHEET_WORKSHEET", "Cancelaciones")
    gclient = get_gspread_client()
    worksheet, valores = obtener_pedidos_desde_sheet(gclient, SHEET_ID, GSHEET_WORKSHEET)

    # Selenium
    driver = build_driver()
    try:
        # Login
        driver.get(SIMPLIROUTE_LOGIN_URL)
        time.sleep(5)
        write_in_input(driver, "loginUser", SIMPLIROUTE_USER, By.ID)
        time.sleep(2)
        write_in_input(driver, "loginPass", SIMPLIROUTE_PASSWORD, By.ID)
        time.sleep(5)
        click_button_js(driver, "//button[contains(@class, 'btn-auth')]", By.XPATH)
        time.sleep(10)

        driver.get(SIMPLIROUTE_EXT_URL)
        time.sleep(25)

        fecha_objetivo = obtener_fecha_objetivo()
        pedidos_del_dia = filtrar_pedidos_por_fecha_objetivo(valores, fecha_objetivo)

        if not pedidos_del_dia:
            print(f"❌ No se encontraron pedidos para la fecha {fecha_objetivo}")
            print("🏁 Script finalizado - Sin pedidos para procesar")
            return

        print(f"📅 Configurando fecha en el sistema: {fecha_objetivo}")
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
                    driver.get(SIMPLIROUTE_EXT_URL)
                    time.sleep(25)
                else:
                    print("❌ Falló 3 veces al configurar fecha. Abortando script.")
                    return
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
    finally:
        driver.quit()
        enviar_notificacion_slack("Finalizó anulación de SIMPLIROUTE ✅")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ Error fatal: {e}")
        sys.exit(1)
