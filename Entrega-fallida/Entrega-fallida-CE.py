import os
import sys
import json
import time
import ssl

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


def escribir_punto_de_entrega(driver, texto, intentos):
    try:
        print(f"📝 Escribiendo '{texto}' en deliveryPoint")
        input_campo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "deliveryPoint"))
        )
        input_campo.clear()
        input_campo.send_keys(texto)
        input_campo.send_keys(Keys.ENTER)
        time.sleep(1)
        opcion = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//li[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{texto.lower()}')]"))
        )
        opcion.click()
        time.sleep(2)
        print("✅ Punto de entrega seleccionado.")
        return True
    except Exception as e:
        print(f"❌ Error seleccionando punto de entrega: {e}")
        if intentos < 2:
            print("🔄 Reintentando...")
            return escribir_punto_de_entrega(driver, texto, intentos + 1)
        print("❌ No se pudo seleccionar el punto de entrega.")
        return False


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


def guardar_cambios(driver, wait_time=10):
    try:
        modal = WebDriverWait(driver, wait_time).until(
            EC.visibility_of_element_located((By.XPATH, "//h2[contains(text(), 'Cambiar el estado del pedido')]/ancestor::div[@role='dialog']"))
        )
        guardar_btn = modal.find_element(By.XPATH, ".//button[normalize-space(text())='Guardar cambios']")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", guardar_btn)
        driver.execute_script("arguments[0].click();", guardar_btn)
        print("✅ 'Guardar cambios' clickeado.")
        return True
    except Exception as e:
        print(f"❌ No se pudo hacer clic en el botón correcto: {e}")
        return False


def get_gspread_client():
    json_content = os.getenv("GSERVICE_CREDENTIALS_JSON_CONTENT", "").strip()
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    if not json_content:
        raise RuntimeError("Falta GSERVICE_CREDENTIALS_JSON_CONTENT")
    try:
        creds_info = json.loads(json_content)
    except json.JSONDecodeError:
        raise RuntimeError("GSERVICE_CREDENTIALS_JSON_CONTENT no es un JSON válido")
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
    return gspread.authorize(creds)


def obtener_pedidos_planilla(client, spreadsheet_id, sheet_name,
                             country_col=0, fecha_col=1, punto_col=2, header_rows=1):
    sh = client.open_by_key(spreadsheet_id)
    sheets = sh.worksheets()
    print("DEBUG: Spreadsheet:", sh.title)
    print("DEBUG: Hojas disponibles:", [s.title for s in sheets])
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        raise RuntimeError(f"Hoja '{sheet_name}' no encontrada en el spreadsheet")
    all_values = ws.get_all_values()
    data = all_values[header_rows:]

    pedidos = []
    filas_saltadas = []
    for idx, row in enumerate(data, start=header_rows + 1):
        if len(row) <= max(country_col, fecha_col, punto_col):
            filas_saltadas.append((idx, "fila incompleta"))
            continue
        pais = str(row[country_col]).strip()
        fecha = str(row[fecha_col]).strip()
        punto = str(row[punto_col]).strip()
        if not pais or not fecha or not punto:
            filas_saltadas.append((idx, {"pais": pais, "fecha": fecha, "punto": punto}))
            continue
        pedidos.append((pais, fecha, punto))

    print(f"✅ Pedidos válidos: {len(pedidos)}")
    if filas_saltadas:
        print(f"⚠️ Filas saltadas: {len(filas_saltadas)}")
    return pedidos


def build_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    chrome_bin = os.getenv("GOOGLE_CHROME_BIN", "").strip()
    if chrome_bin:
        options.binary_location = chrome_bin
    return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)


def main():
    enviar_notificacion_slack("El script de entrega fallida de ARG y MX ha comenzado 🚀")

    # Variables
    BACKOFFICE_BASE_URL = os.getenv("BACKOFFICE_BASE_URL", "https://backoffice.nilus.co").rstrip("/")
    BACKOFFICE_EMAIL = get_env("BACKOFFICE_EMAIL")
    BACKOFFICE_PASSWORD = get_env("BACKOFFICE_PASSWORD")

    SHEET_ID = get_env("SHEET_ID")
    GSHEET_WORKSHEET = os.getenv("GSHEET_WORKSHEET", "Fallidas CEs")
    MOTIVO_FIJO = os.getenv("MOTIVO_FIJO", "Error en la entrega")

    # Google Sheets
    gclient = get_gspread_client()
    pedidos_planilla = obtener_pedidos_planilla(
        client=gclient,
        spreadsheet_id=SHEET_ID,
        sheet_name=GSHEET_WORKSHEET,
        country_col=0,
        fecha_col=1,
        punto_col=2,
        header_rows=1
    )

    # Selenium
    driver = build_driver()
    try:
        driver.get(f"{BACKOFFICE_BASE_URL}/es-AR/login")
        time.sleep(5)
        driver.find_element(By.ID, "email").send_keys(BACKOFFICE_EMAIL)
        driver.find_element(By.ID, "password").send_keys(BACKOFFICE_PASSWORD)
        click_button(driver, "//button[text()='INGRESAR']", By.XPATH)
        time.sleep(5)

        country_map = {"AR": "1", "MX": "2"}
        start_time = time.time()
        MAX_DURATION = 60 * 60

        for pais, fecha, punto in pedidos_planilla:
            codigo_pais = country_map.get(str(pais).strip().upper(), "1")
            url = f"{BACKOFFICE_BASE_URL}/es-AR/orders?country={codigo_pais}&status=confirmed"
            print(f"🔗 Navegando a {url} (pais hoja: '{pais}')")
            driver.get(url)
            time.sleep(2)

            if not escribir_punto_de_entrega(driver, punto, 0):
                print(f"❌ No se pudo colocar el punto de entrega para: {punto}. Continuando...")
                continue
            time.sleep(2)

            fecha_normalizada = normalizar_fecha_para_selenium(fecha)
            if not escribir_fecha_entrega(driver, fecha_normalizada, 0):
                print(f"❌ No se pudo escribir la fecha '{fecha_normalizada}' para: {punto}. Continuando...")
                continue
            time.sleep(2)

            while True:
                if time.time() - start_time > MAX_DURATION:
                    print("⏰ Tiempo máximo alcanzado. Finalizando script...")
                    return

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
                        try:
                            cambiar_estado = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='email']"))
                            )
                            cambiar_estado.click()
                            print("✅ Click en combobox id='email'")
                        except Exception:
                            cambiar_estado = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='status']"))
                            )
                            cambiar_estado.click()
                            print("✅ Click en combobox id='status'")

                        cancelado_opcion = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Cancelado')]"))
                        )
                        cancelado_opcion.click()

                        motivo_combo = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "reason_of_canceled"))
                        )
                        motivo_combo.click()
                        time.sleep(2)

                        try:
                            motivo_elemento = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, f"//li[@role='option' and contains(text(), '{MOTIVO_FIJO}')]"))
                            )
                            motivo_elemento.click()
                            print(f"✅ Motivo '{MOTIVO_FIJO}' seleccionado")
                        except Exception as e:
                            print(f"❌ No se pudo seleccionar el motivo '{MOTIVO_FIJO}': {e}")

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
                except Exception as e:
                    print(f"❌ Error en la interacción: {e}")
                time.sleep(5)
        print("✅ Script finalizado correctamente.")
    finally:
        driver.quit()
        enviar_notificacion_slack("Finalizó el script de entrega fallida ✅")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ Error fatal: {e}")
        sys.exit(1)
