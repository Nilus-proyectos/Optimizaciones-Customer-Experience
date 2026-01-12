import os
import sys
import json
import ssl
import time

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
from selenium.common.exceptions import TimeoutException


def get_env(name, required=True, default=None):
    value = os.getenv(name, default)
    if required and (value is None or str(value).strip() == ""):
        raise RuntimeError(f"Falta la variable de entorno requerida: {name}")
    return value


def build_ssl_context():
    # Si necesitas un cacert personalizado, define CUSTOM_CACERT_PATH
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


def guardar_cambios(driver, wait_time=10):
    try:
        modal = WebDriverWait(driver, wait_time).until(
            EC.visibility_of_element_located((By.XPATH, "//h2[contains(text(), 'Cambiar el estado del pedido')]/ancestor::div[@role='dialog']"))
        )
        guardar_btn = modal.find_element(By.XPATH, ".//button[normalize-space(text())='Guardar cambios']")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", guardar_btn)
        driver.execute_script("arguments[0].click();", guardar_btn)
        print("✅ 'Guardar cambios' se clickeó correctamente.")
        return True
    except Exception as e:
        print(f"❌ No se pudo hacer clic en el botón correcto: {e}")
        return False


def get_gspread_client():
    # Prefiere JSON en variable GSERVICE_CREDENTIALS_JSON_CONTENT
    json_content = os.getenv("GSERVICE_CREDENTIALS_JSON_CONTENT", "").strip()
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    if json_content:
        try:
            creds_info = json.loads(json_content)
        except json.JSONDecodeError:
            raise RuntimeError("GSERVICE_CREDENTIALS_JSON_CONTENT no es un JSON válido")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        return gspread.authorize(creds)

    # Alternativa opcional: ruta a archivo si la provees
    json_path = os.getenv("GSERVICE_CREDENTIALS_JSON_PATH", "").strip()
    if json_path:
        creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)
        return gspread.authorize(creds)

    raise RuntimeError("Falta GSERVICE_CREDENTIALS_JSON_CONTENT o GSERVICE_CREDENTIALS_JSON_PATH")


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

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    return driver


def main():
    # ======= Variables de entorno requeridas/útiles =======
    SHEET_ID = get_env("SHEET_ID")
    WORKSHEET_NAME = os.getenv("GSHEET_WORKSHEET", "Cancelaciones zona insegura")

    BACKOFFICE_BASE_URL = os.getenv("BACKOFFICE_BASE_URL", "https://backoffice.nilus.co").rstrip("/")
    BACKOFFICE_EMAIL = get_env("BACKOFFICE_EMAIL")
    BACKOFFICE_PASSWORD = get_env("BACKOFFICE_PASSWORD")

    enviar_notificacion_slack("Inicio anulación de ZONA INSEGURA 🇦🇷")

    # Google Sheets
    cliente = get_gspread_client()
    sheet = cliente.open_by_key(SHEET_ID)
    worksheet = sheet.worksheet(WORKSHEET_NAME)
    valores = worksheet.get_all_values()

    # Selenium
    driver = build_driver()

    try:
        # LOGIN
        driver.get(f"{BACKOFFICE_BASE_URL}/es-AR/login")
        time.sleep(5)
        driver.find_element(By.ID, "email").send_keys(BACKOFFICE_EMAIL)
        driver.find_element(By.ID, "password").send_keys(BACKOFFICE_PASSWORD)
        click_button(driver, "//button[text()='INGRESAR']", By.XPATH)
        time.sleep(10)

        for i, row in enumerate(valores[1:], start=1):
            pedido_id = str(row[2]).strip() if len(row) > 2 else ""  # Columna 3
            estado_actual = str(row[4]).strip() if len(row) > 4 else ""  # Columna 5

            if estado_actual == "✅ Hecho":
                print(f"⏭️ Pedido {pedido_id} ya procesado. Saltando fila {i+1}.")
                continue

            if not pedido_id:
                print(f"❌ No se encontró un ID válido en fila {i+1}")
                worksheet.update_cell(i + 1, 5, "❌ ID inválido")
                continue

            print(f"🔄 Procesando pedido {pedido_id} con motivo 'Zona insegura'...")
            try:
                driver.get(f"{BACKOFFICE_BASE_URL}/es-AR/orders/{pedido_id}")
            except Exception:
                print(f"⚠️ Error al abrir el pedido {pedido_id}")
                worksheet.update_cell(i + 1, 5, "❌")
                continue

            try:
                try:
                    cambiar_estado = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='email']"))
                    )
                    cambiar_estado.click()
                except TimeoutException:
                    print("⚠️ No se pudo hacer clic en cambiar estado intento 1. Intentando con id='status'...")
                    try:
                        cambiar_estado = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='status']"))
                        )
                        cambiar_estado.click()
                    except TimeoutException:
                        print("❌ No se pudo hacer clic en cambiar estado intento 2. Continuando con el siguiente pedido...")
                        continue

                cancelado_opcion = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Cancelado')]"))
                )
                cancelado_opcion.click()
                time.sleep(1)

                motivo_combo = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "reason_of_canceled"))
                )
                motivo_combo.click()
                time.sleep(2)
                try:
                    motivos = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Zona insegura')]"))
                    )
                    motivos.click()
                except Exception:
                    print("❌ No se pudo seleccionar el motivo 'Zona insegura'")

                guardar_cambios(driver)
                print(f"��� Pedido {pedido_id} cancelado con motivo 'Zona insegura'")
                worksheet.update_cell(i + 1, 5, "✅ Hecho")
            except Exception:
                print(f"❌ Error procesando pedido {pedido_id}")
                worksheet.update_cell(i + 1, 5, "❌")
                continue
    finally:
        driver.quit()

    enviar_notificacion_slack("Finalizó anulación de ZONA INSEGURA ✅")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ Error fatal: {e}")
        sys.exit(1)
