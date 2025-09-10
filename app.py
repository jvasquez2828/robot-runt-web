# -*- coding: utf-8 -*-
# ==================================================================================
# === SCRIPT FINAL: VERSIÓN WEB SIMPLIFICADA, ESTABLE Y RÁPIDA (BASADA EN CÓDIGO LOCAL FUNCIONAL) ===
# ==================================================================================
import time
import pandas as pd
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from twocaptcha import TwoCaptcha
import base64
from PIL import Image, ImageEnhance
import io
import os
from datetime import datetime
import asyncio
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from flask import Flask, render_template, send_file, url_for, redirect
import json

# --- CONFIGURACIÓN ---
API_KEY_2CAPTCHA = os.environ.get('API_KEY_2CAPTCHA')
GOOGLE_CREDENTIALS_JSON_STR = os.environ.get('GOOGLE_CREDENTIALS_JSON')
URL_CONSULTA = "https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana"
GOOGLE_SHEET_NAME = "Vehiculos a Consultar RUNT"
MAX_RETRIES = 3
CONCURRENCY_LIMIT = 4

# --- INICIALIZACIÓN DE FLASK ---
app = Flask(__name__)
# Variable para almacenar el nombre del último archivo generado
latest_file = None

# ==================================================================================
# === INICIO: NÚCLEO DEL ROBOT (LÓGICA EXACTA DE TU SCRIPT LOCAL FUNCIONAL) ===
# ==================================================================================
async def consultar_vehiculo(page, placa, num_doc):
    captcha_id = None
    try:
        await page.goto(URL_CONSULTA, wait_until='domcontentloaded', timeout=15000)
        await page.wait_for_selector("//input[@formcontrolname='placa']", timeout=10000)
        
        await page.fill("//input[@formcontrolname='placa']", placa)
        await page.click("//mat-select[@formcontrolname='tipoDocumento']")
        await page.click("//mat-option//span[contains(text(), 'NIT')]")
        await page.fill("//*[@id='mat-input-1']", str(num_doc), timeout=8000)
        
        solver = TwoCaptcha(API_KEY_2CAPTCHA)
        captcha_img_element = page.locator("xpath=//img[contains(@src, 'data:image/png')]")
        screenshot_bytes = await captcha_img_element.screenshot()
        
        image = Image.open(io.BytesIO(screenshot_bytes))
        image = image.convert('L')
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.0)
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(2.0)
        threshold = 150 
        image = image.point(lambda p: 0 if p < threshold else 255)
        
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        base64_captcha = base64.b64encode(buffer.getvalue()).decode('utf-8')
        
        resultado = await asyncio.to_thread(
            solver.normal, base64_captcha, caseSensitive=1, min_len=5, max_len=5, numeric=4, hint_text='The code is 5 characters, case-sensitive.'
        )
        codigo_captcha = resultado['code']
        captcha_id = resultado['captchaId']
        
        await page.fill("//*[@id='mat-input-2']", codigo_captcha)
        await page.press("//*[@id='mat-input-2']", "Enter")

        try:
            error_locator = page.locator("xpath=//div[contains(text(), 'código de verificación es incorrecto')]")
            await error_locator.wait_for(timeout=3500) 
            error_text = await error_locator.inner_text()
            if captcha_id: await asyncio.to_thread(solver.report, captcha_id, False)
            raise Exception("Error de CAPTCHA.")
        except PlaywrightTimeoutError:
            pass

        await page.wait_for_selector("text=Información general del vehículo", timeout=12000)
        
        soat_header_locator = page.locator("xpath=//mat-expansion-panel-header[contains(., 'Póliza SOAT')]")
        await soat_header_locator.click()
        await asyncio.sleep(0.25)
        
        estado_locator = page.locator(f"xpath=//*[@id='cdk-accordion-child-1']/div/mat-card-content/div/mat-table/mat-row[1]/mat-cell[7]")
        texto_completo_soat = (await estado_locator.inner_text(timeout=5000)).strip().lower()
        soat_info = 'Vigente' if 'vigente' in texto_completo_soat and 'no vigente' not in texto_completo_soat else 'No Vigente'
        
        limitaciones_header_locator = page.locator("xpath=//mat-expansion-panel-header[contains(., 'Limitaciones a la Propiedad')]")
        await limitaciones_header_locator.click()
        await asyncio.sleep(0.25)
        
        limitaciones_content_locator = limitaciones_header_locator.locator("xpath=./ancestor::mat-expansion-panel//div[contains(@class, 'mat-expansion-panel-content')]")
        limitaciones_info = (await limitaciones_content_locator.inner_text(timeout=5000)).strip().replace('\n', ' ')
        
        return {"SOAT": soat_info, "Limitaciones": limitaciones_info, "error": None}

    except Exception as e:
        error_msg = str(e).split('\n')[0]
        return {"SOAT": "Error", "Limitaciones": "Error", "error": error_msg}

async def process_vehicle_with_retries(browser, placa, num_doc, semaphore):
    async with semaphore:
        for attempt in range(MAX_RETRIES):
            context = await browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36")
            page = await context.new_page()
            
            async def handle_route(route):
                if route.request.resource_type in ["stylesheet", "font", "media"]: await route.abort()
                elif route.request.resource_type == "image" and not route.request.url.startswith("data:image"): await route.abort()
                else: await route.continue_()
            await page.route("**/*", handle_route)

            resultado = await consultar_vehiculo(page, placa, num_doc)
            await context.close()

            if resultado['error'] is None:
                return {'placa': placa, **resultado}
            else:
                if attempt < MAX_RETRIES - 1: await asyncio.sleep(1.5)
        return {'placa': placa, **resultado}

async def run_main_process():
    global latest_file
    print("Iniciando proceso principal del scraper...")
    try:
        creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON_STR)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict)
        client = gspread.authorize(creds)
        sheet = client.open(GOOGLE_SHEET_NAME).sheet1
        data = sheet.get_all_records()
        df_entrada = pd.DataFrame(data)
        print(f"Se encontraron {len(df_entrada)} registros en Google Sheets.")
        
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            semaphore = asyncio.Semaphore(CONCURRENCY_LIMIT)
            tasks = []
            for _, fila in df_entrada.iterrows():
                tasks.append(process_vehicle_with_retries(browser, fila['placa'], str(fila['numero_documento']), semaphore))
            
            lista_resultados = await asyncio.gather(*tasks)
            await browser.close()

        df_resultados = pd.DataFrame(lista_resultados)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        if not os.path.exists('tmp'): os.makedirs('tmp')
        output_filename = f'tmp/resultados_consulta_{timestamp}.xlsx'
        df_resultados.to_excel(output_filename, index=False)
        
        # Colorear el archivo
        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        wb = load_workbook(output_filename)
        ws = wb.active
        for row in range(2, ws.max_row + 1):
            celda_soat = ws.cell(row=row, column=2)
            celda_limitaciones = ws.cell(row=row, column=3)
            if celda_soat.value and 'Vigente' in str(celda_soat.value): ws.cell(row=row, column=2).fill = green_fill
            elif celda_soat.value and celda_soat.value != 'Error': ws.cell(row=row, column=2).fill = red_fill
            if celda_limitaciones.value and 'no tiene limitaciones a la propiedad' not in str(celda_limitaciones.value).lower() and 'No se encontró' not in str(celda_limitaciones.value): ws.cell(row=row, column=3).fill = red_fill
        wb.save(output_filename)
        
        latest_file = output_filename
        print(f"Proceso completado. Archivo generado: {latest_file}")
        return True
    except Exception as e:
        print(f"Error CRÍTICO en el proceso principal: {e}")
        latest_file = None
        return False
# ==================================================================================
# === FIN: NÚCLEO DEL ROBOT ===
# ==================================================================================

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/run-process")
def run_process_route():
    # El bucle de eventos de asyncio no puede ser reiniciado en el mismo hilo.
    # Corremos el proceso asíncrono y esperamos a que termine.
    success = asyncio.run(run_main_process())
    if success and latest_file:
        # Redirigimos al usuario a la página de descarga.
        return redirect(url_for('download_page'))
    else:
        # Si hay un error, lo mostramos.
        return "Ocurrió un error durante el proceso. Por favor, revisa los logs en Render para más detalles.", 500

@app.route("/download-page")
def download_page():
    # Una página intermedia que le dice al usuario que el archivo está listo.
    return f"""
    <html>
        <head><title>Descarga Lista</title></head>
        <body>
            <h1>¡Proceso Completado!</h1>
            <p>Tu archivo de resultados está listo para descargar.</p>
            <a href="{url_for('download_file_route')}">Haga clic aquí para descargar el archivo Excel</a>
        </body>
    </html>
    """

@app.route("/download")
def download_file_route():
    if latest_file and os.path.exists(latest_file):
        return send_file(latest_file, as_attachment=True)
    return "No se encontró el archivo para descargar o el proceso aún no ha terminado.", 404

# Este bloque es necesario para que Gunicorn (en Render) pueda encontrar la app.
if __name__ == '__main__':
    app.run()

