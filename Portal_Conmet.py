import os
import time
import re
import sqlite3
import smtplib
from datetime import datetime, timedelta
import pandas as pd
from email.message import EmailMessage
from dotenv import load_dotenv
from imap_tools import MailBox, A
from playwright.sync_api import Playwright, sync_playwright

# ==========================================
# 1. CONFIGURACIÓN GLOBAL Y CREDENCIALES
# ==========================================
load_dotenv()
USER_VERTICAL = os.getenv("USER_VERTICAL")
PASSWORD_VERTICAL = os.getenv("PASSWORD_VERTICAL")
IMAP_USER = os.getenv("IMAP_USER")
IMAP_PASSWORD = os.getenv("IMAP_PASSWORD")
USER_CLIENTE = os.getenv("USER_CLIENTE")
PASSWORD_CLIENTE = os.getenv("PASSWORD_CLIENTE")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))

# Actualizado con el dominio @grupoei.com.mx
DESTINATARIOS_BLOQUEO = ["vcruz.mty@aduax.com", "jtrujillo.mty@aduax.com"]
DESTINATARIO_RESUMEN = [IMAP_USER] 

# Nombre unificado de la base de datos
DB_NAME = 'historial_cargas_Conmet.db'

# ==========================================
# 2. FUNCIONES CENTRALIZADAS DE BASE DE DATOS
# ==========================================
def inicializar_base_datos():
    """Crea la base de datos SQLite y la tabla completa si no existen."""
    print("Verificando/Creando base de datos unificada...")
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historial (
            factura TEXT PRIMARY KEY,
            fecha_factura TEXT,
            fecha_carga TEXT,
            monto_facturado REAL,
            pedimento TEXT,
            estatus TEXT,
            detalle_error TEXT,
            folio_operacion TEXT
        )
    ''')
    # Por si migras desde la BD vieja, aseguramos que existan las columnas nuevas
    for col in ["detalle_error", "folio_operacion"]:
        try:
            cursor.execute(f"ALTER TABLE historial ADD COLUMN {col} TEXT")
        except sqlite3.OperationalError:
            pass
    conn.commit()
    conn.close()

# --- Funciones BD Fase 2 ---
def registrar_excepcion_bd(factura, fecha, total, pedimento, detalle):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    estatus_principal = "Carga Pendiente" 
    cursor.execute('''
        INSERT INTO historial (factura, fecha_factura, monto_facturado, pedimento, estatus, detalle_error)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(factura) DO UPDATE SET estatus=excluded.estatus, detalle_error=excluded.detalle_error
    ''', (factura, fecha, total, pedimento, estatus_principal, detalle))
    conn.commit()
    conn.close()

def registrar_pendiente_bd(factura, fecha, total, pedimento):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    estatus_principal = "Pendiente"
    detalle = "Archivos listos en carpeta"
    cursor.execute('''
        INSERT INTO historial (factura, fecha_factura, monto_facturado, pedimento, estatus, detalle_error)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(factura) DO UPDATE SET estatus=excluded.estatus, detalle_error=excluded.detalle_error
    ''', (factura, fecha, total, pedimento, estatus_principal, detalle))
    conn.commit()
    conn.close()

def obtener_facturas_exitosas():
    conn = sqlite3.connect(DB_NAME)
    try:
        df_exitosas = pd.read_sql("SELECT factura FROM historial WHERE estatus = 'Ok cargada'", conn)
        exitosas = df_exitosas['factura'].tolist()
    except Exception:
        exitosas = []
    conn.close()
    return exitosas

# --- Funciones BD Fase 3 ---
def obtener_facturas_pendientes():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT factura, monto_facturado FROM historial WHERE estatus != 'Ok cargada' OR estatus IS NULL")
    pendientes = cursor.fetchall()
    conn.close()
    return pendientes 

def actualizar_estatus_masivo(pendientes, estatus, detalle):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    for factura, _ in pendientes:
        cursor.execute('''
            UPDATE historial SET estatus = ?, detalle_error = ?, fecha_carga = datetime('now', 'localtime')
            WHERE factura = ?
        ''', (estatus, detalle, factura))
    conn.commit()
    conn.close()

def actualizar_estatus_individual(factura, estatus, detalle, folio="N/A"):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE historial SET estatus = ?, detalle_error = ?, folio_operacion = ?, fecha_carga = datetime('now', 'localtime')
        WHERE factura = ?
    ''', (estatus, detalle, folio, factura))
    conn.commit()
    conn.close()

# ==========================================
# 3. FUNCIONES DE CORREO Y REPORTES (FASE 3)
# ==========================================
def enviar_correo(asunto, cuerpo_html, destinatarios, df_adjunto=None):
    try:
        msg = EmailMessage()
        msg['Subject'] = asunto
        msg['From'] = IMAP_USER
        msg['To'] = ", ".join(destinatarios)
        msg.set_content("Tu cliente de correo no soporta HTML. Revisa el adjunto.")
        msg.add_alternative(cuerpo_html, subtype='html')

        if df_adjunto is not None and not df_adjunto.empty:
            archivo_temp = "Reporte_Ejecucion.xlsx"
            df_adjunto.to_excel(archivo_temp, index=False)
            with open(archivo_temp, 'rb') as f:
                excel_data = f.read()
            msg.add_attachment(excel_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=archivo_temp)
            os.remove(archivo_temp)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(IMAP_USER, IMAP_PASSWORD)
            server.send_message(msg)
        print(f"📧 Correo enviado exitosamente a {destinatarios}")
    except Exception as e:
        print(f"❌ Error al enviar correo: {e}")

def procesar_resumen_y_enviar(resultados):
    total_intentos = len(resultados)
    exitosas = [r for r in resultados if r['Estatus Principal'] == 'Ok cargada']
    errores = [r for r in resultados if r['Estatus Principal'] == 'Carga Pendiente']
    
    monto_exito = sum(r['Monto'] for r in exitosas)
    monto_error = sum(r['Monto'] for r in errores)

    cuerpo_html = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <h2 style="color: #2e6c80;">Resumen de Ejecución Diaria: CONMET</h2>
        <p>Hola, la ejecución del robot ha finalizado. Aquí están los resultados <b>EXCLUSIVOS</b> de las facturas procesadas hoy:</p>
        <ul>
            <li><b>Total Intentos de Carga Hoy:</b> {total_intentos}</li>
            <li><span style="color: green;"><b>Cargas Exitosas:</b></span> {len(exitosas)} (Monto: ${monto_exito:,.2f})</li>
            <li><span style="color: red;"><b>Cargas Pendientes / Problemas:</b></span> {len(errores)} (Monto: ${monto_error:,.2f})</li>
        </ul>
        <p>Se adjunta el archivo Excel únicamente con el detalle de las facturas intentadas en este bloque.</p>
    </body>
    </html>
    """
    df_resumen = pd.DataFrame(resultados)
    enviar_correo("Resumen de Ejecución Diaria: CONMET", cuerpo_html, DESTINATARIO_RESUMEN, df_adjunto=df_resumen)

def reiniciar_navegacion_carga(page):
    """Fuerza al portal a regresar al menú inicial de carga para limpiar errores previos."""
    page.get_by_role("link", name="Subir Múltiples Facturas").click()
    page.wait_for_timeout(2000)
    page.locator(".k-icon").first.click()
    page.get_by_role("option", name="CONMET DE MEXICO").click()
    page.wait_for_timeout(1000)
    page.get_by_role("button", name="select").click()
    page.get_by_role("option", name="EMPRESA DE LOGISTICA").click()
    page.wait_for_timeout(1000)
    page.get_by_role("link", name="SIGUIENTE").click()
    page.wait_for_timeout(2000)

# ==========================================
# 4. LÓGICA DE LAS FASES (MÓDULOS)
# ==========================================
def fase1_descarga_reporte(playwright: Playwright) -> None:
    fecha_final = datetime.now()
    fecha_inicial = fecha_final - timedelta(days=7)
    str_fecha_inicial = fecha_inicial.strftime("%Y-%m-%d")
    str_fecha_final = fecha_final.strftime("%Y-%m-%d")
    
    print(f"Periodo de consulta: {str_fecha_inicial} a {str_fecha_final}")

    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    print("Navegando a Vertical...")
    page.goto("https://web.aduax.com/vertical/login.html")
    page.wait_for_timeout(2000)

    page.get_by_role("textbox", name="Usuario").fill(USER_VERTICAL)
    page.wait_for_timeout(1000)
    page.get_by_role("textbox", name="Contraseña").fill(PASSWORD_VERTICAL)
    page.wait_for_timeout(1000)
    page.get_by_role("button", name="Entrar").click()
    page.wait_for_timeout(3000)

    print("Seleccionando empresa ELI y navegando...")
    page.locator("#ddlEmpresas").select_option("1")
    page.wait_for_timeout(2000)

    page.get_by_role("link", name="Facturacion").click()
    page.wait_for_timeout(1000)
    page.get_by_role("link", name="Facturacion").click()
    page.wait_for_timeout(1000)
    page.get_by_role("link", name="Facturacion").click()
    page.wait_for_timeout(1000)
    page.locator("#ul361 a").filter(has_text="Documentos por Referencia").click()
    page.wait_for_timeout(3000)

    frame = page.get_by_text("</div> </div> </body> </html>").content_frame
    
    print("Llenando formulario de búsqueda...")
    frame.get_by_text("Cliente").click()
    page.wait_for_timeout(1000)
    
    frame.get_by_placeholder("Inicial").fill(str_fecha_inicial)
    page.wait_for_timeout(1000)
    frame.get_by_placeholder("Final").fill(str_fecha_final)
    page.wait_for_timeout(1000)
    
    frame.get_by_role("textbox", name="Cliente").click()
    frame.get_by_role("textbox", name="Cliente").fill("9826-CONMET DE MEXICO SA DE CV")
    page.wait_for_timeout(1000)
    
    frame.get_by_role("button", name=" Buscar").click()
    
    print("Buscando información... (Esperando 5 segundos)")
    page.wait_for_timeout(5000)

    print("Descargando archivo Excel...")
    with page.expect_download() as download_info:
        frame.locator("button[onclick='Exportar();']").click()
    
    download = download_info.value
    ruta_descarga = f"reporte_{str_fecha_final}.xlsx"
    download.save_as(ruta_descarga)
    print(f"✅ Excel descargado en: {ruta_descarga}")

    context.close()
    browser.close()

    print("Iniciando limpieza de datos...")
    df = pd.read_excel(ruta_descarga, header=1)
    
    columnas_indices = [1, 4, 6, 11, 23]
    df_filtrado = df.iloc[:, columnas_indices].copy()
    df_filtrado.columns = ['Factura', 'Fecha', 'UUID', 'Total', 'Pedimento']
    df_filtrado = df_filtrado.dropna(subset=['UUID'])
    df_limpio = df_filtrado.drop_duplicates(subset=['Factura'], keep='first')
    
    print(f"✅ Limpieza finalizada. {len(df_limpio)} facturas válidas encontradas.")
    df_limpio.to_excel("datos_limpios_fase1.xlsx", index=False)

def fase2_procesar_imap():
    if not os.path.exists("datos_limpios_fase1.xlsx"):
        print("❌ Error: No se encontró 'datos_limpios_fase1.xlsx'.")
        return
        
    df = pd.read_excel("datos_limpios_fase1.xlsx")
    facturas_exitosas = obtener_facturas_exitosas()
    df = df[~df['Factura'].isin(facturas_exitosas)]
    
    print(f"📦 Cargadas {len(df)} facturas pendientes para buscar en IMAP.\n")
    if df.empty:
        print("✅ Todas las facturas de este periodo ya están cargadas con éxito.")
        return

    print("Conectando al servidor IMAP...")
    try:
        with MailBox('imap.gmail.com', port=993).login(IMAP_USER, IMAP_PASSWORD) as mailbox:
            print("✅ Conexión IMAP exitosa.")
            mailbox.folder.set('PA CONMET')
            
            for index, row in df.iterrows():
                factura = row['Factura']
                fecha = row['Fecha']
                total = row['Total']
                pedimento = row['Pedimento']
                
                print(f"----------------------------------------")
                print(f"🔍 Buscando Factura: {factura} ...")
                
                correos = list(mailbox.fetch(A(subject=factura)))
                
                if not correos:
                    detalle = "No se encontro la factura en correo"
                    print(f"⚠️ {detalle}: {factura}")
                    registrar_excepcion_bd(factura, fecha, total, pedimento, detalle)
                    continue
                
                archivos_descargados = 0
                
                for correo in correos:
                    for att in correo.attachments:
                        nombre_archivo = att.filename
                        es_pdf_o_xml = nombre_archivo.lower().endswith('.pdf') or nombre_archivo.lower().endswith('.xml')
                        contiene_folio = factura.lower() in nombre_archivo.lower()
                        
                        if es_pdf_o_xml and contiene_folio:
                            ruta_base = os.path.join("Descargas")
                            os.makedirs(ruta_base, exist_ok=True)
                            ruta_factura = os.path.join(ruta_base, factura)
                            os.makedirs(ruta_factura, exist_ok=True)
                            ruta_guardado = os.path.join(ruta_factura, nombre_archivo)
                            
                            with open(ruta_guardado, 'wb') as f:
                                f.write(att.payload)
                            
                            print(f"  📥 Descargado: {ruta_guardado}")
                            archivos_descargados += 1
                
                if archivos_descargados == 0:
                    detalle = "Correo encontrado, pero sin adjuntos validos"
                    print(f"⚠️ {detalle} para {factura}.")
                    registrar_excepcion_bd(factura, fecha, total, pedimento, detalle)
                else:
                    print(f"✅ Descarga completada para {factura}.")
                    registrar_pendiente_bd(factura, fecha, total, pedimento)

    except Exception as e:
        print(f"❌ Ocurrió un error con el IMAP: {e}")

def fase3_carga_portal(playwright: Playwright) -> None:
    facturas_pendientes = obtener_facturas_pendientes()
    
    if not facturas_pendientes:
        print("No hay facturas pendientes por cargar.")
        return

    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    
    print("Navegando al portal del cliente...")
    page.goto("https://apconmet.azurewebsites.net/login.aspx")
    
    page.get_by_role("textbox", name="Usuario").fill(USER_CLIENTE)
    page.get_by_role("textbox", name="Contraseña").fill(PASSWORD_CLIENTE)
    page.get_by_role("link", name="ENTRAR").click()
    
    page.wait_for_timeout(3000)
    icono_cerrar = page.get_by_text("")
    if icono_cerrar.is_visible():
        icono_cerrar.click()

    page.wait_for_timeout(1000)
    page.get_by_role("button", name=" Facturas").click()
    page.wait_for_timeout(500)
    
    reiniciar_navegacion_carga(page)
    
    # === 1. BLOQUEO POR COMPLEMENTO DE PAGO ===
    mensaje_bloqueo_pago = page.get_by_text("presenta un bloqueo por tener pagos sin Complemento")
    if mensaje_bloqueo_pago.is_visible():
        detalle = "Faltan complementos de pago en el portal, solicitar a cobranza"
        actualizar_estatus_masivo(facturas_pendientes, "Carga Pendiente", detalle)
        cuerpo_bloqueo = "<p>El portal Conmet se encuentra bloqueado para la recepcion de facturas, por favor cargar el complemento de pago pendiente.</p>"
        enviar_correo("Bloqueo de portal Conmet - Complemento de Pago", cuerpo_bloqueo, DESTINATARIOS_BLOQUEO)
        context.close()
        browser.close()
        return 
        
    # === 2. BLOQUEO POR CARTA DE OPINIÓN ===
    mensaje_bloqueo_carta = page.get_by_text("Carta de Opinión de Obligaciones Fiscales")
    if mensaje_bloqueo_carta.is_visible():
        detalle = "Falta actualizar Carta de Opinión de Obligaciones Fiscales"
        actualizar_estatus_masivo(facturas_pendientes, "Carga Pendiente", detalle)
        cuerpo_bloqueo = "<p>El portal Conmet se encuentra bloqueado para la recepción de facturas. <b>Motivo:</b> La razón social seleccionada presenta un bloqueo por no tener actualizada la Carta de Opinión de Obligaciones Fiscales en el portal.</p>"
        enviar_correo("Bloqueo de portal Conmet - Carta de Opinión", cuerpo_bloqueo, DESTINATARIOS_BLOQUEO)
        context.close()
        browser.close()
        return 

    print("✅ Cuenta libre. Iniciando ciclo de carga...")
    resultados_ejercicio = []
    
    for i, (factura, monto) in enumerate(facturas_pendientes):
        print(f"\n--- Procesando factura {i+1}/{len(facturas_pendientes)}: {factura} ---")
        
        ruta_xml = os.path.join("Descargas", factura, f"{factura}.xml")
        ruta_pdf = os.path.join("Descargas", factura, f"{factura}.pdf")
        
        if not os.path.exists(ruta_xml) or not os.path.exists(ruta_pdf):
            detalle = "Archivos no encontrados en carpeta Descargas"
            actualizar_estatus_individual(factura, "Carga Pendiente", detalle)
            resultados_ejercicio.append({"Factura": factura, "Monto": monto, "Estatus Principal": "Carga Pendiente", "Detalle": detalle, "Folio Operacion": "N/A"})
            continue

        try:
            print("Cargando XML...")
            with page.expect_file_chooser() as fc_info:
                page.locator("a:has(img[src*='xml_activo.svg'])").click()
            fc_info.value.set_files(ruta_xml)
            
            page.wait_for_timeout(5000) 
            
            print("Cargando PDF...")
            with page.expect_file_chooser() as fc_info:
                page.locator("a:has(img[src*='pdf_activo.svg'])").click()
            fc_info.value.set_files(ruta_pdf)
            
            page.wait_for_timeout(4000)
            
            page.locator("#btnAvanzarPaso2").click()
            page.wait_for_timeout(3000)
            
            errores = page.locator("#lblFacturasConProblemas").inner_text(timeout=5000)
            if errores != "0":
                if page.get_by_text("El uuid ya ha sido usado").first.is_visible():
                    detalle = "Cargada por intervencion manual"
                    print(f"✅ {detalle} (UUID duplicado detectado).")
                    actualizar_estatus_individual(factura, "Ok cargada", detalle, "N/A")
                    resultados_ejercicio.append({"Factura": factura, "Monto": monto, "Estatus Principal": "Ok cargada", "Detalle": detalle, "Folio Operacion": "N/A"})
                else:
                    detalle = "Error de validación en portal (Revisar CFDI)"
                    print(f"⚠️ {detalle}")
                    actualizar_estatus_individual(factura, "Carga Pendiente", detalle)
                    resultados_ejercicio.append({"Factura": factura, "Monto": monto, "Estatus Principal": "Carga Pendiente", "Detalle": detalle, "Folio Operacion": "N/A"})
                
                if i < len(facturas_pendientes) - 1:
                    print("Limpiando estado del portal tras validación...")
                    reiniciar_navegacion_carga(page)
                continue

            page.get_by_role("link", name="FINALIZAR").click()
            page.wait_for_timeout(3000)
            
            texto_exito = page.locator("body").inner_text()
            match_folio = re.search(r"Folio Operación:\s*(\d+)", texto_exito)
            folio_operacion = match_folio.group(1) if match_folio else "Desconocido"
            
            detalle = "Portal genero folio"
            actualizar_estatus_individual(factura, "Ok cargada", detalle, folio_operacion)
            resultados_ejercicio.append({"Factura": factura, "Monto": monto, "Estatus Principal": "Ok cargada", "Detalle": detalle, "Folio Operacion": folio_operacion})
            
            if i < len(facturas_pendientes) - 1:
                page.get_by_text("SUBIR NUEVA FACTURA").click()
                page.wait_for_timeout(2000)
                page.get_by_role("link", name="SIGUIENTE").click()
                page.wait_for_timeout(2000)
                
        except Exception as e:
            detalle = "Interrupción web (Timeout, elemento no encontrado o bloqueo)"
            print(f"❌ Ocurrió una interrupción crítica: {e}")
            actualizar_estatus_individual(factura, "Carga Pendiente", detalle)
            resultados_ejercicio.append({"Factura": factura, "Monto": monto, "Estatus Principal": "Carga Pendiente", "Detalle": detalle, "Folio Operacion": "N/A"})
            
            if i < len(facturas_pendientes) - 1:
                print("Intentando recuperar el estado del portal para la siguiente factura...")
                try:
                    reiniciar_navegacion_carga(page)
                except Exception as reset_error:
                    print("❌ El portal se ha bloqueado irremediablemente. Terminando proceso de carga actual.")
                    break 

    print("\nEnviando reporte de resumen...")
    procesar_resumen_y_enviar(resultados_ejercicio)
    context.close()
    browser.close()

# ==========================================
# 5. CONTROLADOR PRINCIPAL
# ==========================================
def main():
    print("=== INICIANDO AUTOMATIZACIÓN CONMET ===")
    inicializar_base_datos()
    
    print("\n--- INICIANDO FASE 1: DESCARGA DE REPORTE ---")
    with sync_playwright() as playwright:
        fase1_descarga_reporte(playwright)
        
    print("\n--- INICIANDO FASE 2: PROCESAMIENTO IMAP ---")
    fase2_procesar_imap()
    
    print("\n--- INICIANDO FASE 3: CARGA AL PORTAL ---")
    with sync_playwright() as playwright:
        fase3_carga_portal(playwright)
        
    print("\n=== AUTOMATIZACIÓN COMPLETADA CON ÉXITO ===")

if __name__ == "__main__":
    main()
