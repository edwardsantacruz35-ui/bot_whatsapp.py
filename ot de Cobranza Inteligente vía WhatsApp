import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import time
import os
import sys

# --- CONFIGURACIÓN ---
NOMBRE_HOJA = "Cobranzas_Bot"
NOMBRE_ARCHIVO_CREDENCIALES = "credentials.json"

# Ubicación de la Tasa de Cambio en la hoja
UBICACION_TASA = "K1" 

# --- MENSAJES PERSONALIZADOS POR NIVEL DE URGENCIA ---
BASE_MENSAJES = {
    1: "Buen día {nombre}, te escribo para validar en relación del pago pendiente del servicio Wawa.",
    2: "Buen día {nombre}, por favor quedamos pendientes en relación al pago.",
    3: "Buen día {nombre}, requerimos respuesta con urgencia en relación al pago."
}

def log(texto):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {texto}")

# --- 1. CONEXIÓN A SHEETS ---
print("--- 🧠 INICIANDO CEREBRO DEL ROBOT ---")
try:
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(NOMBRE_ARCHIVO_CREDENCIALES, scope)
    client = gspread.authorize(creds)
    sheet = client.open(NOMBRE_HOJA).sheet1
    print("✅ Conexión a Base de Datos: EXITOSA")
except Exception as e:
    print(f"❌ Error conectando a Sheets: {e}")
    sys.exit()

# --- BUCLE INFINITO (DAEMON) ---
while True:
    log("🕵️ Analizando clientes en silencio...")
    
    lista_para_enviar = [] 

    try:
        registros = sheet.get_all_records()
        
        # --- LECTURA DE TASA DEL DÍA ---
        try:
            val_tasa = sheet.acell(UBICACION_TASA).value 
            
            if val_tasa:
                tasa_bcv = float(str(val_tasa).replace(",", "."))
            else:
                tasa_bcv = 0
            
            log(f"💰 Tasa BCV leída: {tasa_bcv}") 

        except Exception as e:
            log(f"⚠️ No se pudo leer la tasa: {e}")
            tasa_bcv = 0

    except Exception as e:
        log(f"⚠️ Error leyendo Excel: {e}")
        time.sleep(60)
        continue

    # --- FASE 1: FILTRADO ---
    for i, fila in enumerate(registros, start=2):
        estatus = str(fila.get('Estatus', '')).strip().lower()
        accion = str(fila.get('Accion', '')).strip().lower()

        if estatus == "no pagado" and accion == "enviar whatsapp":
            nombre = fila.get('Nombre', 'Cliente')
            ultimo_envio_str = str(fila.get('Ultimo Envio', ''))
            frecuencia = float(str(fila.get('Frecuencia (Horas)', 24)) or 24)
            hora_minima_str = str(fila.get('Hora Minima', '')).strip()
            
            try: tipo_mensaje = int(fila.get('Tipo Mensaje', 1))
            except: tipo_mensaje = 1
            
            if tipo_mensaje not in BASE_MENSAJES:
                tipo_mensaje = 1

            # 1. Chequeo de Frecuencia (Anti-Spam)
            cumple_frecuencia = True
            if ultimo_envio_str != "":
                try:
                    ultimo_envio = datetime.strptime(ultimo_envio_str, '%Y-%m-%d %H:%M:%S')
                    horas_pasadas = (datetime.now() - ultimo_envio).total_seconds() / 3600
                    cumple_frecuencia = (horas_pasadas >= frecuencia)
                except:
                    cumple_frecuencia = True 

            # 2. Chequeo de Horario Comercial
            cumple_hora = True
            if hora_minima_str: 
                try:
                    hora_pautada = datetime.strptime(hora_minima_str, "%H:%M").time()
                    hora_actual = datetime.now().time()
                    if hora_actual < hora_pautada:
                        cumple_hora = False
                except:
                    log(f"⚠️ Error hora en {nombre}")

            if cumple_frecuencia and cumple_hora:
                lista_para_enviar.append({
                    "fila": i,
                    "nombre": nombre,
                    "telefono": str(fila.get('Telefono', '')),
                    "deuda": float(str(fila.get('Deuda USD', 0)).replace(",", ".") or 0),
                    "mostrar_tasa": str(fila.get('Mostrar Tasa?', 'NO')).upper(),
                    "tipo_msg": tipo_mensaje
                })

    # --- FASE 2: EJECUCIÓN (SELENIUM) ---
    cantidad = len(lista_para_enviar)
    
    if cantidad > 0:
        log(f"🔥 ¡HORA DEL SHOW! {cantidad} mensajes por enviar.")
        
        # Persistencia de Sesión (No escanear QR cada vez)
        ruta_perfil = os.path.join(os.getcwd(), "perfil_chrome")
        options = Options()
        options.add_argument(f"user-data-dir={ruta_perfil}") 
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        
        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            driver.get("https://web.whatsapp.com")
            log("⏳ Esperando WhatsApp (20s)...")
            time.sleep(20)

            for item in lista_para_enviar:
                nombre = item["nombre"]
                deuda = item['deuda']
                tipo = item['tipo_msg']
                
                log(f"🚀 Enviando a: {nombre} (Mensaje Tipo {tipo})...")

                # --- CONSTRUCCIÓN DEL MENSAJE ---
                texto_base = BASE_MENSAJES[tipo]
                mensaje = texto_base.format(nombre=nombre)
                mensaje += f" Saldo pendiente: ${deuda}."

                # Cálculo de conversión dinámica
                if item["mostrar_tasa"] == "SI" and tasa_bcv > 0:
                      monto_bs = round(deuda * tasa_bcv, 2)
                      # Formato moneda local: 1.234,56
                      monto_str = "{:,.2f}".format(monto_bs).replace(",", "X").replace(".", ",").replace("X", ".")
                      mensaje += f" (Tasa BCV {tasa_bcv}: Bs. {monto_str})."
                
                mensaje += " Atentos a su soporte."

                try:
                    # Uso de API URL Scheme para abrir chat directo
                    link = f"https://web.whatsapp.com/send?phone={item['telefono']}&text={mensaje}"
                    driver.get(link)
                    
                    wait = WebDriverWait(driver, 35)
                    # Esperar a que el botón de enviar sea clicable
                    box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')))
                    time.sleep(3)
                    box.send_keys(Keys.ENTER)
                    
                    time.sleep(5) 
                    
                    # Registrar envío en BD
                    fecha_hoy = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    sheet.update_cell(item["fila"], 7, fecha_hoy)
                    log("   ✅ Enviado.")
                    
                    time.sleep(8) # Pausa humana (Anti-ban)

                except Exception as e:
                    log(f"   ❌ Falló envío a {nombre}: {e}")

            log("🏁 Cerrando Chrome...")
            driver.quit()
            
        except Exception as e:
            log(f"❌ Error Chrome: {e}")
            try: driver.quit()
            except: pass
            
    else:
        log("💤 Nada que enviar por ahora.")

    # Esperar 10 minutos para el siguiente ciclo
    time.sleep(600)
