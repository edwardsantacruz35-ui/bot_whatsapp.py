📱 Bot de Cobranza Inteligente vía WhatsApp (Python + Selenium)

📋 Descripción del Proyecto

Sistema de automatización de mensajería (RPA) que conecta una base de datos en Google Sheets con WhatsApp Web. El bot funciona como un asistente virtual de cobranza que monitorea el estado de cuenta de los clientes en tiempo real y gestiona el envío de notificaciones de pago personalizadas.

A diferencia de las listas de difusión tradicionales, este sistema aplica Lógica de Negocio Condicional: respeta horarios de oficina, intervalos de frecuencia para evitar saturación (anti-spam) y realiza conversiones de moneda dinámicas (USD a Bs) basadas en la tasa diaria.

Impacto: Recuperación proactiva de cartera con una tasa de lectura del 98% (WhatsApp) vs 20% (Email tradicional).

⚙️ Arquitectura

El sistema opera mediante un bucle infinito de monitoreo (Daemon Process):

Cerebro (Google Sheets): Actúa como base de datos y panel de control. El bot lee las reglas de negocio (Frecuencia, Hora Mínima, Tasa de Cambio) directamente de la hoja.

Motor (Python): Procesa la lógica de filtrado y decide a quién contactar.

Ejecutor (Selenium): Abre una instancia de Chrome con persistencia de sesión (User Data) para interactuar con WhatsApp Web sin necesidad de escanear el código QR en cada ejecución.

🛠 Stack Tecnológico

Lenguaje: Python 3.x

Base de Datos: Google Sheets API (vía gspread y oauth2client).

Automatización Web: Selenium WebDriver (Chrome).

Persistencia: Chrome User Profiles (para mantener sesión de WhatsApp abierta).

🚀 Características Clave

Persistencia de Sesión: Uso de user-data-dir en Chrome Options para guardar cookies y LocalStorage, evitando el Login/QR repetitivo.

Smart Currency Conversion: Lectura dinámica de la celda K1 en Sheets para calcular el monto en moneda local al momento del envío.

Filtros de "No Molestar":

Frecuencia: Calcula el delta de tiempo desde el último envío (datetime.now() - ultimo_envio).

Horario: Respeta la columna "Hora Mínima" para no escribir a clientes corporativos fuera de horario laboral.

Manejo de Tiempos de Espera: Uso de WebDriverWait y EC.presence_of_element_located para sincronizar la velocidad del script con la carga de la interfaz de WhatsApp (que varía según la conexión).

📦 Instalación y Uso

Habilitar Google Sheets API y descargar credentials.json.

Instalar dependencias: pip install gspread selenium webdriver-manager.

Crear carpeta perfil_chrome en la raíz para guardar la sesión.

Configurar el archivo bot_whatsapp.py (asegúrate de no subir credenciales reales).

Ejecutar: python bot_whatsapp.py.

Escanear el QR una única vez; el sistema recordará la sesión en ejecuciones futuras.

Desarrollado por Edward Gabriel Santacruz - Especialista en Automatización Financiera
