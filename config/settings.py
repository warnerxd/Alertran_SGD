# config/settings.py
"""
Configuraciones generales de tiempo y proceso
"""

# Tiempos de espera (en milisegundos)
TIEMPO_ESPERA_RECUPERACION = 4000
TIEMPO_ESPERA_NAVEGACION = 3000
TIEMPO_ESPERA_CLICK = 2000
TIEMPO_ESPERA_CARGA = 8000
TIEMPO_ESPERA_ENTRE_GUIAS = 2000
TIEMPO_ESPERA_INGRESO_CODIGOS = 1500
TIEMPO_ESPERA_VOLVER = 5000

# Configuración de proceso
MAX_REINTENTOS = 3
MAX_NAVEGADORES = 6
URL_ALERTRAN = "https://alertran.latinlogistics.com.co/padua/inicio.do"