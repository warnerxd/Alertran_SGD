##ALERTRAN_SGD V.8.0
##Cualquier Pull Request notificar por teams para pronta respuesta eduar fabian vargas

##importacion de librerias IMPORTANTE EN ENTORNO EMPRESARIAN RESTRINGE PANDAS.

import sys
import asyncio
from PySide6.QtWidgets import QApplication
import qasync

from ui.main_window import VentanaPrincipal

def main():
    """Función principal"""
    app = QApplication(sys.argv)
    loop = qasync.QEventLoop(app)
    asyncio.set_event_loop(loop)
    app.setStyle('Fusion')
    
    ventana = VentanaPrincipal()
    ventana.show()
    
    with loop:
        sys.exit(loop.run_forever())

if __name__ == "__main__":
    main()