# models/signals.py
"""
Señales para comunicación entre threads
"""
from PySide6.QtCore import QObject, Signal

class ProcesoSenales(QObject):
    """Señales para el proceso de creación de desviaciones"""
    progreso = Signal(int)
    estado = Signal(str)
    log = Signal(str)
    error = Signal(str)
    finalizado = Signal()
    archivo_errores = Signal(str)
    guia_procesada = Signal(str, str, str, str, str)  # guia, estado, resultado, navegador, fecha
    proceso_cancelado = Signal()
    tiempo_restante = Signal(str)