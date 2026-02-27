# ui/resumen_window.py
"""
Ventana de resumen del proceso
"""
from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame, QWidget
from PySide6.QtGui import QFont
from PySide6.QtCore import Qt

class ResumenWindow(QDialog):
    """Ventana de resumen al finalizar el proceso"""
    
    def __init__(self, total_guias, desviadas, entregadas, errores, advertencias, tiempo_total, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📊 RESUMEN DEL PROCESO")
        self.setMinimumWidth(500)
        self.setMinimumHeight(400)
        self.setModal(True)
        
        self._setup_ui(total_guias, desviadas, entregadas, errores, advertencias, tiempo_total)
        self._setup_styles()

    def _setup_ui(self, total_guias, desviadas, entregadas, errores, advertencias, tiempo_total):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        titulo = QLabel("✅ PROCESO COMPLETADO")
        titulo.setProperty("class", "titulo")
        titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(titulo)
        
        icono = QLabel("🎉")
        icono.setFont(QFont("Arial", 48))
        icono.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(icono)
        
        linea = QFrame()
        linea.setFrameShape(QFrame.Shape.HLine)
        linea.setStyleSheet("background-color: #bdc3c7;")
        layout.addWidget(linea)
        
        grid_layout = QHBoxLayout()
        
        stats = [
            ("TOTAL GUÍAS", str(total_guias), "#3498db"),
            ("DESVIACIONES", str(desviadas), "#27ae60"),
            ("ENTREGADAS", str(entregadas), "#f39c12"),
            ("ERRORES", str(errores), "#e74c3c"),
            ("ADVERTENCIAS", str(advertencias), "#f39c12")
        ]
        
        for titulo_stat, valor, color in stats:
            widget = self._crear_stat_widget(titulo_stat, valor, color)
            grid_layout.addWidget(widget)
        
        layout.addLayout(grid_layout)
        
        tiempo_label = QLabel(f"⏱️ TIEMPO TOTAL: {tiempo_total}")
        tiempo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tiempo_label.setStyleSheet("font-size: 14pt; font-weight: bold; color: #bdc3c7; margin-top: 15px;")
        layout.addWidget(tiempo_label)
        
        btn_cerrar = QPushButton("ACEPTAR")
        btn_cerrar.clicked.connect(self.accept)
        btn_cerrar.setObjectName("btn_cerrar")
        layout.addWidget(btn_cerrar)

    def _crear_stat_widget(self, titulo, valor, color):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        num = QLabel(valor)
        num.setProperty("class", "numero")
        num.setStyleSheet(f"color: {color};")
        num.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(num)
        
        label = QLabel(titulo)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)
        
        return widget

    def _setup_styles(self):
        self.setStyleSheet("""
            QDialog {
                background-color: #212124;
            }
            QLabel {
                font-size: 12pt;
                padding: 5px;
            }
            .titulo {
                font-size: 18pt;
                font-weight: bold;
                color: #2c3e50;
            }
            .numero {
                font-size: 24pt;
                font-weight: bold;
            }
            QPushButton#btn_cerrar {
                background-color: #27ae60;
                color: white;
                font-weight: bold;
                padding: 15px;
                border-radius: 8px;
                font-size: 12pt;
                min-width: 200px;
                margin-top: 20px;
            }
            QPushButton#btn_cerrar:hover {
                background-color: #2ecc71;
            }
        """)