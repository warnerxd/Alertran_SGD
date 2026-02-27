# ui/login_window.py
"""
Ventana de inicio de sesión
"""
from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QFormLayout, QWidget
from PySide6.QtGui import QFont
from PySide6.QtCore import Qt

class LoginWindow(QDialog):
    """Ventana de login de ALERTRAN"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🔐 Iniciar Sesión - ALERTRAN")
        self.setMinimumWidth(400)
        self.setModal(True)
        
        self._setup_ui()
        self._setup_styles()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        titulo = QLabel("🔐 INICIAR SESIÓN")
        titulo.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(titulo)
        
        form_widget = QWidget()
        form_layout = QFormLayout(form_widget)
        form_layout.setSpacing(15)
        form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        
        self.usuario_input = QLineEdit()
        self.usuario_input.setPlaceholderText("Ingrese su usuario")
        self.usuario_input.setMinimumHeight(35)
        form_layout.addRow("👤 Usuario:", self.usuario_input)
        
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setPlaceholderText("Ingrese su contraseña")
        self.password_input.setMinimumHeight(35)
        form_layout.addRow("🔒 Contraseña:", self.password_input)
        
        layout.addWidget(form_widget)
        
        button_layout = QHBoxLayout()
        button_layout.setSpacing(15)
        
        self.btn_login = QPushButton("✅ INICIAR SESIÓN")
        self.btn_login.clicked.connect(self.accept)
        
        self.btn_cancel = QPushButton("❌ CANCELAR")
        self.btn_cancel.clicked.connect(self.reject)
        
        button_layout.addWidget(self.btn_login)
        button_layout.addWidget(self.btn_cancel)
        
        layout.addLayout(button_layout)
        
        self.usuario_input.returnPressed.connect(self.password_input.setFocus)
        self.password_input.returnPressed.connect(self.accept)

    def _setup_styles(self):
        self.setStyleSheet("""
            QDialog {
                background-color:#212124;
                border-radius: 10px;
            }
            QLineEdit {
                padding: 8px;
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                font-size: 11pt;
            }
            QLineEdit:focus {
                border-color: #3498db;
            }
            QLabel {
                font-size: 11pt;
            }
            QPushButton {
                font-weight: bold;
                padding: 12px;
                border-radius: 5px;
                font-size: 11pt;
                min-width: 150px;
            }
            QPushButton#btn_login {
                background-color: #27ae60;
                color: white;
            }
            QPushButton#btn_login:hover {
                background-color: #2ecc71;
            }
            QPushButton#btn_cancel {
                background-color: #ebc7c7;
                color: #616060;
            }
            QPushButton#btn_cancel:hover {
                background-color: #e84646;
            }
        """)
        
        self.btn_login.setObjectName("btn_login")
        self.btn_cancel.setObjectName("btn_cancel")
    
    def get_credentials(self):
        return self.usuario_input.text(), self.password_input.text()