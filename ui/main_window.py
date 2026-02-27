# ui/main_window.py
"""
Ventana principal de la aplicación
"""

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QComboBox, QPushButton, QTextEdit, QFileDialog,
    QMessageBox, QGroupBox, QFormLayout, QSpinBox,QDialog,QApplication,
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QTextCursor
from datetime import datetime, timedelta
from pathlib import Path
import os
import subprocess

from ui.login_window import LoginWindow
from ui.resumen_window import ResumenWindow
from ui.historial_window import HistorialWindow
from ui.widgets.progress_bar import MacProgressBar
from workers.proceso_thread import ProcesoThread
from config.constants import CIUDADES, TIPOS_INCIDENCIA, ERROR_MESSAGES
from utils.file_utils import FileUtils

class VentanaPrincipal(QMainWindow):
    """Ventana principal de la aplicación"""
    
    def __init__(self):
        super().__init__()
        self.excel_path = None
        self.proceso_thread = None
        self.sesion_activa = False
        self.usuario_actual = ""
        self.password_actual = ""
        self.historial_datos = []
        self.historial_window = None
        self.tiempo_inicio = None
        self.total_guias = 0
        self.guias_ent = []
        self.guias_error_count = 0
        self.guias_advertencia_count = 0
        self.desviaciones_creadas = 0
        self.carpeta_descargas = FileUtils.obtener_carpeta_descargas()
        
        self._setup_ui()
        self._setup_styles()

    def _setup_ui(self):
        """Configura la interfaz de usuario"""
        self.setWindowTitle("ALERTRAN - Gestión Desviaciones")
        self.setMinimumSize(800, 800)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout_principal = QVBoxLayout(central_widget)
        layout_principal.setSpacing(10)
        layout_principal.setContentsMargins(20, 20, 20, 20)

        layout_principal.addWidget(self._crear_panel_sesion())
        layout_principal.addWidget(self._crear_panel_configuracion())
        layout_principal.addWidget(self._crear_panel_excel())
        layout_principal.addWidget(self._crear_panel_progreso())
        layout_principal.addWidget(self._crear_panel_log())
        layout_principal.addLayout(self._crear_panel_botones())
        
        self._agregar_version_label(layout_principal)
        
        self.showMaximized()

    def _crear_panel_sesion(self):
        grupo = QGroupBox(" 🔐 CONTROL DE ACCESO")
        grupo.setStyleSheet("QGroupBox { font-weight: bold; font-size: 11pt; }")
        
        layout = QHBoxLayout(grupo)
        layout.setSpacing(5)
        
        self.btn_login = QPushButton("🔑 INICIAR SESIÓN")
        self.btn_login.setObjectName("btn_login")
        self.btn_login.clicked.connect(self.abrir_login)
        layout.addWidget(self.btn_login)
        
        self.btn_logout = QPushButton("🚪 CERRAR SESIÓN")
        self.btn_logout.setObjectName("btn_logout")
        self.btn_logout.clicked.connect(self.cerrar_sesion)
        self.btn_logout.setEnabled(False)
        layout.addWidget(self.btn_logout)
        
        self.lbl_estado_sesion = QLabel("⛔ SESIÓN NO INICIADA")
        self.lbl_estado_sesion.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_estado_sesion)
        
        return grupo

    def _crear_panel_configuracion(self):
        grupo = QGroupBox("⚙️ CONFIGURACIÓN")
        grupo.setStyleSheet("QGroupBox { font-weight: bold; font-size: 11pt; }")
        
        layout_config = QFormLayout(grupo)
        layout_config.setSpacing(3)

        self.ciudad_combo = QComboBox()
        self.ciudad_combo.addItems(CIUDADES)
        self.ciudad_combo.setCurrentText("ABA")
        layout_config.addRow("📍 Regional :", self.ciudad_combo)
        
        self.tipo_combo = QComboBox()
        self.tipo_combo.addItems(TIPOS_INCIDENCIA)
        self.tipo_combo.setCurrentText("22")
        layout_config.addRow("📌 desviación :", self.tipo_combo)
        
        self.ampliacion_input = QLineEdit()
        self.ampliacion_input.setPlaceholderText("Amplación Desviación :")
        layout_config.addRow("📝 Ampliación:", self.ampliacion_input)
        
        nav_layout = QHBoxLayout()
        self.num_navegadores_spin = QSpinBox()
        self.num_navegadores_spin.setMinimum(1)
        self.num_navegadores_spin.setMaximum(6)
        self.num_navegadores_spin.setValue(1)
        self.num_navegadores_spin.setPrefix("🚀 ")
        self.num_navegadores_spin.setSuffix(" navegador(es)")
        nav_layout.addWidget(QLabel("Navegadores:"))
        nav_layout.addWidget(self.num_navegadores_spin)
        nav_layout.addStretch()
        layout_config.addRow("", nav_layout)
        
        return grupo

    def _crear_panel_excel(self):
        grupo = QGroupBox("📁 ARCHIVO DE GUÍAS")
        grupo.setStyleSheet("QGroupBox { font-weight: bold; font-size: 11pt; }")
        
        layout_excel = QVBoxLayout(grupo)
        layout_boton_excel = QHBoxLayout()
        
        self.btn_cargar_excel = QPushButton("📂 CARGAR EXCEL")
        self.btn_cargar_excel.setObjectName("btn_cargar_excel")
        self.btn_cargar_excel.clicked.connect(self.cargar_excel)
        self.btn_cargar_excel.setEnabled(False)
        layout_boton_excel.addWidget(self.btn_cargar_excel)
        
        self.lbl_archivo = QLabel("❌ NINGÚN ARCHIVO")
        self.lbl_archivo.setStyleSheet("color: #e74c3c; font-style: italic;")
        layout_boton_excel.addWidget(self.lbl_archivo)
        layout_boton_excel.addStretch()
        
        layout_excel.addLayout(layout_boton_excel)
        return grupo

    def _crear_panel_progreso(self):
        grupo = QGroupBox("📊 PROGRESO")
        grupo.setStyleSheet("QGroupBox { font-weight: bold; font-size: 11pt; }")
        
        layout_progreso = QVBoxLayout(grupo)
        
        self.progress_bar = MacProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout_progreso.addWidget(self.progress_bar)
        
        self.lbl_tiempo_restante = QLabel("")
        self.lbl_tiempo_restante.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_tiempo_restante.setStyleSheet(
            "font-size: 10pt; color: #3498db; font-weight: bold; padding: 5px;"
        )
        layout_progreso.addWidget(self.lbl_tiempo_restante)
        
        self.lbl_estado = QLabel("💤 LISTO")
        self.lbl_estado.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_estado.setStyleSheet("font-weight: bold; font-size: 11pt;")
        layout_progreso.addWidget(self.lbl_estado)
        
        return grupo

    def _crear_panel_log(self):
        grupo = QGroupBox("📋 REGISTRO DE ACTIVIDAD")
        grupo.setStyleSheet("QGroupBox { font-weight: bold; font-size: 11pt; }")
        
        layout_log = QVBoxLayout(grupo)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(150)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #2c3e50;
                font-family: 'Consolas', monospace;
                font-size: 10pt;
                border: 2px solid #34495e;
                border-radius: 5px;
                color: #ecf0f1;
                padding: 8px;
            }
        """)
        layout_log.addWidget(self.log_text)
        
        return grupo

    def _crear_panel_botones(self):
        layout_botones = QHBoxLayout()
        layout_botones.setSpacing(10)
        
        self.btn_iniciar = QPushButton("▶ INICIAR PROCESO")
        self.btn_iniciar.setObjectName("btn_iniciar")
        self.btn_iniciar.clicked.connect(self.iniciar_proceso)
        self.btn_iniciar.setEnabled(False)
        layout_botones.addWidget(self.btn_iniciar)
        
        self.btn_cancelar = QPushButton("⏹ CANCELAR PROCESO")
        self.btn_cancelar.setObjectName("btn_cancelar")
        self.btn_cancelar.clicked.connect(self.cancelar_proceso)
        self.btn_cancelar.setEnabled(False)
        layout_botones.addWidget(self.btn_cancelar)
        
        self.btn_historial = QPushButton("📋 VER HISTORIAL")
        self.btn_historial.setObjectName("btn_historial")  # ← AGREGAR AQUÍ
        self.btn_historial.clicked.connect(self.ver_historial)
        layout_botones.addWidget(self.btn_historial)
        
        self.btn_errores = QPushButton("📥 EXCEL ERRORES")
        self.btn_errores.setObjectName("btn_errores")  # ← AGREGAR AQUÍ
        self.btn_errores.clicked.connect(self.mostrar_errores)
        self.btn_errores.setEnabled(False)
        layout_botones.addWidget(self.btn_errores)
        
        self.btn_descargar_log = QPushButton("💾 DESCARGAR LOG")
        self.btn_descargar_log.setObjectName("btn_descargar_log")  # ← AGREGAR AQUÍ
        self.btn_descargar_log.clicked.connect(self.descargar_log)
        layout_botones.addWidget(self.btn_descargar_log)
        
        return layout_botones

    def _agregar_version_label(self, layout):
        info_label = QLabel("🤖 V.8.0")
        info_label.setStyleSheet("color: #3498db; font-size: 9pt; font-weight: bold;")
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(info_label)

    def _setup_styles(self):
        self.setStyleSheet("""
            QPushButton {
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
                min-width: 150px;
            }
            QPushButton#btn_login {
                background-color: #3498db;
                color: white;
            }
            QPushButton#btn_login:hover {
                background-color: #2980b9;
            }
            QPushButton#btn_logout {
                background-color: #ebc7c7;
                color: white;
            }
            QPushButton#btn_logout:hover {
                background-color: #d35400;
            }
            QPushButton#btn_cargar_excel {
                background-color: #f39c12;
                color: white;
            }
            QPushButton#btn_cargar_excel:hover {
                background-color: #e67e22;
            }
            QPushButton#btn_iniciar {
                background-color: #27ae60;
                color: white;
                font-size: 12pt;
            }
            QPushButton#btn_iniciar:hover {
                background-color: #2ecc71;
            }
            QPushButton#btn_cancelar {
                background-color: #e67e22;
                color: white;
                font-size: 12pt;
            }
            QPushButton#btn_cancelar:hover {
                background-color: #d35400;
            }
            QPushButton#btn_historial {
                background-color: #3498db;
                color: white;
            }
            QPushButton#btn_historial:hover {
                background-color: #2980b9;
            }
            QPushButton#btn_errores {
                background-color: #e74c3c;
                color: white;
            }
            QPushButton#btn_errores:hover {
                background-color: #c0392b;
            }
            QPushButton#btn_descargar_log {
                background-color: #9b59b6;
                color: white;
            }
            QPushButton#btn_descargar_log:hover {
                background-color: #8e44ad;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)
        
        self.btn_login.setObjectName("btn_login")
        self.btn_logout.setObjectName("btn_logout")
        self.btn_cargar_excel.setObjectName("btn_cargar_excel")
        self.btn_iniciar.setObjectName("btn_iniciar")
        self.btn_cancelar.setObjectName("btn_cancelar")
        self.btn_historial.setObjectName("btn_historial")
        self.btn_errores.setObjectName("btn_errores")
        self.btn_descargar_log.setObjectName("btn_descargar_log")

    def abrir_login(self):
        login = LoginWindow(self)
        if login.exec() == QDialog.DialogCode.Accepted:
            usuario, password = login.get_credentials()
            if usuario and password:
                self.usuario_actual = usuario
                self.password_actual = password
                self.sesion_activa = True
                self.actualizar_estado_sesion()
                self.log(f"✅ Sesión iniciada: {usuario}")
                self.habilitar_controles(True)

    def cerrar_sesion(self):
        reply = QMessageBox.question(
            self, "Cerrar Sesión", 
            f"¿Cerrar sesión de {self.usuario_actual}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.sesion_activa = False
            self.usuario_actual = ""
            self.password_actual = ""
            self.actualizar_estado_sesion()
            self.log("🔒 Sesión cerrada")
            self.habilitar_controles(False)
            self.historial_datos.clear()

    def actualizar_estado_sesion(self):
        if self.sesion_activa:
            self.lbl_estado_sesion.setText(f"✅ ACTIVA - {self.usuario_actual}")
            self.lbl_estado_sesion.setStyleSheet("""
                QLabel {
                    background-color: #e8f8f5;
                    color: #27ae60;
                    font-weight: bold;
                    padding: 8px;
                    border-radius: 5px;
                    border: 2px solid #27ae60;
                }
            """)
            self.btn_login.setEnabled(False)
            self.btn_logout.setEnabled(True)
        else:
            self.lbl_estado_sesion.setText("⛔ SESIÓN NO INICIADA")
            self.lbl_estado_sesion.setStyleSheet("""
                QLabel {
                    background-color: #fdeded;
                    color: #e74c3c;
                    font-weight: bold;
                    padding: 8px;
                    border-radius: 5px;
                    border: 2px solid #e74c3c;
                }
            """)
            self.btn_login.setEnabled(True)
            self.btn_logout.setEnabled(False)

    def habilitar_controles(self, habilitar):
        self.btn_cargar_excel.setEnabled(habilitar)
        self.btn_iniciar.setEnabled(habilitar)
        if not habilitar:
            self.excel_path = None
            self.lbl_archivo.setText("❌ NINGÚN ARCHIVO")
            self.lbl_archivo.setStyleSheet("color: #e74c3c; font-style: italic;")
            self.progress_bar.setValue(0)
            self.lbl_estado.setText("💤 LISTO")
            self.lbl_tiempo_restante.setText("")

    def cargar_excel(self):
        archivo, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar Excel", str(Path.home()), "Excel (*.xlsx)"
        )
        if archivo:
            self.excel_path = archivo
            nombre = Path(archivo).name
            
            try:
                guias_count = len(FileUtils.leer_guias_excel(Path(archivo)))
                
                self.lbl_archivo.setText(f"📄 {nombre} ({guias_count} guías)")
                self.lbl_archivo.setStyleSheet("color: #27ae60; font-weight: bold; font-size: 18px;")
                self.log(f"✅ Archivo cargado: {nombre} - {guias_count} guías a procesar")
                
                self.total_guias = guias_count
                
            except Exception as e:
                self.lbl_archivo.setText(f"📄 {nombre} (Error al leer)")
                self.lbl_archivo.setStyleSheet("color: #e74c3c; font-weight: bold;")
                self.log(f"⚠️ Error al leer el archivo: {str(e)}")

    def log(self, mensaje):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{ts}] {mensaje}")
        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.log_text.setTextCursor(cursor)

    def descargar_log(self):
        try:
            contenido_log = self.log_text.toPlainText()
            ruta = FileUtils.guardar_log(contenido_log, self.carpeta_descargas)
            
            QMessageBox.information(
                self, "✅ Éxito", 
                f"📄 Log guardado en:\n{ruta}\n\n📁 Carpeta: Descargas"
            )
            
            self.log(f"✅ Log guardado en: {ruta}")
            
        except Exception as e:
            QMessageBox.critical(self, "❌ Error", f"No se pudo guardar el log:\n{str(e)}")

    def ver_historial(self):
        if not self.historial_window:
            self.historial_window = HistorialWindow(self)
        
        self.historial_window.actualizar_historial(self.historial_datos)
        self.historial_window.show()

    def agregar_al_historial(self, guia, estado, resultado, navegador, fecha):
        self.historial_datos.append((guia, estado, resultado, navegador, fecha))
        
        if len(self.historial_datos) > 1000:
            self.historial_datos = self.historial_datos[-1000:]
        
        if "📦" in estado:
            self.guias_ent.append(guia)
        elif "❌" in estado:
            self.guias_error_count += 1
        elif "⚠️" in estado:
            self.guias_advertencia_count += 1
        elif "✅" in estado:
            self.desviaciones_creadas += 1

    def actualizar_tiempo_restante(self, tiempo):
        self.lbl_tiempo_restante.setText(tiempo)

    def mostrar_resumen(self):
        tiempo_total = datetime.now() - self.tiempo_inicio if self.tiempo_inicio else timedelta(0)
        tiempo_formateado = str(tiempo_total).split('.')[0]
        
        resumen = ResumenWindow(
            total_guias=self.total_guias,
            desviadas=self.desviaciones_creadas,
            entregadas=len(self.guias_ent),
            errores=self.guias_error_count,
            advertencias=self.guias_advertencia_count,
            tiempo_total=tiempo_formateado,
            parent=self
        )
        resumen.exec()

    def iniciar_proceso(self):
        if not self.sesion_activa:
            self._mostrar_error_validacion(ERROR_MESSAGES['NO_SESSION'])
            return
        
        if not self.ampliacion_input.text().strip():
            self.ampliacion_input.setStyleSheet("border: 2px solid red;")
            self._mostrar_error_validacion(ERROR_MESSAGES['NO_AMPLIACION'])
            self.ampliacion_input.setFocus()
            return
        else:
            self.ampliacion_input.setStyleSheet("")
        
        if not self.excel_path:
            self._mostrar_error_validacion(ERROR_MESSAGES['NO_FILE'])
            return

        num_nav = self.num_navegadores_spin.value()
        
        if not self._confirmar_inicio_proceso(num_nav):
            return

        self.tiempo_inicio = datetime.now()
        self.guias_ent = []
        self.guias_error_count = 0
        self.guias_advertencia_count = 0
        self.desviaciones_creadas = 0

        self.btn_iniciar.setEnabled(False)
        self.btn_cancelar.setEnabled(True)
        self.btn_cargar_excel.setEnabled(False)
        self.btn_errores.setEnabled(False)
        self.btn_login.setEnabled(False)
        self.btn_logout.setEnabled(False)
        self.num_navegadores_spin.setEnabled(False)
        self.progress_bar.setValue(0)
        self.lbl_tiempo_restante.setText("⏱️ Calculando tiempo restante...")
        self.log_text.clear()
        self.historial_datos.clear()

        self.log(f"🚀 Iniciando con {num_nav} navegador(es)...")
        self.log(f"👤 Usuario: {self.usuario_actual}")
        self.log(f"📊 Total guías a procesar: {self.total_guias}")
        self.log(f"📁 Los archivos se guardarán en: {self.carpeta_descargas}")

        self.proceso_thread = ProcesoThread(
            self.usuario_actual,
            self.password_actual,
            self.ciudad_combo.currentText(),
            self.tipo_combo.currentText(),
            self.ampliacion_input.text(),
            self.excel_path,
            num_nav
        )

        senales = self.proceso_thread.senales
        senales.progreso.connect(self.progress_bar.setValue)
        senales.estado.connect(self.lbl_estado.setText)
        senales.log.connect(self.log)
        senales.error.connect(self.mostrar_error)
        senales.finalizado.connect(self.proceso_finalizado)
        senales.archivo_errores.connect(self.archivo_errores_generado)
        senales.guia_procesada.connect(self.agregar_al_historial)
        senales.proceso_cancelado.connect(self.proceso_cancelado)
        senales.tiempo_restante.connect(self.actualizar_tiempo_restante)

        self.proceso_thread.start()

    def _mostrar_error_validacion(self, mensaje):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Icon.Warning)
        msg.setWindowTitle("⚠️ Validación")
        msg.setText(f"<b>{mensaje}</b>")
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def _confirmar_inicio_proceso(self, num_nav):
        mensaje = self._crear_mensaje_confirmacion(num_nav)
        
        reply = QMessageBox(self)
        reply.setWindowTitle("🔔 CONFIRMAR PROCESO")
        reply.setText(mensaje)
        reply.setIcon(QMessageBox.Icon.Question)
        
        btn_si = reply.addButton("✅ SÍ, INICIAR", QMessageBox.ButtonRole.YesRole)
        btn_no = reply.addButton("❌ NO, CANCELAR", QMessageBox.ButtonRole.NoRole)
        reply.setDefaultButton(btn_no)
        
        reply.setStyleSheet("""
            QMessageBox QLabel { color: #ffffff; font-size: 10pt; min-width: 400px; }
            QPushButton { padding: 8px 20px; font-weight: bold; }
        """)
        
        reply.exec()
        
        return reply.clickedButton() == btn_si

    def _crear_mensaje_confirmacion(self, num_nav):
        color_guias = "#4CAF50" if self.total_guias < 50 else "#FF9800" if self.total_guias < 100 else "#f44336"
        emoji_guias = "✅" if self.total_guias < 50 else "⚠️" if self.total_guias < 100 else "🔴"
        
        mensaje = f"""
        <div style="font-family: Arial; text-align: center; color: white;">
            <h3 style="color: #ffffff;">📋 RESUMEN DE OPERACIÓN</h3>
            
            <table style="width: 100%; margin: 10px 0; border-collapse: collapse;">
                <tr><td style="padding: 5px; background: #333;">🌐 Navegadores:</td>
                    <td style="padding: 5px; background: #2d2d2d;">{num_nav} simultáneos</td></tr>
                <tr><td style="padding: 5px; background: #333;">👤 Usuario:</td>
                    <td style="padding: 5px; background: #2d2d2d;">{self.usuario_actual}</td></tr>
                <tr><td style="padding: 5px; background: #333;">📋 Total guías:</td>
                    <td style="padding: 5px; background: #2d2d2d;">
                        <span style="background: {color_guias}; color: white; padding: 3px 12px; border-radius: 15px;">
                            {emoji_guias} {self.total_guias}
                        </span>
                    </td></tr>
            </table>
            
            <div style="background: #2a2a2a; padding: 10px; border-radius: 5px; margin: 10px 0;">
                <p>📝 Ampliación N°: {self.ampliacion_input.text().strip()}</p>
                <p style="font-family: monospace; font-size: 8pt;">📂 {self.carpeta_descargas}</p>
            </div>
            
            <p style="margin-top: 15px;"><b>¿Desea continuar con el proceso?</b></p>
        </div>
        """
        
        if self.total_guias > 50:
            mensaje += """
            <div style="background: #331111; padding: 8px; border-radius: 5px; border: 1px solid #ff4444;">
                <p style="color: #ff4444;"><b>⚠️ PROCESO EXTENSO</b><br>Espere sin interrumpir.</p>
            </div>
            """
        
        return mensaje

    def cancelar_proceso(self):
        reply = QMessageBox.question(
            self, "Cancelar Proceso",
            "¿Está seguro que desea cancelar el proceso?\n\nLas guías no procesadas quedarán pendientes.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes and self.proceso_thread:
            self.log("🛑 Cancelando proceso...")
            self.proceso_thread.cancelar()
            self.btn_cancelar.setEnabled(False)
            self.btn_cancelar.setText("⏹ CANCELANDO...")

    def proceso_cancelado(self):
        self.log("✅ Proceso cancelado por usuario")
        self.btn_cancelar.setText("⏹ CANCELAR PROCESO")
        self.lbl_tiempo_restante.setText("")
        self.proceso_finalizado()

    def mostrar_error(self, mensaje):
        QMessageBox.critical(self, "Error", mensaje)
        self.log(f"🔴 ERROR: {mensaje}")
        self.lbl_tiempo_restante.setText("")
        self.proceso_finalizado()

    def proceso_finalizado(self):
        self.btn_iniciar.setEnabled(True)
        self.btn_cancelar.setEnabled(False)
        self.btn_cancelar.setText("⏹ CANCELAR PROCESO")
        self.btn_cargar_excel.setEnabled(True)
        self.btn_login.setEnabled(False)
        self.btn_logout.setEnabled(True)
        self.num_navegadores_spin.setEnabled(True)
        self.lbl_estado.setText("✅ Finalizado")
        
        self.mostrar_resumen()

    def archivo_errores_generado(self, ruta):
        self.btn_errores.setEnabled(True)
        self.error_path = ruta
        
        QMessageBox.information(
            self, "✅ Archivo Generado", 
            f"📊 Archivo de errores y advertencias guardado en:\n{ruta}\n\n📁 Carpeta: Descargas"
        )
        
        self.log(f"✅ Archivo de errores guardado en: {ruta}")

    def mostrar_errores(self):
        if hasattr(self, 'error_path') and self.error_path:
            reply = QMessageBox.question(
                self, "📂 Archivo de Errores",
                f"📊 Archivo guardado en:\n{self.error_path}\n\n"
                f"¿Desea abrir la carpeta que contiene el archivo?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                carpeta = Path(self.error_path).parent
                if os.name == 'nt':
                    os.startfile(carpeta)
                else:
                    subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', str(carpeta)])