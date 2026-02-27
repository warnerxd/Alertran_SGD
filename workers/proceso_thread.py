# workers/proceso_thread.py
"""
Thread principal para el procesamiento con múltiples navegadores
"""
from PySide6.QtCore import QThread
from datetime import datetime, timedelta
import asyncio
import time
from playwright.async_api import async_playwright
from typing import List, Union
from pathlib import Path

from models.signals import ProcesoSenales
from utils.file_utils import FileUtils
from config.settings import (
    MAX_REINTENTOS, TIEMPO_ESPERA_CLICK, TIEMPO_ESPERA_NAVEGACION,
    TIEMPO_ESPERA_INGRESO_CODIGOS, TIEMPO_ESPERA_VOLVER, URL_ALERTRAN
)

class ProcesoThread(QThread):
    """Thread principal para el procesamiento"""
    
    def __init__(self, usuario, password, ciudad, tipo, ampliacion, excel_path, num_navegadores):
        super().__init__()
        self.usuario = usuario
        self.password = password
        self.ciudad = ciudad
        self.tipo = tipo
        self.ampliacion = ampliacion
        self.excel_path = excel_path
        self.num_navegadores = min(num_navegadores, 6)
        self.senales = ProcesoSenales()
        
        # Estado del proceso
        self.guias_error = []
        self.guias_advertencia = []
        self.guias_ent = []
        self.guias_procesadas_exito = set()
        self.guias_procesadas_ent = set()
        self.guias_en_error = set()
        
        # Recursos
        self.pages = []
        self.browsers = []
        self.contexts = []
        self.lock = asyncio.Lock()
        self.cola_guias = []
        
        # Control
        self.procesando = True
        self.cancelado = False
        self.tiempo_inicio = None
        self.total_guias = 0
        self.carpeta_descargas = FileUtils.obtener_carpeta_descargas()
        self.file_utils = FileUtils()

    def leer_excel(self, ruta: Union[str, Path]) -> List[str]:
        """Lee el archivo Excel y extrae las guías"""
        return self.file_utils.leer_guias_excel(Path(ruta))

    async def esperar_overlay(self, page, timeout=10000):
        """Espera a que desaparezca el overlay"""
        try:
            await page.wait_for_selector("#capa_selector", state="hidden", timeout=timeout)
        except:
            pass
        await asyncio.sleep(1.5)

    async def verificar_pagina_activa(self, page):
        """Verifica si la página está activa"""
        try:
            await page.title()
            return True
        except:
            return False

    async def verificar_estado_ent(self, page, nav_idx):
        """Verifica si la guía tiene estado ENT"""
        try:
            elemento_ent = (
                page
                .frame_locator("frame[name=\"menu\"]")
                .frame_locator("iframe[name=\"principal\"]")
                .frame_locator("frame[name=\"resultado\"]")
                .get_by_role("cell", name="ENT", exact=True)
            )
            
            if await elemento_ent.count() > 0:
                self.senales.log.emit(f"📦 [Nav{nav_idx}] Estado ENT detectado")
                return True
            return False
        except:
            return False

    async def calcular_tiempo_restante(self, procesadas, total):
        """Calcula y emite el tiempo restante"""
        if self.tiempo_inicio and procesadas > 0:
            elapsed = time.time() - self.tiempo_inicio
            velocidad = procesadas / elapsed if elapsed > 0 else 0
            if velocidad > 0:
                restantes = total - procesadas
                segundos_restantes = restantes / velocidad
                tiempo_restante = str(timedelta(seconds=int(segundos_restantes)))
                self.senales.tiempo_restante.emit(f"⏱️ Tiempo restante: {tiempo_restante}")

    async def hacer_login(self, page, nav_idx):
        """Realiza el login en ALERTRAN"""
        try:
            self.senales.log.emit(f"🔐 [Nav{nav_idx}] Iniciando sesión...")
            await page.fill('input[name="j_username"]', self.usuario)
            await asyncio.sleep(0.2)
            await page.fill('input[name="j_password"]', self.password)
            await asyncio.sleep(0.2)
            await page.get_by_role("button", name="Aceptar").click()
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(1)
            return True
        except Exception as e:
            self.senales.log.emit(f"❌ [Nav{nav_idx}] Error login: {str(e)}")
            return False

    async def navegar_a_funcionalidad_7_8(self, page, nav_idx):
        """Navega a la funcionalidad 7.8"""
        try:
            self.senales.log.emit(f"🧭 [Nav{nav_idx}] Navegando a 7.8...")
            
            if not await self.verificar_pagina_activa(page):
                return False
            
            menu = page.frame_locator('frame[name="menu"]')
            
            try:
                base_selector = menu.get_by_role("cell", name="ABA BARRANQUILLA AEROPUE").locator("span")
                if await base_selector.count() > 0:
                    await base_selector.click(timeout=3000)
            except:
                pass

            try:
                ciudad_selector = menu.get_by_role("list").get_by_text(self.ciudad)
                if await ciudad_selector.count() > 0:
                    await ciudad_selector.click(timeout=2000)
                    await asyncio.sleep(TIEMPO_ESPERA_CLICK * 2 / 1000)
            except:
                pass

            funcionalidad = menu.locator('input[name="funcionalidad_codigo"]:not([type="hidden"])')
            await funcionalidad.wait_for(state="visible", timeout=20000)
            await funcionalidad.fill("")
            await asyncio.sleep(0.2)
            await funcionalidad.fill("7.8")
            await asyncio.sleep(0.2)
            await funcionalidad.press("Enter")

            await self.esperar_overlay(page)
            await asyncio.sleep(TIEMPO_ESPERA_NAVEGACION / 1000)
            
            self.senales.log.emit(f"✅ [Nav{nav_idx}] Navegación completada")
            return True
            
        except Exception as e:
            self.senales.log.emit(f"❌ [Nav{nav_idx}] Error navegación: {str(e)}")
            return False

    async def ingresar_codigos(self, contenido, tipo, origen, nav_idx):
        """Ingresa los códigos de tipo y origen"""
        try:
            tipo_input = contenido.locator('input[name="tipo_incidencia_codigo"]:not([type="hidden"])')
            await tipo_input.wait_for(state="visible", timeout=10000)
            await tipo_input.fill("")
            await tipo_input.fill(tipo)
            await tipo_input.press("Enter")

            origen_input = contenido.locator('input[name="tipo_origen_incidencia_codigo"]:not([type="hidden"])')
            await origen_input.wait_for(state="visible", timeout=10000)
            await origen_input.fill("")
            await asyncio.sleep(TIEMPO_ESPERA_INGRESO_CODIGOS / 1000)
            await origen_input.fill(origen)
            await asyncio.sleep(TIEMPO_ESPERA_INGRESO_CODIGOS / 1000)
            await origen_input.press("Enter")
            await asyncio.sleep(TIEMPO_ESPERA_CLICK / 1000)

            return True
        except Exception as e:
            self.senales.log.emit(f"⚠️ [Nav{nav_idx}] Error códigos: {str(e)}")
            return False

    async def manejar_boton_volver(self, solapas, guia, nav_idx):
        """Maneja el botón Volver"""
        try:
            self.senales.log.emit(f"⏎ [Nav{nav_idx}] Clic en Volver...")
            await asyncio.sleep(2)
            
            boton_volver = solapas.get_by_role("button", name="Volver")
            
            if await boton_volver.count() == 0:
                self.senales.log.emit(f"⚠️ [Nav{nav_idx}] Botón Volver no encontrado")
                return False
            
            await boton_volver.click(timeout=TIEMPO_ESPERA_VOLVER)
            await self.esperar_overlay(self.pages[nav_idx - 1])
            await asyncio.sleep(3)
            
            return await self.verificar_pagina_activa(self.pages[nav_idx - 1])
            
        except Exception as e:
            self.senales.log.emit(f"⚠️ [Nav{nav_idx}] Error Volver: {str(e)}")
            return False

    async def verificar_incidencia_creada(self, page, nav_idx, guia):
        """Verifica si la incidencia se creó correctamente"""
        try:
            await asyncio.sleep(3)
            
            mensajes_exito = [
                "Incidencia creada", "Éxito", "Success", 
                "Creado correctamente", "Operación exitosa"
            ]
            
            for mensaje in mensajes_exito:
                if await page.get_by_text(mensaje, exact=False).count() > 0:
                    self.senales.log.emit(f"✅ [Nav{nav_idx}] Incidencia creada exitosamente")
                    return True
            
            mensajes_error = [
                "Error", "Falló", "No se pudo crear", 
                "Exception", "No fue posible", "Reintente"
            ]
            
            for mensaje in mensajes_error:
                if await page.get_by_text(mensaje, exact=False).count() > 0:
                    self.senales.log.emit(f"❌ [Nav{nav_idx}] Error detectado: {mensaje}")
                    return False
            
            return None
            
        except Exception as e:
            self.senales.log.emit(f"⚠️ [Nav{nav_idx}] Error en verificación: {str(e)}")
            return None

    async def detectar_error_guia(self, page):
        """Detecta si hay error en la guía"""
        errores = ["No se encontraron", "Error", "No existe", "sin resultados"]
        for texto in errores:
            try:
                if await page.get_by_text(texto, exact=False).count() > 0:
                    return True
            except:
                pass
        return False

    async def _registrar_error(self, guia, error_msg, nav_idx):
        """Registra un error en la lista"""
        async with self.lock:
            self.guias_error.append((guia, f"[Nav{nav_idx}] {error_msg}"))
            self.guias_en_error.add(guia)
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.senales.guia_procesada.emit(
            guia, "❌ ERROR", error_msg, f"Nav{nav_idx}", fecha
        )

    async def _manejar_ent(self, guia, nav_idx, solapas):
        """Maneja el caso de guía ENT"""
        mensaje = f"📦 [Nav{nav_idx}] {guia} - GUÍA ENTREGADA (ENT)"
        self.senales.log.emit(mensaje)
        async with self.lock:
            self.guias_ent.append(guia)
            self.guias_procesadas_ent.add(guia)
        
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.senales.guia_procesada.emit(guia, "📦 ENTREGADA", "ENT", f"Nav{nav_idx}", fecha)
        
        try:
            boton_volver = solapas.get_by_role("button", name="Volver")
            if await boton_volver.count() > 0:
                await boton_volver.click(timeout=10000)
                await self.esperar_overlay(self.pages[nav_idx - 1])
                await asyncio.sleep(2)
        except:
            pass
        
        return True

    async def _evaluar_resultado(self, guia, nav_idx, incidencia_creada, exito_volver, intento):
        """Evalúa el resultado de la creación"""
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if incidencia_creada is True and exito_volver:
            async with self.lock:
                self.guias_procesadas_exito.add(guia)
            self.senales.guia_procesada.emit(
                guia, "✅ PROCESADA", "COMPLETADO", f"Nav{nav_idx}", fecha
            )
            self.senales.log.emit(f"✅ [Nav{nav_idx}] {guia} OK")
            return True
        elif incidencia_creada is None:
            async with self.lock:
                self.guias_advertencia.append((guia, f"[Nav{nav_idx}] Estado indeterminado"))
            self.senales.guia_procesada.emit(
                guia, "⚠️ ADVERTENCIA", "NO CONFIRMADO", f"Nav{nav_idx}", fecha
            )
            return True
        else:
            error_msg = "Error en procesamiento"
            if not incidencia_creada:
                error_msg = "Incidencia no creada"
            if not exito_volver:
                error_msg = "Error al volver"
            
            await self._registrar_error(guia, error_msg, nav_idx)
            
            if intento < MAX_REINTENTOS and not incidencia_creada:
                return False
            return True

    async def _ejecutar_creacion(self, page, guia, nav_idx, contenido):
        """Ejecuta la creación de la incidencia"""
        try:
            async with page.expect_popup(timeout=10000) as pop_info:
                await contenido.get_by_role("button", name="Crear").click()
            popup = await pop_info.value
            await popup.close()
            await asyncio.sleep(2)
            self.senales.log.emit(f"✅ [Nav{nav_idx}] Popup cerrado correctamente")
            return True
        except Exception as e:
            self.senales.log.emit(f"⚠️ [Nav{nav_idx}] Timeout/Error en creación - {str(e)}")
            return await self.verificar_incidencia_creada(page, nav_idx, guia)

    async def _procesar_creacion_incidencia(self, page, guia, nav_idx, resultado, contenido, solapas, intento):
        """Procesa la creación de la incidencia"""
        if await self.detectar_error_guia(page):
            error_msg = "Guía sin resultados"
            await self._registrar_error(guia, error_msg, nav_idx)
            raise Exception(error_msg)

        try:
            await resultado.get_by_role("link", name=guia).click(timeout=10000)
        except Exception as e:
            error_msg = f"No se pudo abrir la guía: {str(e)}"
            await self._registrar_error(guia, error_msg, nav_idx)
            raise Exception(error_msg)

        await self.esperar_overlay(page)
        await asyncio.sleep(TIEMPO_ESPERA_CLICK / 1000)

        if not await self.ingresar_codigos(contenido, self.tipo, "018", nav_idx):
            error_msg = "Error ingresando códigos"
            await self._registrar_error(guia, error_msg, nav_idx)
            raise Exception(error_msg)

        await contenido.locator('textarea[name="ampliacion_incidencia"]').fill(self.ampliacion)

        resultado_creacion = await self._ejecutar_creacion(page, guia, nav_idx, contenido)
        exito_volver = await self.manejar_boton_volver(solapas, guia, nav_idx)
        
        return await self._evaluar_resultado(
            guia, nav_idx, resultado_creacion, exito_volver, intento
        )

    async def crear_incidencia(self, page, guia, nav_idx, intento=1):
        """Crea una incidencia para una guía"""
        async with self.lock:
            if any(guia in s for s in [self.guias_procesadas_exito, 
                                        self.guias_procesadas_ent, 
                                        self.guias_en_error]):
                self.senales.log.emit(f"⏭️ [Nav{nav_idx}] Guía {guia} ya procesada - omitiendo")
                return True
        
        if not await self.verificar_pagina_activa(page):
            raise Exception("Página no activa")
        
        principal = (
            page
            .frame_locator('frame[name="menu"]')
            .frame_locator('iframe[name="principal"]')
        )

        filtro = principal.frame_locator('frame[name="filtro"]')
        resultado = principal.frame_locator('frame[name="resultado"]')
        contenido = principal.frame_locator('frame[name="contenido"]')
        solapas = principal.frame_locator('frame[name="solapas"]')

        try:
            envio = filtro.locator('input[name="nenvio"]:not([type="hidden"])')
            await envio.wait_for(state="visible", timeout=15000)
            await envio.fill("")
            await asyncio.sleep(0.5)
            await envio.fill(guia)
            await asyncio.sleep(0.5)
            await envio.press("Enter")
        except Exception as e:
            error_msg = f"Campo búsqueda no disponible: {str(e)}"
            await self._registrar_error(guia, error_msg, nav_idx)
            raise Exception(error_msg)

        await self.esperar_overlay(page)
        await asyncio.sleep(TIEMPO_ESPERA_CLICK / 1000)

        if await self.verificar_estado_ent(page, nav_idx):
            return await self._manejar_ent(guia, nav_idx, solapas)

        return await self._procesar_creacion_incidencia(
            page, guia, nav_idx, resultado, contenido, solapas, intento
        )

    async def trabajador_navegador(self, nav_idx, total_guias, resultados):
        """Worker para cada navegador"""
        try:
            page = self.pages[nav_idx - 1]
            guias_procesadas_local = 0
            
            while self.procesando and not self.cancelado:
                async with self.lock:
                    if not self.cola_guias:
                        break
                    guia = self.cola_guias.pop(0)
                    
                    if any(guia in s for s in [self.guias_procesadas_exito,
                                                self.guias_procesadas_ent,
                                                self.guias_en_error]):
                        self.senales.log.emit(f"⏭️ [Nav{nav_idx}] Guía {guia} ya procesada - saltando")
                        continue
                
                try:
                    self.senales.log.emit(f"🌐 [Nav{nav_idx}] Procesando: {guia}")
                    exito = await self.crear_incidencia(page, guia, nav_idx)
                    
                    if exito:
                        guias_procesadas_local += 1
                        async with self.lock:
                            resultados['exitosas'] += 1
                    
                except Exception as e:
                    self.senales.log.emit(f"❌ [Nav{nav_idx}] Error: {str(e)}")
                
                async with self.lock:
                    resultados['progreso'] += 1
                    progreso = int(resultados['progreso'] / total_guias * 100)
                    self.senales.progreso.emit(progreso)
                    await self.calcular_tiempo_restante(resultados['progreso'], total_guias)
                    self.senales.estado.emit(
                        f"Progreso: {resultados['progreso']}/{total_guias} ({progreso}%) "
                        f"- Éxitos: {resultados['exitosas']}"
                    )
            
            self.senales.log.emit(f"📊 [Nav{nav_idx}] Procesó {guias_procesadas_local} guías")
            
        except Exception as e:
            self.senales.log.emit(f"❌ [Nav{nav_idx}] Error fatal: {str(e)}")

    async def _inicializar_navegadores(self, p):
        """Inicializa los navegadores"""
        for i in range(self.num_navegadores):
            if self.cancelado:
                break
                
            self.senales.log.emit(f"▶️ Iniciando navegador {i+1}/{self.num_navegadores}...")
            
            browser = await p.chromium.launch(
                headless=False,
                args=['--start-maximized', '--disable-dev-shm-usage']
            )
            
            context = await browser.new_context(
                viewport={'width': 1280, 'height': 800},
                locale="es-ES"
            )
            
            page = await context.new_page()
            page.set_default_timeout(60000)
            
            self.browsers.append(browser)
            self.contexts.append(context)
            self.pages.append(page)
            
            await page.goto(URL_ALERTRAN, timeout=60000)
            await asyncio.sleep(3)
            
            if not await self.hacer_login(page, i+1):
                self.senales.error.emit(f"Error login navegador {i+1}")
                return False
            
            if not await self.navegar_a_funcionalidad_7_8(page, i+1):
                self.senales.error.emit(f"Error navegación navegador {i+1}")
                return False
        
        return True

    def _finalizar_proceso(self, exitosas):
        """Finaliza el proceso y guarda resultados"""
        if not self.cancelado:
            if self.guias_error or self.guias_advertencia:
                ruta = self.file_utils.guardar_errores_excel(
                    self.guias_error, self.guias_advertencia, self.carpeta_descargas
                )
                self.senales.archivo_errores.emit(ruta)
            
            tiempo_total = time.time() - self.tiempo_inicio
            tiempo_formateado = str(timedelta(seconds=int(tiempo_total)))
            
            self.senales.log.emit(f"\n 🕑 Completado en {tiempo_formateado}")
            self.senales.log.emit(f" 📝 Desviaciones creadas: {exitosas - len(self.guias_ent)}")
            self.senales.log.emit(f" 📦 Guías ENT (omitidas): {len(self.guias_ent)}")
            self.senales.log.emit(f" ❌ Errores: {len(self.guias_error)}")
            self.senales.log.emit(f" ⚠️ Advertencias: {len(self.guias_advertencia)}")
            
            self.senales.finalizado.emit()
        else:
            self.senales.proceso_cancelado.emit()

    async def proceso_principal(self):
        """Método principal con múltiples navegadores"""
        try:
            guias = self.leer_excel(self.excel_path)
            self.total_guias = len(guias)
            
            if self.total_guias == 0:
                self.senales.error.emit("El archivo Excel no contiene guías")
                return

            self.tiempo_inicio = time.time()
            self.senales.estado.emit(f"Procesando {self.total_guias} guías con {self.num_navegadores} navegador(es)...")

            async with async_playwright() as p:
                if not await self._inicializar_navegadores(p):
                    return

                if not self.cancelado:
                    self.cola_guias = guias.copy()
                    resultados = {'progreso': 0, 'exitosas': 0}

                    tareas = []
                    for i in range(self.num_navegadores):
                        tarea = self.trabajador_navegador(i+1, self.total_guias, resultados)
                        tareas.append(tarea)

                    await asyncio.gather(*tareas)

                for browser in self.browsers:
                    await browser.close()

                self._finalizar_proceso(resultados['exitosas'])

        except Exception as e:
            self.senales.error.emit(f"Error: {str(e)}")

    def cancelar(self):
        """Cancela el proceso"""
        self.cancelado = True
        self.procesando = False

    def run(self):
        """Ejecuta el thread"""
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(self.proceso_principal())
        finally:
            loop.close()