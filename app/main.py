import sys
import os
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QFileDialog, QListWidget, QListWidgetItem,
                             QMessageBox, QProgressDialog)
from PyQt6.QtGui import QPixmap, QIcon
from PyQt6.QtCore import QSize, Qt, QObject, QThread, pyqtSignal
from app.core.processing import (cargar_zip, procesar_zip, reaplicar_recomendaciones, _find_image_data,
                                 cargar_directorio)
from app.report.pptx_writer import export_groups_to_pptx_report
from app.report.xlsx_writer import export_groups_to_xlsx_report

class DataProcessorWorker(QObject):
    finished = pyqtSignal(dict, dict, list)
    progress = pyqtSignal(str)

    def __init__(self, paths, hist_path, mode='zip'):
        super().__init__()
        self.paths = paths
        self.hist_path = hist_path
        self.mode = mode
        self._is_running = True

    def run(self):
        grupos_acumulados, archivos_acumulados, errors = {}, {}, []
        
        loader_func = cargar_zip if self.mode == 'zip' else cargar_directorio

        for i, path in enumerate(self.paths):
            if not self._is_running: break
            try:
                base_name = os.path.basename(path)
                self.progress.emit(f"Procesando {i+1}/{len(self.paths)}: {base_name}...")
                
                # Usar la función de carga correspondiente
                nuevos_archivos = loader_func(path)

                archivos_acumulados.update(nuevos_archivos)
                nuevos_grupos, error = procesar_zip(nuevos_archivos, hist_path=self.hist_path)
                if error: errors.append(f"Error en {os.path.basename(path)}: {error}")
                for key, grupo_nuevo in nuevos_grupos.items():
                    if key in grupos_acumulados: grupos_acumulados[key].fotos.extend(grupo_nuevo.fotos)
                    else: grupos_acumulados[key] = grupo_nuevo
            except Exception as e:
                errors.append(f"Error crítico procesando {base_name}: {e}")
        self.finished.emit(grupos_acumulados, archivos_acumulados, errors)

    def stop(self): self._is_running = False

class ReportWorker(QObject):
    finished = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, report_type, grupos, archivos, destino):
        super().__init__()
        self.report_type = report_type
        self.grupos = grupos
        self.archivos = archivos
        self.destino = destino
        self._is_running = True

    def run(self):
        try:
            if not self._is_running: 
                self.finished.emit("Generación de informe cancelada.")
                return

            if self.report_type == 'xlsx':
                export_groups_to_xlsx_report(self.grupos, self.archivos, self.destino, progress_callback=self.progress)
            elif self.report_type == 'pptx':
                export_groups_to_pptx_report(self.grupos, self.archivos, self.destino, progress_callback=self.progress)
            
            if self._is_running:
                self.finished.emit(f"¡Informe guardado con éxito en:\n{self.destino}")
        except Exception as e:
            self.finished.emit(f"Ocurrió un error al generar el informe:\n{e}")

    def stop(self): self._is_running = False

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("InspectW Desktop")
        self.resize(900, 600)
        self.thread = None
        self.worker = None

        # --- Widgets ---
        self.btnZip = QPushButton("Cargar ZIP(s)")
        self.btnDir = QPushButton("Cargar Carpeta")
        self.btnClear = QPushButton("Limpiar")
        self.btnPptReport = QPushButton("Generar Informe A4 (PPTX)")
        self.btnXlsxReport = QPushButton("Generar Informe (XLSX)")
        self.btnHist = QPushButton("Cargar historico.csv (opcional)")
        self.lista = QListWidget()
        self.listaFotos = QListWidget()
        self.listaFotos.setViewMode(QListWidget.ViewMode.IconMode)
        self.listaFotos.setIconSize(QSize(128, 128))
        self.listaFotos.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.listaFotos.setWordWrap(True)

        # --- Layouts ---
        main_layout = QVBoxLayout(self)
        top_buttons_layout = QHBoxLayout()
        top_buttons_layout.addWidget(self.btnZip)
        top_buttons_layout.addWidget(self.btnDir)
        top_buttons_layout.addWidget(self.btnClear)
        top_buttons_layout.addWidget(self.btnHist)
        h_layout = QHBoxLayout()
        h_layout.addWidget(self.lista, 2)
        h_layout.addWidget(self.listaFotos, 2)
        export_layout = QHBoxLayout()
        export_layout.addWidget(self.btnPptReport)
        export_layout.addWidget(self.btnXlsxReport)
        main_layout.addLayout(top_buttons_layout)
        main_layout.addLayout(h_layout)
        main_layout.addLayout(export_layout)

        # --- Conexiones ---
        self.btnZip.clicked.connect(self.on_cargar_zip)
        self.btnDir.clicked.connect(self.on_cargar_carpeta)
        self.btnClear.clicked.connect(self.on_limpiar)
        self.btnHist.clicked.connect(self.on_cargar_hist)
        self.btnPptReport.clicked.connect(lambda: self.generar_informe('pptx'))
        self.btnXlsxReport.clicked.connect(lambda: self.generar_informe('xlsx'))
        self.lista.currentItemChanged.connect(self.on_grupo_seleccionado)

        self.on_limpiar()

    def set_ui_busy(self, busy, message=""):
        self.btnZip.setEnabled(not busy)
        self.btnDir.setEnabled(not busy)
        self.btnClear.setEnabled(not busy)
        self.btnHist.setEnabled(not busy)
        self.btnPptReport.setEnabled(not busy)
        self.btnXlsxReport.setEnabled(not busy)
        self.setWindowTitle(f"InspectW Desktop {message}".strip())

    def on_limpiar(self):
        if self.thread and self.thread.isRunning():
            QMessageBox.warning(self, "Aviso", "No se puede limpiar mientras se procesan archivos.")
            return
        self.grupos, self.archivos, self.hist_path = {}, {}, None
        self.lista.clear()
        self.listaFotos.clear()

    def on_cargar_zip(self):
        if self.thread and self.thread.isRunning(): return
        paths, _ = QFileDialog.getOpenFileNames(self, "Selecciona uno o más archivos ZIP", "", "ZIP (*.zip)")
        if not paths: return

        self.iniciar_procesamiento(paths, mode='zip')

    def on_cargar_carpeta(self):
        if self.thread and self.thread.isRunning(): return
        path = QFileDialog.getExistingDirectory(self, "Selecciona la carpeta del proyecto")
        if not path: return

        # Verificar que los archivos necesarios existen
        if not os.path.exists(os.path.join(path, "descriptions.txt")) or not os.path.exists(os.path.join(path, "grupos.txt")):
            QMessageBox.warning(self, "Archivos Faltantes", "La carpeta seleccionada debe contener 'descriptions.txt' y 'grupos.txt'.")
            return

        self.iniciar_procesamiento([path], mode='dir')

    def iniciar_procesamiento(self, paths, mode):
        self.set_ui_busy(True, "(Iniciando...)")
        self.thread = QThread()
        self.worker = DataProcessorWorker(paths, self.hist_path, mode=mode)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_processing_finished)
        self.worker.progress.connect(lambda msg: self.set_ui_busy(True, f"({msg})"))
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self.clear_thread_references) # Limpiar referencia
        self.thread.start()

    def on_processing_finished(self, nuevos_grupos, nuevos_archivos, errors):
        self.archivos.update(nuevos_archivos)
        for key, grupo_nuevo in nuevos_grupos.items():
            if key in self.grupos: self.grupos[key].fotos.extend(grupo_nuevo.fotos)
            else: self.grupos[key] = grupo_nuevo
        self.actualizar_lista_grupos()
        self.set_ui_busy(False)
        if errors:
            QMessageBox.warning(self, "Errores Durante el Procesamiento", f"Se encontraron problemas:\n\n- {'\n- '.join(errors)}")

    def generar_informe(self, report_type):
        if not self.grupos: 
            QMessageBox.warning(self, "Aviso", "Carga primero uno o más ZIPs.")
            return
        if self.thread and self.thread.isRunning():
            QMessageBox.warning(self, "Aviso", "Espera a que termine el proceso actual.")
            return

        ext, file_filter = ('pptx', "PowerPoint (*.pptx)") if report_type == 'pptx' else ('xlsx', "Excel (*.xlsx)")
        destino, _ = QFileDialog.getSaveFileName(self, f"Guardar {ext.upper()}", "", file_filter)
        if not destino: return

        progress = QProgressDialog(f"Generando informe {ext.upper()}...", "Cancelar", 0, 100, self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        
        self.thread = QThread()
        self.worker = ReportWorker(report_type, self.grupos, self.archivos, destino)
        self.worker.moveToThread(self.thread)
        self.worker.progress.connect(progress.setValue)
        progress.canceled.connect(self.worker.stop)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(progress.close)
        self.worker.finished.connect(lambda msg: QMessageBox.information(self, "Proceso Terminado", msg))
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self.clear_thread_references) # Limpiar referencia
        
        self.thread.start()
        progress.exec()

    def clear_thread_references(self):
        """Slot para limpiar las referencias al worker y al thread cuando terminan."""
        self.worker = None
        self.thread = None

    def actualizar_lista_grupos(self):
        self.lista.clear()
        self.listaFotos.clear()
        for k, g in sorted(self.grupos.items()):
            item = QListWidgetItem(f"{k} ({len(g.fotos)} fotos)")
            item.setData(Qt.ItemDataRole.UserRole, k)
            self.lista.addItem(item)

    def on_grupo_seleccionado(self, current, previous):
        self.listaFotos.clear()
        if not current: return
        key = current.data(Qt.ItemDataRole.UserRole)
        if key not in self.grupos: return
        grupo = self.grupos[key]
        for foto in grupo.fotos:
            img_data = _find_image_data(self.archivos, foto)
            if img_data:
                pixmap = QPixmap()
                pixmap.loadFromData(img_data)
                if not pixmap.isNull():
                    item = QListWidgetItem(foto.filename)
                    item.setIcon(QIcon(pixmap))
                    self.listaFotos.addItem(item)

    def on_cargar_hist(self):
        if self.thread and self.thread.isRunning():
            QMessageBox.warning(self, "Aviso", "Espera a que termine el proceso actual.")
            return
        path, _ = QFileDialog.getOpenFileName(self, "Selecciona historico.csv", "", "CSV (*.csv)")
        if not path: return
        self.hist_path = path
        QMessageBox.information(self, "Histórico Cargado", f"Se usará el archivo:\n{path}")
        if self.grupos:
            reply = QMessageBox.question(self, 'Aplicar Histórico', "¿Deseas aplicar las recomendaciones a los datos ya cargados?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes)
            if reply == QMessageBox.StandardButton.Yes:
                error = reaplicar_recomendaciones(self.grupos, self.hist_path)
                if error: QMessageBox.critical(self, "Error al Aplicar Histórico", f"No se pudieron aplicar las recomendaciones.\n\nError: {error}")
                else: QMessageBox.information(self, "Éxito", "Se han actualizado las recomendaciones.")

    def closeEvent(self, event):
        if self.thread and self.thread.isRunning():
            self.worker.stop()
            self.thread.quit()
            self.thread.wait(500)
        event.accept()

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()