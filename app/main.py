import sys
import os
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QFileDialog, QListWidget, QListWidgetItem,
                             QMessageBox)
from PyQt6.QtGui import QPixmap, QIcon
from PyQt6.QtCore import QSize, Qt
from app.core.processing import cargar_zip, procesar_zip, reaplicar_recomendaciones
from app.report.pptx_writer import export_groups_to_pptx_report
from app.report.xlsx_writer import export_groups_to_xlsx_report

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("InspectW Desktop")
        self.resize(900, 600)

        # --- Botones ---
        self.btnZip = QPushButton("Cargar ZIP(s)")
        self.btnClear = QPushButton("Limpiar")
        self.btnPptReport = QPushButton("Generar Informe A4 (PPTX)")
        self.btnXlsxReport = QPushButton("Generar Informe (XLSX)")
        self.btnHist = QPushButton("Cargar historico.csv (opcional)")

        # --- Listas ---
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
        self.btnClear.clicked.connect(self.on_limpiar)
        self.btnHist.clicked.connect(self.on_cargar_hist)
        self.btnPptReport.clicked.connect(self.on_generar_reporte_pptx)
        self.btnXlsxReport.clicked.connect(self.on_generar_reporte_xlsx)
        self.lista.currentItemChanged.connect(self.on_grupo_seleccionado)

        # --- Estado inicial ---
        self.on_limpiar() # Usamos on_limpiar para establecer el estado inicial

    def on_limpiar(self):
        """Limpia los datos cargados y la interfaz."""
        self.grupos = {}
        self.archivos = {}
        self.hist_path = None
        self.lista.clear()
        self.listaFotos.clear()

    def on_cargar_zip(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Selecciona uno o más archivos ZIP", "", "ZIP (*.zip)")
        if not paths:
            return

        error_mostrado = False
        for path in paths:
            nuevos_archivos = cargar_zip(path)
            self.archivos.update(nuevos_archivos)
            
            nuevos_grupos, error = procesar_zip(nuevos_archivos, hist_path=self.hist_path)
            
            if error and not error_mostrado:
                QMessageBox.warning(self, "Error al Cargar Histórico",
                    f"No se pudieron aplicar las recomendaciones del archivo histórico. "
                    f"El procesamiento del ZIP continuó, pero sin sugerencias.\n\n"
                    f"Error: {error}\n\n"
                    f"Asegúrate de que el archivo .csv tenga el formato correcto (separado por ';', codificación 'latin1' y columnas 'TAG' y 'RECOMENDACIÓN').")
                error_mostrado = True # Mostrar el error solo una vez por tanda

            for key, grupo_nuevo in nuevos_grupos.items():
                if key in self.grupos:
                    self.grupos[key].fotos.extend(grupo_nuevo.fotos)
                else:
                    self.grupos[key] = grupo_nuevo
        
        self.actualizar_lista_grupos()

    def actualizar_lista_grupos(self):
        """Refresca la lista de grupos en la UI con los datos actuales."""
        self.lista.clear()
        self.listaFotos.clear()
        for k, g in sorted(self.grupos.items()):
            item = QListWidgetItem(f"{k} ({len(g.fotos)} fotos)")
            item.setData(Qt.ItemDataRole.UserRole, k)
            self.lista.addItem(item)

    def on_grupo_seleccionado(self, current, previous):
        self.listaFotos.clear()
        if not current:
            return

        key = current.data(Qt.ItemDataRole.UserRole)
        if key not in self.grupos:
            return

        grupo = self.grupos[key]
        
        for foto in grupo.fotos:
            path_in_zip = f"{foto.carpeta}/{foto.filename}"
            img_data = self.archivos.get(path_in_zip)
            
            if img_data is None:
                path_in_zip_alt = path_in_zip.replace('/', '\\')
                img_data = self.archivos.get(path_in_zip_alt)

            if img_data:
                pixmap = QPixmap()
                pixmap.loadFromData(img_data)
                if not pixmap.isNull():
                    item = QListWidgetItem(foto.filename)
                    item.setIcon(QIcon(pixmap))
                    self.listaFotos.addItem(item)

    def on_generar_reporte_pptx(self):
        if not self.grupos:
            QMessageBox.warning(self, "Aviso", "Carga primero uno o más ZIPs.")
            return
        destino, _ = QFileDialog.getSaveFileName(self, "Guardar PPTX", "", "PowerPoint (*.pptx)")
        if not destino: return
        export_groups_to_pptx_report(self.grupos, self.archivos, destino)
        QMessageBox.information(self, "Listo", f"Guardado en:\n{destino}")

    def on_generar_reporte_xlsx(self):
        if not self.grupos:
            QMessageBox.warning(self, "Aviso", "Carga primero uno o más ZIPs.")
            return
            
        destino, _ = QFileDialog.getSaveFileName(self, "Guardar XLSX", "", "Excel (*.xlsx)")
        if not destino: 
            return
            
        try:
            with open(destino, 'a'): pass
        except PermissionError:
            QMessageBox.critical(self, "Error", 
                "No se puede guardar el archivo porque está abierto en otro programa.\n"
                "Cierra Excel u otro programa que pueda estar usando el archivo e intenta nuevamente.")
            return
        except Exception:
            pass
            
        try:
            export_groups_to_xlsx_report(self.grupos, self.archivos, destino)
            QMessageBox.information(self, "Listo", f"Guardado en:\n{destino}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error al guardar el archivo:\n{str(e)}")
        
    def on_cargar_hist(self):
        path, _ = QFileDialog.getOpenFileName(self, "Selecciona historico.csv", "", "CSV (*.csv)")
        if not path:
            return

        self.hist_path = path
        QMessageBox.information(self, "Histórico Cargado", 
            f"Se usará el archivo:\n{path}\n\n"
            "Los próximos ZIPs que cargues usarán este histórico para las recomendaciones.")

        if self.grupos:
            reply = QMessageBox.question(self, 'Aplicar Histórico',
                "Ya hay datos ZIP cargados. ¿Deseas aplicar las recomendaciones de este histórico a los datos existentes?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes)
            
            if reply == QMessageBox.StandardButton.Yes:
                error = reaplicar_recomendaciones(self.grupos, self.hist_path)
                if error:
                    QMessageBox.critical(self, "Error al Aplicar Histórico",
                        f"No se pudieron aplicar las recomendaciones.\n\n"
                        f"Error: {error}\n\n"
                        f"Asegúrate de que el archivo .csv tenga el formato correcto (separado por ';', codificación 'latin1' y columnas 'TAG' y 'RECOMENDACIÓN').")
                else:
                    QMessageBox.information(self, "Éxito", "Se han actualizado las recomendaciones para todos los grupos cargados.")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
