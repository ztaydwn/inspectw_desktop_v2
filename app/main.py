from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QFileDialog, QListWidget, QListWidgetItem,
                             QMessageBox)
from PyQt6.QtGui import QPixmap, QIcon
from PyQt6.QtCore import QSize, Qt
from app.core.processing import cargar_zip, procesar_zip
from app.report.pptx_writer import export_groups_to_pptx_report
from app.report.xlsx_writer import export_groups_to_xlsx_report
import sys, os

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("InspectW Desktop")
        self.resize(900, 600)

        self.btnZip = QPushButton("Cargar ZIP")
        self.btnPptReport = QPushButton("Generar Informe A4 (PPTX)")
        self.btnXlsxReport = QPushButton("Generar Informe (XLSX)")
        self.lista = QListWidget()
        self.listaFotos = QListWidget()
        self.listaFotos.setViewMode(QListWidget.ViewMode.IconMode)
        self.listaFotos.setIconSize(QSize(128, 128))
        self.listaFotos.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.listaFotos.setWordWrap(True)
        self.btnHist = QPushButton("Cargar historico.csv (opcional)")
        self.btnHist.clicked.connect(self.on_cargar_hist)
        self.hist_path = None

        main_layout = QVBoxLayout(self)
        
        h_layout = QHBoxLayout()
        h_layout.addWidget(self.lista, 2)
        h_layout.addWidget(self.listaFotos, 2)

        # Crear layout para los botones de exportación
        export_layout = QHBoxLayout()
        export_layout.addWidget(self.btnPptReport)
        export_layout.addWidget(self.btnXlsxReport)
        
        

        main_layout.addWidget(self.btnZip)
        main_layout.addWidget(self.btnHist)
        main_layout.addLayout(h_layout)
        main_layout.addLayout(export_layout)

        self.btnZip.clicked.connect(self.on_cargar_zip)
        self.btnPptReport.clicked.connect(self.on_generar_reporte_pptx)
        self.btnXlsxReport.clicked.connect(self.on_generar_reporte_xlsx)
        self.lista.currentItemChanged.connect(self.on_grupo_seleccionado)

        self.grupos = {}
        self.archivos = {}

    def on_cargar_zip(self):
        path, _ = QFileDialog.getOpenFileName(self, "Selecciona ZIP", "", "ZIP (*.zip)")
        if not path: return
        self.archivos = cargar_zip(path)
        # pasa la ruta del histórico si el usuario la cargó; si no, usa la ruta por defecto
        self.grupos = procesar_zip(self.archivos, hist_path=self.hist_path)
        self.lista.clear()
        self.listaFotos.clear()
        for k, g in self.grupos.items():
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
            QMessageBox.warning(self, "Aviso", "Carga primero un ZIP.")
            return
        destino, _ = QFileDialog.getSaveFileName(self, "Guardar PPTX", "", "PowerPoint (*.pptx)")
        if not destino: return
        export_groups_to_pptx_report(self.grupos, self.archivos, destino)
        QMessageBox.information(self, "Listo", f"Guardado en:\n{destino}")

    def on_generar_reporte_xlsx(self):
        if not self.grupos:
            QMessageBox.warning(self, "Aviso", "Carga primero un ZIP.")
            return
            
        destino, _ = QFileDialog.getSaveFileName(self, "Guardar XLSX", "", "Excel (*.xlsx)")
        if not destino: 
            return
            
        try:
            # Verificar si el archivo está en uso
            try:
                with open(destino, 'a'):
                    pass
            except PermissionError:
                QMessageBox.critical(self, "Error", 
                    "No se puede guardar el archivo porque está abierto en otro programa.\n"
                    "Cierra Excel u otro programa que pueda estar usando el archivo e intenta nuevamente.")
                return
            except Exception:
                pass  # Si el archivo no existe, está bien
                
            # Intentar generar el reporte
            export_groups_to_xlsx_report(self.grupos, self.archivos, destino)
            QMessageBox.information(self, "Listo", f"Guardado en:\n{destino}")
            
        except PermissionError:
            QMessageBox.critical(self, "Error", 
                "No tienes permisos para guardar en esta ubicación.\n"
                "Intenta guardar en otra carpeta o ejecuta el programa como administrador.")
        except Exception as e:
            QMessageBox.critical(self, "Error", 
                f"Ocurrió un error al guardar el archivo:\n{str(e)}\n\n"
                "Intenta guardar en otra ubicación o con otro nombre.")
        
    def on_cargar_hist(self):
        path, _ = QFileDialog.getOpenFileName(self, "Selecciona historico.csv", "", "CSV (*.csv)")
        if path:
            self.hist_path = path

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
