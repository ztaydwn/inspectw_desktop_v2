from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QFileDialog, QListWidget, QListWidgetItem,
                             QMessageBox)
from PyQt6.QtGui import QPixmap, QIcon
from PyQt6.QtCore import QSize, Qt
from .core.processing import cargar_zip, procesar_zip
from .report.pptx_writer import export_groups_to_pptx_report
import sys, os

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("InspectW Desktop")
        self.resize(900, 600)

        self.btnZip = QPushButton("Cargar ZIP")
        self.btnPptMosaic = QPushButton("Generar Mosaico (PPTX)")
        self.btnPptReport = QPushButton("Generar Informe A4 (PPTX)")
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
        h_layout.addWidget(self.lista, 1)
        h_layout.addWidget(self.listaFotos, 3)

        # Crear layout para los botones de exportación
        export_layout = QHBoxLayout()
        export_layout.addWidget(self.btnPptMosaic)
        export_layout.addWidget(self.btnPptReport)
        
        self.btnPptMosaic.setEnabled(False)
        self.btnPptMosaic.setToolTip("Función no disponible en esta versión.")

        main_layout.addWidget(self.btnZip)
        main_layout.addWidget(self.btnHist)
        main_layout.addLayout(h_layout)
        main_layout.addLayout(export_layout)

        self.btnZip.clicked.connect(self.on_cargar_zip)
        self.btnPptReport.clicked.connect(self.on_generar_reporte)
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

    def on_generar_reporte(self):
        if not self.grupos:
            QMessageBox.warning(self, "Aviso", "Carga primero un ZIP.")
            return
        destino, _ = QFileDialog.getSaveFileName(self, "Guardar PPTX", "", "PowerPoint (*.pptx)")
        if not destino: return
        export_groups_to_pptx_report(self.grupos, self.archivos, destino)
        QMessageBox.information(self, "Listo", f"Guardado en:\n{destino}")
        
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
