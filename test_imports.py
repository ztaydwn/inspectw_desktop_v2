print("Importing PyQt6...")
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QFileDialog, QListWidget, QListWidgetItem,
                             QMessageBox)
from PyQt6.QtGui import QPixmap, QIcon
from PyQt6.QtCore import QSize, Qt
print("PyQt6 imported successfully.")

print("Importing app.core.processing...")
from app.core.processing import cargar_zip, procesar_zip
print("app.core.processing imported successfully.")

print("Importing app.report.pptx_writer...")
from app.report.pptx_writer import export_groups_to_pptx_report
print("app.report.pptx_writer imported successfully.")
