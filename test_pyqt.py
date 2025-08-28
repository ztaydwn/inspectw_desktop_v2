import sys
from PyQt6.QtWidgets import QApplication, QWidget

if __name__ == '__main__':
    print("Creating QApplication...")
    app = QApplication(sys.argv)
    print("Creating QWidget...")
    w = QWidget()
    w.resize(250, 150)
    w.move(300, 300)
    w.setWindowTitle('Simple')
    print("Showing QWidget...")
    w.show()
    print("Executing app...")
    sys.exit(app.exec())
