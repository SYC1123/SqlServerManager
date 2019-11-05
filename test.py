import sys
import first
from PyQt5.QtWidgets import QApplication, QMainWindow

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = QMainWindow()
    ui = first.Ui_Dialog()
    ui.setupUi(mainWindow)
    mainWindow.show()
    sys.exit(app.exec())
