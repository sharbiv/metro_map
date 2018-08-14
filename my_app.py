import sys
from PyQt5.QtWidgets import QApplication, QWidget

if __name__ == '__main__':
    app = QApplication(sys.argv)

    w = QWidget()
    w.setGeometry(300, 300, 650, 550)
    w.setWindowTitle('Metro Maps')
    w.setGeometry(300, 350, 650, 550)
    w.show()

    sys.exit(app.exec_())
