import sys
from PyQt5 import QtWidgets, QtCore, QtGui
import tkinter as tk
from PIL import ImageGrab
import numpy as np
import cv2
import os
import re


class MyWidget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        root = tk.Tk()
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        self.setGeometry(0, 0, screen_width, screen_height)
        self.setWindowTitle(' ')
        self.begin = QtCore.QPoint()
        self.end = QtCore.QPoint()
        self.setWindowOpacity(0.3)
        self.num_snip = 0
        self.is_snipping = False
        QtWidgets.QApplication.setOverrideCursor(
            QtGui.QCursor(QtCore.Qt.CrossCursor)
        )
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        print('Capture the screen...')
        print('Press Esc if you want to quit...')
        self.show()

    def paintEvent(self, event):
        if self.is_snipping:
            brush_color = (0, 0, 0, 0)
            lw = 0
            opacity = 0
        else:
            brush_color = (128, 128, 255, 128)
            lw = 3
            opacity = 0.3

        self.setWindowOpacity(opacity)
        qp = QtGui.QPainter(self)
        qp.setPen(QtGui.QPen(QtGui.QColor('black'), lw))
        qp.setBrush(QtGui.QColor(*brush_color))
        qp.drawRect(QtCore.QRect(self.begin, self.end))

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Escape:
            print('Quit')
            self.close()
        event.accept()

    def mousePressEvent(self, event):
        self.begin = event.pos()
        self.end = self.begin
        self.update()

    def mouseMoveEvent(self, event):
        self.end = event.pos()
        self.update()

    def mouseReleaseEvent(self, event):
        # self.close()
        self.num_snip += 1
        x1 = min(self.begin.x(), self.end.x())
        y1 = min(self.begin.y(), self.end.y())
        x2 = max(self.begin.x(), self.end.x())
        y2 = max(self.begin.y(), self.end.y())

        self.is_snipping = True
        self.repaint()
        QtWidgets.QApplication.processEvents()
        img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        self.is_snipping = False
        self.repaint()
        QtWidgets.QApplication.processEvents()
        if not os.path.isdir(os.path.join(os.path.dirname(os.path.realpath(__file__)), "map")):
            os.makedirs(os.path.join(os.path.dirname(os.path.realpath(__file__)), "map"))
        else:
            file_list = os.listdir(os.path.join(os.path.dirname(os.path.realpath(__file__)), "map"))
            if file_list:
                regex = re.compile(r'\d+')
                num_list = []
                for filename in file_list:
                    num_list.append(int(regex.findall(filename)[0]))
                self.num_snip = max(num_list) + 1
        img_name = 'snip_{}.png'.format(self.num_snip)
        img.save(os.path.join(os.path.dirname(os.path.realpath(__file__)), "map\{}".format(img_name)))
        print(img_name, 'saved')
        img = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2RGB)

        # cv2.imshow('Captured Image', img)
        # cv2.waitKey(0)
        # cv2.destroyAllWindows()


if __name__ == '__main__':
# def snip():
    app = QtWidgets.QApplication(sys.argv)
    window = MyWidget()
    # window.show()
    app.aboutToQuit.connect(app.deleteLater)
    sys.exit(app.exec_())
