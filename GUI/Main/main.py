# coding=utf-8
"""
Author: vision
date: 2019/4/4 19:46
"""

from PyQt5 import QtWidgets
import sys
from GUI.ui import Ui

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QWidget()
    ui = Ui()
    ui.setupUi(MainWindow)
    ui.function()
    MainWindow.show()
    sys.exit(app.exec_())
