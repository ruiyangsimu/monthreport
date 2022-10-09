import logging
import os
import sys

import yaml
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QDialog, QMainWindow, QMessageBox
from PyQt5.uic import loadUi

from gen_data import Ui_Gen_Data
from gen_pic import Ui_Gen_Pic
from gen_word import Ui_Gen_Word
from main_window_ui import Ui_MainWindow
from picture import Picture

with open('./config/log.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)

logger = logging.getLogger("app")


class Window(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.picture = Picture()
        self.picture.load()
        self.setupUi(self)
        self.connectSignalsSlots()

    def connectSignalsSlots(self):
        self.action_Exit.triggered.connect(self.close)
        # self.action_Find_Replace.triggered.connect(self.findAndReplace)
        self.action_Config.triggered.connect(self.openConfig)
        self.action_genword.triggered.connect(self.genWord)
        self.action_GenPic.triggered.connect(self.genPic)
        self.action_About.triggered.connect(self.about)
        self.action_Data.triggered.connect(self.genData)

    def openConfig(self):
        current_path = os.path.abspath(__file__)
        dir_name = os.path.dirname(current_path)
        os.system(r"start "+dir_name+"//config//config.yaml")

    def genData(self):
        self.window = QtWidgets.QMainWindow()
        self.gen_data = Ui_Gen_Data()
        self.gen_data.setupUi(self.window)
        self.gen_data.setData(self.picture)
        self.window.show()

    def genWord(self):
        self.window = QtWidgets.QMainWindow()
        self.gen_word = Ui_Gen_Word()
        self.gen_word.setupUi(self.window)
        self.gen_word.setData(self.picture)
        self.window.show()

    def findAndReplace(self):
        dialog = FindReplaceDialog(self)
        dialog.exec()

    def genPic(self):
        self.window = QtWidgets.QMainWindow()
        self.gen_pic = Ui_Gen_Pic()
        self.gen_pic.setupUi(self.window)
        self.gen_pic.setData(self.picture)
        self.window.show()

    def about(self):
        QMessageBox.about(
            self,
            "关于",
            "<p>睿扬运营系统:</p>"
            "<p>- 月报生成功能</p>",
        )


class FindReplaceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        loadUi("ui/find_replace.ui", self)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = Window()
    win.show()
    sys.exit(app.exec())
