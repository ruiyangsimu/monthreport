import logging
import sys

import yaml
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QDialog, QMainWindow, QMessageBox
from PyQt5.uic import loadUi

from Picture import Picture
from gen_pic import Ui_Gen_Pic
from main_window_ui import Ui_MainWindow

with open('log.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)

logger = logging.getLogger("app")


class Window(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.connectSignalsSlots()

    def connectSignalsSlots(self):
        self.action_Exit.triggered.connect(self.close)
        # self.action_Find_Replace.triggered.connect(self.findAndReplace)
        self.action_genword.triggered.connect(self.genDocProcessBar)
        self.action_GenPic.triggered.connect(self.genPic)
        self.action_About.triggered.connect(self.about)

    def genDocProcessBar(self):
        # processBar = CommonProgressBar()
        # processBar.show()
        pass

    def findAndReplace(self):
        dialog = FindReplaceDialog(self)
        dialog.exec()

    def genPic(self):
        # pic = Picture()
        # processBar = CommonProgressBar(10, pic, "图片生成")
        # genPicThread = GenPicThread(pic)
        # genPicThread._signal.connect(processBar.signal_accept)
        # processBar.show()
        # genPicThread.start()
        self.window = QtWidgets.QMainWindow()
        self.gen_pic = Ui_Gen_Pic()
        self.gen_pic.setupUi(self.window)
        self.gen_pic.setData(Picture())
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
