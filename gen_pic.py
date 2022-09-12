# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ui/gen_pic.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.
import logging

import yaml
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal

from gen_success import Ui_GenSuccessDialog

with open('log.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)

logger = logging.getLogger("uigenpic")


class Ui_Gen_Pic(object):
    def setupUi(self, Gen_Pic):
        Gen_Pic.setObjectName("Gen_Pic")
        Gen_Pic.setWindowModality(QtCore.Qt.WindowModal)
        Gen_Pic.resize(429, 105)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/images/logo.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Gen_Pic.setWindowIcon(icon)
        self.progressBar = QtWidgets.QProgressBar(Gen_Pic)
        self.progressBar.setGeometry(QtCore.QRect(20, 20, 401, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setInvertedAppearance(False)
        self.progressBar.setObjectName("progressBar")
        self.pushButton = QtWidgets.QPushButton(Gen_Pic)
        self.pushButton.setGeometry(QtCore.QRect(160, 60, 81, 23))
        self.pushButton.setObjectName("pushButton")

        self.retranslateUi(Gen_Pic)
        QtCore.QMetaObject.connectSlotsByName(Gen_Pic)

    def retranslateUi(self, Gen_Pic):
        _translate = QtCore.QCoreApplication.translate
        Gen_Pic.setWindowTitle(_translate("Gen_Pic", "生成图片"))
        self.pushButton.setText(_translate("Gen_Pic", "生成图片"))

    def setData(self, picture):
        self.picture = picture
        self.picture.load()
        self.progressBar.setMinimum(0)
        self.progressBar.setMaximum(self.picture.get_num())
        self.pushButton.clicked.connect(self.btnFunc)

    def btnFunc(self):
        self.thread = GenPicThread(self.picture)
        self.thread._signal.connect(self.signal_accept)
        self.thread.start()
        self.pushButton.setEnabled(False)

    def signal_accept(self, msg):
        if (self.progressBar.value() == self.picture.get_num()):
            self.progressBar.setValue(0)
        logger.debug("accect signal : %s", msg)
        value = self.progressBar.value() + int(msg)
        logger.debug("signal value: %s", value)
        self.progressBar.setValue(value)
        if self.progressBar.value() == self.picture.get_num():
            self.window = QtWidgets.QMainWindow()
            self.gsuccess = Ui_GenSuccessDialog()
            self.gsuccess.setupUi(self.window)
            self.window.show()
            self.pushButton.setEnabled(True)
            _translate = QtCore.QCoreApplication.translate
            self.pushButton.setText(_translate("Gen_Pic", "重新生成"))


class GenPicThread(QThread):
    _signal = pyqtSignal(int)

    def __init__(self, pic):
        super(GenPicThread, self).__init__()
        self.pic = pic

    def __del__(self):
        self.wait()

    def run(self):
        for name in self.pic.get_product_name():
            self.pic.gen(name)
            logger.debug("%s gen ok", name)
            self._signal.emit(1)


