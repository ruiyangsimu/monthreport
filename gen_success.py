# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ui/gen_success.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_GenSuccessDialog(object):
    def setupUi(self, GenSuccessDialog):
        GenSuccessDialog.setObjectName("GenSuccessDialog")
        GenSuccessDialog.setWindowModality(QtCore.Qt.WindowModal)
        GenSuccessDialog.resize(228, 75)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/images/logo.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        GenSuccessDialog.setWindowIcon(icon)
        self.label = QtWidgets.QLabel(GenSuccessDialog)
        self.label.setGeometry(QtCore.QRect(90, 10, 111, 51))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.retranslateUi(GenSuccessDialog)
        QtCore.QMetaObject.connectSlotsByName(GenSuccessDialog)

    def retranslateUi(self, GenSuccessDialog):
        _translate = QtCore.QCoreApplication.translate
        GenSuccessDialog.setWindowTitle(_translate("GenSuccessDialog", "结果"))
        self.label.setText(_translate("GenSuccessDialog", "成功"))


