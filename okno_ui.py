# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'okno.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(450, 193)
        Form.setMinimumSize(QtCore.QSize(450, 0))
        Form.setMaximumSize(QtCore.QSize(450, 16777215))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("logo.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        Form.setStyleSheet("")
        self.gridLayout = QtWidgets.QGridLayout(Form)
        self.gridLayout.setContentsMargins(5, 5, 5, 5)
        self.gridLayout.setSpacing(5)
        self.gridLayout.setObjectName("gridLayout")
        self.progressBar_1 = QtWidgets.QProgressBar(Form)
        self.progressBar_1.setMinimumSize(QtCore.QSize(0, 8))
        self.progressBar_1.setMaximumSize(QtCore.QSize(16777215, 8))
        self.progressBar_1.setStyleSheet("QProgressBar::chunk {\n"
"    background-color: rgb(100, 150, 150);\n"
"    margin: 2px;\n"
"}")
        self.progressBar_1.setMinimum(0)
        self.progressBar_1.setMaximum(100)
        self.progressBar_1.setProperty("value", 0)
        self.progressBar_1.setAlignment(QtCore.Qt.AlignCenter)
        self.progressBar_1.setTextVisible(False)
        self.progressBar_1.setTextDirection(QtWidgets.QProgressBar.TopToBottom)
        self.progressBar_1.setObjectName("progressBar_1")
        self.gridLayout.addWidget(self.progressBar_1, 4, 0, 1, 1)
        self.plainTextEdit = QtWidgets.QPlainTextEdit(Form)
        self.plainTextEdit.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.plainTextEdit.sizePolicy().hasHeightForWidth())
        self.plainTextEdit.setSizePolicy(sizePolicy)
        self.plainTextEdit.setMinimumSize(QtCore.QSize(400, 50))
        self.plainTextEdit.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.gridLayout.addWidget(self.plainTextEdit, 0, 0, 1, 1)
        self.label = QtWidgets.QLabel(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        self.label.setMinimumSize(QtCore.QSize(0, 25))
        self.label.setMaximumSize(QtCore.QSize(16777215, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.label.setFont(font)
        self.label.setFrameShape(QtWidgets.QFrame.Box)
        self.label.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.label.setText("")
        self.label.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignTop)
        self.label.setWordWrap(True)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 2, 0, 1, 1)
        self.tableWidget = QtWidgets.QTableWidget(Form)
        self.tableWidget.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.tableWidget.sizePolicy().hasHeightForWidth())
        self.tableWidget.setSizePolicy(sizePolicy)
        self.tableWidget.setMinimumSize(QtCore.QSize(400, 55))
        self.tableWidget.setMaximumSize(QtCore.QSize(16777215, 55))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.tableWidget.setFont(font)
        self.tableWidget.setFocusPolicy(QtCore.Qt.NoFocus)
        self.tableWidget.setStyleSheet("")
        self.tableWidget.setInputMethodHints(QtCore.Qt.ImhNone)
        self.tableWidget.setFrameShape(QtWidgets.QFrame.Box)
        self.tableWidget.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.tableWidget.setLineWidth(1)
        self.tableWidget.setMidLineWidth(0)
        self.tableWidget.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.tableWidget.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.tableWidget.setAutoScrollMargin(16)
        self.tableWidget.setAlternatingRowColors(False)
        self.tableWidget.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.tableWidget.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectItems)
        self.tableWidget.setTextElideMode(QtCore.Qt.ElideLeft)
        self.tableWidget.setShowGrid(True)
        self.tableWidget.setGridStyle(QtCore.Qt.SolidLine)
        self.tableWidget.setWordWrap(True)
        self.tableWidget.setCornerButtonEnabled(False)
        self.tableWidget.setRowCount(2)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(4)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        item.setFont(font)
        brush = QtGui.QBrush(QtGui.QColor(240, 240, 240))
        brush.setStyle(QtCore.Qt.SolidPattern)
        item.setBackground(brush)
        item.setFlags(QtCore.Qt.NoItemFlags)
        self.tableWidget.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        brush = QtGui.QBrush(QtGui.QColor(240, 240, 240))
        brush.setStyle(QtCore.Qt.SolidPattern)
        item.setBackground(brush)
        self.tableWidget.setItem(0, 1, item)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        brush = QtGui.QBrush(QtGui.QColor(240, 240, 240))
        brush.setStyle(QtCore.Qt.SolidPattern)
        item.setBackground(brush)
        self.tableWidget.setItem(0, 2, item)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        brush = QtGui.QBrush(QtGui.QColor(240, 240, 240))
        brush.setStyle(QtCore.Qt.SolidPattern)
        item.setBackground(brush)
        self.tableWidget.setItem(0, 3, item)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        brush = QtGui.QBrush(QtGui.QColor(254, 254, 254))
        brush.setStyle(QtCore.Qt.SolidPattern)
        item.setBackground(brush)
        self.tableWidget.setItem(1, 0, item)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        brush = QtGui.QBrush(QtGui.QColor(254, 254, 254))
        brush.setStyle(QtCore.Qt.SolidPattern)
        item.setBackground(brush)
        self.tableWidget.setItem(1, 1, item)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        brush = QtGui.QBrush(QtGui.QColor(254, 254, 254))
        brush.setStyle(QtCore.Qt.SolidPattern)
        item.setBackground(brush)
        self.tableWidget.setItem(1, 2, item)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        brush = QtGui.QBrush(QtGui.QColor(254, 254, 254))
        brush.setStyle(QtCore.Qt.SolidPattern)
        item.setBackground(brush)
        self.tableWidget.setItem(1, 3, item)
        self.tableWidget.horizontalHeader().setVisible(False)
        self.tableWidget.horizontalHeader().setCascadingSectionResizes(False)
        self.tableWidget.horizontalHeader().setDefaultSectionSize(50)
        self.tableWidget.horizontalHeader().setMinimumSectionSize(50)
        self.tableWidget.horizontalHeader().setSortIndicatorShown(False)
        self.tableWidget.horizontalHeader().setStretchLastSection(True)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.verticalHeader().setCascadingSectionResizes(False)
        self.tableWidget.verticalHeader().setDefaultSectionSize(25)
        self.tableWidget.verticalHeader().setMinimumSectionSize(25)
        self.tableWidget.verticalHeader().setSortIndicatorShown(False)
        self.tableWidget.verticalHeader().setStretchLastSection(True)
        self.gridLayout.addWidget(self.tableWidget, 1, 0, 1, 1)
        self.pushButton = QtWidgets.QPushButton(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton.sizePolicy().hasHeightForWidth())
        self.pushButton.setSizePolicy(sizePolicy)
        self.pushButton.setMinimumSize(QtCore.QSize(0, 25))
        self.pushButton.setMaximumSize(QtCore.QSize(16777215, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.gridLayout.addWidget(self.pushButton, 5, 0, 1, 1)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.plainTextEdit.setPlainText(_translate("Form", "C:\\vxvproj\\Smets_for_Subpodryd\\xxx\\2\\1394 16.03.23 Сапсан"))
        self.plainTextEdit.setPlaceholderText(_translate("Form", "Перетащите сюда или скопируйте адрес папки с файлами для переименования и раскладки по структуре папок для модуля Субподряд. Комплекный объект заполняется в формате:  ХХХ.ХХХ.ХХХ"))
        self.tableWidget.setSortingEnabled(False)
        item = self.tableWidget.verticalHeaderItem(0)
        item.setText(_translate("Form", "1"))
        item = self.tableWidget.verticalHeaderItem(1)
        item.setText(_translate("Form", "2"))
        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("Form", "Шифр"))
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("Form", "Стадия"))
        item = self.tableWidget.horizontalHeaderItem(2)
        item.setText(_translate("Form", "Комплексный объект"))
        item = self.tableWidget.horizontalHeaderItem(3)
        item.setText(_translate("Form", "Марка"))
        __sortingEnabled = self.tableWidget.isSortingEnabled()
        self.tableWidget.setSortingEnabled(False)
        item = self.tableWidget.item(0, 0)
        item.setText(_translate("Form", "Шифр"))
        item = self.tableWidget.item(0, 1)
        item.setText(_translate("Form", "Стадия"))
        item = self.tableWidget.item(0, 2)
        item.setText(_translate("Form", "Комплексный объект"))
        item = self.tableWidget.item(0, 3)
        item.setText(_translate("Form", "Марка"))
        item = self.tableWidget.item(1, 1)
        item.setText(_translate("Form", "Р"))
        item = self.tableWidget.item(1, 3)
        item.setText(_translate("Form", "СМ-01"))
        self.tableWidget.setSortingEnabled(__sortingEnabled)
        self.pushButton.setText(_translate("Form", "Start"))
        self.pushButton.setShortcut(_translate("Form", "Return"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())