# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.

import os
from convert import handle_xmind,handle_title,handle_topics,write_to_temp1,write_to_temp_jira,write_to_temp2
from PyQt5 import QtCore, QtWidgets



class Ui_MainWindow(object):
    fileName = ''
    filePath = ''
    template = ''

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 450)
        MainWindow.setMinimumSize(QtCore.QSize(800, 450))
        MainWindow.setMaximumSize(QtCore.QSize(800, 450))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(410, 150, 341, 31))
        self.pushButton.setObjectName("pushButton")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(160, 41, 471, 31))
        self.lineEdit.setTabletTracking(False)
        self.lineEdit.setStatusTip("")
        self.lineEdit.setWhatsThis("")
        self.lineEdit.setInputMask("")
        self.lineEdit.setText("")
        self.lineEdit.setFrame(True)
        self.lineEdit.setEchoMode(QtWidgets.QLineEdit.Normal)
        self.lineEdit.setDragEnabled(False)
        self.lineEdit.setReadOnly(True)
        self.lineEdit.setObjectName("lineEdit")
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(40, 240, 711, 131))
        self.textEdit.setObjectName("textEdit")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(660, 90, 91, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(160, 90, 471, 31))
        self.lineEdit_2.setInputMask("")
        self.lineEdit_2.setReadOnly(True)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(40, 40, 81, 31))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(40, 90, 81, 31))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(40, 210, 81, 31))
        self.label_3.setObjectName("label_3")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(660, 40, 91, 31))
        self.pushButton_3.setObjectName("pushButton_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(40, 150, 101, 31))
        self.label_4.setObjectName("label_4")
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(160, 150, 191, 31))
        self.comboBox.setEditable(False)
        self.comboBox.setDuplicatesEnabled(False)
        self.comboBox.setObjectName("comboBox")
        self.comboBox.addItem("禅道模板")
        self.comboBox.addItem("JIRA模板")
        self.comboBox.addItem("集成测试模板")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        # 信号，选择xmind文件
        self.pushButton_3.clicked.connect(self.open_file)
        # 信号，选择导出路径
        self.pushButton_2.clicked.connect(self.open_filepath)
        # 信号，把取到的值传给convert槽进行转换用例
        self.template = self.comboBox.currentText()
        self.comboBox.activated.connect(self.select_template)
        self.pushButton.clicked.connect(lambda: self.run(self.fileName, self.filePath, self.template))

    # 槽函数
    def open_file(self):
        fileName, fileType = QtWidgets.QFileDialog.getOpenFileName(self, "选取xmind文件", os.getcwd())
        self.fileName = fileName
        self.lineEdit.setText(fileName)

    def open_filepath(self):
        filePath = QtWidgets.QFileDialog.getExistingDirectory(self, "选取输出文件夹", os.getcwd())
        self.filePath = filePath
        self.lineEdit_2.setText(filePath)

    def select_template(self):
        template = self.comboBox.currentText()
        self.template = template

    def run(self, xmind_file, out_filepath, temp):
        if xmind_file == "":
            QtWidgets.QMessageBox.warning(self, '警告', '源文件路径不能为空', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)
        elif xmind_file.endswith('.xmind') is False:
            QtWidgets.QMessageBox.warning(self, '警告', '请选择xmind格式的用例', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)
        else:
            xmind_name = os.path.basename(xmind_file)[:-6]
            # print("xmind文件路径:", xmind_file)
            if out_filepath == '':
                excel_file = os.path.join(os.path.dirname(xmind_file), xmind_name + '.xls')
            else:
                excel_file = os.path.join(out_filepath, xmind_name + '.xls')
            # print("excel输出路径:", excel_file)
            res = handle_topics(handle_xmind(xmind_file))
            if temp == "禅道模板":
                write_to_temp1(handle_title(res), excel_file)
            elif temp == "JIRA模板":
                write_to_temp_jira(handle_title(res), excel_file)
            elif temp == "集成测试模板":
                write_to_temp2(handle_title(res), excel_file)
            # print("转换完成！")
            self.textEdit.setText(f"xmind文件路径:{xmind_file}" + '\n' + f"excel输出路径:{excel_file}" + '\n' + "转换成功")

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton.setText(_translate("MainWindow", "开始转换"))
        self.pushButton_2.setText(_translate("MainWindow", "选择"))
        self.label.setText(_translate("MainWindow", "源文件路径："))
        self.label_2.setText(_translate("MainWindow", "输出路径："))
        self.label_3.setText(_translate("MainWindow", "运行结果："))
        self.pushButton_3.setText(_translate("MainWindow", "选择"))
        self.label_4.setText(_translate("MainWindow", "选择输出模板："))
        self.comboBox.setItemText(0, _translate("MainWindow", "禅道模板"))
        self.comboBox.setItemText(1, _translate("MainWindow", "JIRA模板"))
        self.comboBox.setItemText(2, _translate("MainWindow", "集成测试模板"))