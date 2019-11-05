# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'first.ui'
#
# Created by: PyQt5 UI code generator 5.13.0
#
# WARNING! All changes made in this file will be lost!
import sys

from PyQt5 import QtCore, QtGui, QtWidgets
from socket import *
import pymssql
from PyQt5.QtWidgets import QMessageBox, QApplication
import aas


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(553, 392)
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(340, 50, 75, 23))
        self.pushButton.setObjectName("pushButton")
        self.lineEdit = QtWidgets.QLineEdit(Dialog)
        self.lineEdit.setGeometry(QtCore.QRect(130, 50, 181, 20))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_2.setGeometry(QtCore.QRect(130, 120, 181, 20))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.pushButton_2 = QtWidgets.QPushButton(Dialog)
        self.pushButton_2.setGeometry(QtCore.QRect(340, 120, 75, 23))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3 = QtWidgets.QPushButton(Dialog)
        self.pushButton_3.setGeometry(QtCore.QRect(350, 270, 75, 23))
        self.pushButton_3.setObjectName("pushButton_3")
        self.textEdit = QtWidgets.QTextEdit(Dialog)
        self.textEdit.setGeometry(QtCore.QRect(140, 260, 181, 41))
        self.textEdit.setObjectName("textEdit")
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(60, 50, 71, 21))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(60, 120, 61, 21))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(50, 270, 71, 21))
        self.label_3.setObjectName("label_3")
        self.line = QtWidgets.QFrame(Dialog)
        self.line.setGeometry(QtCore.QRect(0, 190, 531, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.label_4 = QtWidgets.QLabel(Dialog)
        self.label_4.setGeometry(QtCore.QRect(60, 150, 441, 21))
        self.label_4.setObjectName("label_4")

        # 根据序列号查询
        self.pushButton.clicked.connect(lambda: self.findByID(self.lineEdit))
        # 根据日期查询
        self.pushButton_2.clicked.connect(lambda: self.findByDate(self.lineEdit_2))
        # 发送序列号工作
        self.pushButton_3.clicked.connect(lambda: self.workByID(self.textEdit))

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "查询工作系统"))
        self.pushButton.setText(_translate("Dialog", "查询"))
        self.pushButton_2.setText(_translate("Dialog", "查询"))
        self.pushButton_3.setText(_translate("Dialog", "提交"))
        self.label.setText(_translate("Dialog", "序列号："))
        self.label_2.setText(_translate("Dialog", "日期："))
        self.label_3.setText(_translate("Dialog", "工作序列号："))
        self.label_4.setText(_translate("Dialog", "请按照年-月-日格式输入。例如2019-02-03"))

    def findByID(self, textfile):
        data = textfile.text()
        if data == "":
            a = QMessageBox.warning(self.pushButton, "警告", "序列号不能为空！", QMessageBox.Yes)
        else:
            data = textfile.text()
            connect = pymssql.connect('(local)', 'sa', '123456', 'lCS_Common')  # 服务器名,账户,密码,数据库名
            if connect:
                print("连接成功!")
            cursor = connect.cursor()  # 创建一个游标对象,python里的sql语句都要通过cursor来执行
            sql = "SELECT * FROM tbl_EOR_Data WHERE id= %s" % (data)
            cursor.execute(sql)  # 执行sql语句
            row = cursor.fetchone()  # 读取查询结果
            if row == None:
                a = QMessageBox.warning(self.pushButton, "警告", "不存在该记录！", QMessageBox.Yes)
            else:
                while row:  # 循环读取所有结果
                    print("ID=%s, Date=%s" % (row[0], row[1]))  # 输出结果
                    row = cursor.fetchone()
            cursor.close()
            connect.close()

    def findByDate(self, textfile):
        data = textfile.text()
        if data == "":
            a = QMessageBox.warning(self.pushButton, "警告", "日期不能为空！", QMessageBox.Yes)
        else:
            connect = pymssql.connect('(local)', 'sa', '123456', 'lCS_Common')  # 服务器名,账户,密码,数据库名
            if connect:
                print("连接成功!")
            cursor = connect.cursor()  # 创建一个游标对象,python里的sql语句都要通过cursor来执行
            sql = "SELECT * FROM tbl_EOR_Data WHERE Convert(varchar,date,120) LIKE %s" % (data)
            sql = sql + '%'
            sql1 = sql[0:sql.rfind(' ')]
            sql2 = sql[sql.rfind(' ') + 1:]
            finalsql = sql1 + "'" + sql2 + "'"
            cursor.execute(finalsql)  # 执行sql语句
            row = cursor.fetchone()  # 读取查询结果,
            if row == None:
                a = QMessageBox.warning(self.pushButton, "警告", "不存在该记录！", QMessageBox.Yes)
            else:
                while row:  # 循环读取所有结果
                    print("ID=%s, Date=%s" % (row[0], row[1]))  # 输出结果
                    row = cursor.fetchone()
            cursor.close()
            connect.close()

    def workByID(self, texfile):
        try:
            HOST = '127.0.0.1'  # 服务端ip
            PORT = 21566  # 服务端端口号
            BUFSIZ = 1024
            ADDR = (HOST, PORT)
            tcpCliSock = socket(AF_INET, SOCK_STREAM)  # 创建socket对象
            tcpCliSock.connect(ADDR)  # 连接服务器
            data = texfile.toPlainText()
            tcpCliSock.send(data.encode('utf-8'))  # 发送消息
            data = tcpCliSock.recv(BUFSIZ)  # 读取消息
            print(data.decode('utf-8'))
            tcpCliSock.close()  # 关闭客户端
        except Exception:
            a = QMessageBox.warning(self.pushButton, "警告", "控制器未打开！", QMessageBox.Yes)
