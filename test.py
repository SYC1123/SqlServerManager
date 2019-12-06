import os
import sys

from PyQt5.uic.properties import QtWidgets, QtCore

import first
from PyQt5.QtWidgets import *
from socket import *
import pymssql
import time
import xlwt
from PyQt5.QtCore import QObject, pyqtSignal, QThread, QDateTime

# 设置表格样式
import second


def set_style(name, height, bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style


def findByIDONE(button, textfile):
    data = textfile.text()
    if data == "":
        a = QMessageBox.warning(button, "警告", "型号不能为空！", QMessageBox.Yes)
    else:
        data = textfile.text()
        connect = pymssql.connect('(local)', 'sa', '123456', 'lCS_Common')  # 服务器名,账户,密码,数据库名
        if connect:
            print("连接成功!")
        cursor = connect.cursor()  # 创建一个游标对象,python里的sql语句都要通过cursor来执行

        sql = "SELECT * FROM tbl_EOR_Data WHERE fld_Barcode Like %" + "%s" % (data)
        sql = sql + '%'
        sql1 = sql[0:sql.rfind(' ')]
        sql2 = sql[sql.rfind(' ') + 1:]
        finalsql = sql1 + "'" + sql2 + "'"
        print(finalsql)
        cursor.execute(finalsql)  # 执行sql语句
        row = cursor.fetchone()  # 读取查询结果
        if row == None:
            a = QMessageBox.warning(button, "警告", "不存在该记录！", QMessageBox.Yes)
        else:
            count = 1
            f = xlwt.Workbook()
            sheet1 = f.add_sheet('型号号找到的工件', cell_overwrite_ok=True)
            while row:  # 循环读取所有结果
                # 写第一行
                sheet1.write(0, 0, "台次号", set_style('Times New Roman', 220, True))
                sheet1.write(0, 1, "时间", set_style('Times New Roman', 220, True))
                sheet1.write(0, 2, "实测扭矩", set_style('Times New Roman', 220, True))
                sheet1.write(0, 3, "规定扭矩", set_style('Times New Roman', 220, True))
                sheet1.write(0, 4, "总群组计数", set_style('Times New Roman', 220, True))
                sheet1.write(0, 5, "当前群组计数", set_style('Times New Roman', 220, True))
                # 写一行数据
                for a in range(0, 6):
                    if a == 0:
                        sheet1.col(a).width = 256 * 40
                    else:
                        sheet1.col(a).width = 256 * 20
                    sheet1.write(count, a, str(row[a]))
                count = count + 1
                row = cursor.fetchone()
            ticks = time.time()
            f.save(os.path.join(os.path.expanduser('~'), "Desktop") + "\\" + str(ticks) + "型号号.xls")
            a = QMessageBox.warning(button, "", "查询成功！", QMessageBox.Yes)
        cursor.close()
        connect.close()


# 台次号
def findByID(button, textfile):
    data = textfile.text()
    if data == "":
        a = QMessageBox.warning(button, "警告", "台次号不能为空！", QMessageBox.Yes)
    else:
        data = textfile.text()
        connect = pymssql.connect('(local)', 'sa', '123456', 'lCS_Common')  # 服务器名,账户,密码,数据库名
        if connect:
            print("连接成功!")
        cursor = connect.cursor()  # 创建一个游标对象,python里的sql语句都要通过cursor来执行

        sql = "SELECT * FROM tbl_EOR_Data WHERE fld_Barcode Like %s" % (data)
        sql = sql + '%'
        sql1 = sql[0:sql.rfind(' ')]
        sql2 = sql[sql.rfind(' ') + 1:]
        finalsql = sql1 + "'" + sql2 + "'"
        print(finalsql)
        cursor.execute(finalsql)  # 执行sql语句
        row = cursor.fetchone()  # 读取查询结果
        if row == None:
            a = QMessageBox.warning(button, "警告", "不存在该记录！", QMessageBox.Yes)
        else:
            count = 1
            f = xlwt.Workbook()
            sheet1 = f.add_sheet('台次号找到的工件', cell_overwrite_ok=True)
            while row:  # 循环读取所有结果
                # 写第一行
                sheet1.write(0, 0, "台次号", set_style('Times New Roman', 220, True))
                sheet1.write(0, 1, "时间", set_style('Times New Roman', 220, True))
                sheet1.write(0, 2, "实测扭矩", set_style('Times New Roman', 220, True))
                sheet1.write(0, 3, "规定扭矩", set_style('Times New Roman', 220, True))
                sheet1.write(0, 4, "总群组计数", set_style('Times New Roman', 220, True))
                sheet1.write(0, 5, "当前群组计数", set_style('Times New Roman', 220, True))
                # 写一行数据
                for a in range(0, 6):
                    if a == 0:
                        sheet1.col(a).width = 256 * 40
                    else:
                        sheet1.col(a).width = 256 * 20
                    sheet1.write(count, a, str(row[a]))
                count = count + 1
                row = cursor.fetchone()
            ticks = time.time()
            f.save(os.path.join(os.path.expanduser('~'), "Desktop") + "\\" + str(ticks) + "台次号.xls")
            a = QMessageBox.warning(button, "", "查询成功！", QMessageBox.Yes)
        cursor.close()
        connect.close()


# 日期
def findByDate(button, textfile):
    global f1
    data = textfile.text()
    print(data)
    if data == "":
        a = QMessageBox.warning(button, "警告", "日期不能为空！", QMessageBox.Yes)
    else:
        connect = pymssql.connect('(local)', 'sa', '123456', 'lCS_Common')  # 服务器名,账户,密码,数据库名
        if connect:
            print("连接成功!")
        cursor = connect.cursor()  # 创建一个游标对象,python里的sql语句都要通过cursor来执行

        sql = "SELECT * FROM tbl_EOR_Data WHERE Convert(varchar,fld_Date_Time_Of_Cycle,120) LIKE %s" % (data)
        sql = sql + '%'
        sql1 = sql[0:sql.rfind(' ')]
        sql2 = sql[sql.rfind(' ') + 1:]
        finalsql = sql1 + "'" + sql2 + "'"
        cursor.execute(finalsql)  # 执行sql语句
        row = cursor.fetchone()  # 读取查询结果,
        if row == None:
            a = QMessageBox.warning(button, "警告", "不存在该记录！", QMessageBox.Yes)
        else:
            co = 1
            f1 = xlwt.Workbook()
            sheet = f1.add_sheet('日期找到的工件', cell_overwrite_ok=True)
            while row:  # 循环读取所有结果
                # 写第一行
                # for i in range(0, len(new_list)):
                #     sheet.write(0, i, new_list[i], set_style('Times New Roman', 220, True))  # 行 列
                sheet.write(0, 0, "台次号", set_style('Times New Roman', 220, True))
                sheet.write(0, 1, "时间", set_style('Times New Roman', 220, True))
                sheet.write(0, 2, "实测扭矩", set_style('Times New Roman', 220, True))
                sheet.write(0, 3, "规定扭矩", set_style('Times New Roman', 220, True))
                sheet.write(0, 4, "总群组计数", set_style('Times New Roman', 220, True))
                sheet.write(0, 5, "当前群组计数", set_style('Times New Roman', 220, True))
                # 写一行数据
                for a in range(0, 6):
                    if a == 0:
                        sheet.col(a).width = 256 * 40
                    else:
                        sheet.col(a).width = 256 * 20
                    sheet.write(co, a, str(row[a]))
                co = co + 1
                print()
                row = cursor.fetchone()
            ticks = time.time()
            f1.save(os.path.join(os.path.expanduser('~'), "Desktop") + "\\" + str(ticks) + "日期.xls")
            a = QMessageBox.warning(button, "", "查询成功！", QMessageBox.Yes)
        cursor.close()
        connect.close()


class Window(QDialog):
    def __init__(self, taici, xinghao):
        QDialog.__init__(self)
        self.taici1 = taici
        self.xinghao1 = xinghao
        self.setWindowTitle('扫码工作')
        self.resize(500, 250)
        self.input = QLineEdit(self)
        self.input.setGeometry(0, 0, 0, 0)
        self.label = QLabel(self)
        self.label.setGeometry(160, 130, 200, 50)
        self.label.setText("注意打开控制器")
        self.input.selectAll()
        self.input.setFocus()
        self.input.returnPressed.connect(self.updateUi)

    def updateUi(self):
        # print(self.input.text() + '\n')
        print("------------------------------------------------")
        if self.input.text() == '':
            a = QMessageBox.warning(None, "警告", "序列号不能为空！", QMessageBox.Yes)
        else:
            try:
                # HOST = '192.168.4.4'  # 服务端ip
                # PORT = 41000  # 服务端端口号
                HOST = '127.0.0.1'
                PORT = 21566
                BUFSIZ = 10241
                ADDR = (HOST, PORT)
                tcpCliSock = socket(AF_INET, SOCK_STREAM)  # 创建socket对象
                tcpCliSock.connect(ADDR)  # 连接服务器
                data = self.taici1 + "QQQQQQ" + self.xinghao1 + "QQQQQQ" + self.input.text() + '\n'
                # data=self.input.text() + '\n'
                print(data)
                tcpCliSock.send(data.encode('utf-8'))  # 发送消息
                data = tcpCliSock.recv(BUFSIZ)  # 读取消息
                print(data.decode('utf-8'))
                tcpCliSock.close()  # 关闭客户端
            except Exception:
                a = QMessageBox.warning(None, "警告", "控制器未打开！", QMessageBox.Yes)
        self.input.setText("")


class parentWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.main_ui = first.Ui_Main()
        self.main_ui.setupUi(self)


def startwork(win, taici, xinghao):
    if taici == '' or xinghao == '':
        a = QMessageBox.warning(None, "警告", "台次号或者型号不能为空！", QMessageBox.Yes)
    else:
        win.taici1=taici
        win.xinghao1=xinghao
        win.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    first = parentWindow()
    win = Window('','')

    # 根据台次号查询
    first.main_ui.pushButton.clicked.connect(lambda: findByID(first.main_ui.pushButton, first.main_ui.lineEdit))
    # 根据日期查询
    first.main_ui.pushButton_2.clicked.connect(lambda: findByDate(first.main_ui.pushButton_2, first.main_ui.lineEdit_2))
    # 发送序列号工作
    first.main_ui.pushButton_3.clicked.connect(
        lambda: startwork(win, first.main_ui.lineEdit_4.text(), first.main_ui.lineEdit_5.text()))
    # 根据型号查询
    first.main_ui.pushButton_4.clicked.connect(
        lambda: findByIDONE(first.main_ui.pushButton_4, first.main_ui.lineEdit_3))

    # 显示
    first.show()
    sys.exit(app.exec())
