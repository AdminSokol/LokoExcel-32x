import time
from PyQt5.QtCore import QTimer
from PyQt5 import uic, QtGui
from PyQt5.QtWidgets import QApplication, QDialog
from datetime import datetime, timedelta, date
import openpyxl
import threading

stop = 0
wb = openpyxl.load_workbook("lokofile.xlsx")
Form, Window = uic.loadUiType("lokodrom.ui")
app = QApplication([])
window = Window()
form = Form()
form.setupUi(window)
window.show()
today = datetime.today()
sheet = wb.active
pause_flag_1 = sheet['S2'].value  # 0
start_flag_1 = sheet['T2'].value  # 1
pause_flag_2 = sheet['S3'].value  # 0
start_flag_2 = sheet['T3'].value  # 1
pause_flag_3 = sheet['S4'].value  # 0
start_flag_3 = sheet['T4'].value  # 1
pause_flag_4 = sheet['S5'].value  # 0
start_flag_4 = sheet['T5'].value  # 1
pause_flag_5 = sheet['S6'].value  # 0
start_flag_5 = sheet['T6'].value  # 1
pause_flag_6 = sheet['S7'].value  # 0
start_flag_6 = sheet['T7'].value  # 1
flag = 0


def Reset1():
    sheet["M2"].value = str(int(sheet['R2'].value) + int(sheet["M2"].value))
    sheet['R2'].value = "0"
    sheet['P2'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(0, 2).setText(str("0h 0m"))
    form.progressBar.setProperty("value", 0)


def Reset2():
    sheet['M3'].value = str(int(sheet['R3'].value) + int(sheet['M3'].value))
    sheet['R3'].value = "0"
    sheet['P3'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(2, 2).setText(str("0h 0m"))
    form.progressBar_2.setProperty("value", 0)


def Reset3():
    sheet['M4'].value = str(int(sheet['R4'].value) + int(sheet['M4'].value))
    sheet['R4'].value = "0"
    sheet['P4'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(4, 2).setText(str("0h 0m"))
    form.progressBar_3.setProperty("value", 0)


def Reset4():
    sheet['M5'].value = str(int(sheet['R5'].value) + int(sheet['M5'].value))
    sheet['R5'].value = "0"
    sheet['P5'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(6, 2).setText(str("0h 0m"))
    form.progressBar_4.setProperty("value", 0)


def Reset5():
    sheet['M6'].value = str(int(sheet['R6'].value) + int(sheet['M6'].value))
    sheet['R6'].value = "0"
    sheet['P6'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(8, 2).setText(str("0h 0m"))
    form.progressBar_5.setProperty("value", 0)


def Reset6():
    sheet['M7'].value = str(int(sheet['R7'].value) + int(sheet['M7'].value))
    sheet['R7'].value = "0"
    sheet['P7'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(10, 2).setText(str("0h 0m"))
    form.progressBar_6.setProperty("value", 0)


def Reset7():
    sheet['L2'].value = str(int(sheet['Q2'].value) + int(sheet['L2'].value))
    sheet['Q2'].value = "0"
    sheet['O2'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(1, 2).setText(str("0h 0m"))
    form.progressBar_7.setProperty("value", 0)


def Reset8():
    sheet['L3'].value = str(int(sheet['Q3'].value) + int(sheet['L3'].value))
    sheet['Q3'].value = "0"
    sheet['O3'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(3, 2).setText(str("0h 0m"))
    form.progressBar_8.setProperty("value", 0)


def Reset9():
    sheet['L4'].value = str(int(sheet['Q4'].value) + int(sheet['L4'].value))
    sheet['Q4'].value = "0"
    sheet['O4'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(5, 2).setText(str("0h 0m"))
    form.progressBar_9.setProperty("value", 0)


def Reset10():
    sheet['L5'].value = str(int(sheet['Q5'].value) + int(sheet['L5'].value))
    sheet['Q5'].value = "0"
    sheet['O5'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(7, 2).setText(str("0h 0m"))
    form.progressBar_10.setProperty("value", 0)


def Reset11():
    sheet['L6'].value = str(int(sheet['Q6'].value) + int(sheet['L6'].value))
    sheet['Q6'].value = "0"
    sheet['O6'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(9, 2).setText(str("0h 0m"))
    form.progressBar_11.setProperty("value", 0)


def Reset12():
    sheet['L7'].value = str(int(sheet['Q7'].value) + int(sheet['L7'].value))
    sheet['Q7'].value = "0"
    sheet['O7'].value = "0"
    wb.save("lokofile.xlsx")
    form.tableWidget.item(11, 2).setText(str("0h 0m"))
    form.progressBar_12.setProperty("value", 0)


def button():
    form.pushButton_10.clicked.connect(button_clicked1)
    form.pushButton_9.clicked.connect(button_clicked2)
    form.pushButton_6.clicked.connect(button_clicked3)
    form.pushButton_8.clicked.connect(button_clicked4)
    form.pushButton_5.clicked.connect(button_clicked5)
    form.pushButton_7.clicked.connect(button_clicked6)
    form.pushButton_21.clicked.connect(button_clicked7)
    form.pushButton_22.clicked.connect(button_clicked8)
    form.pushButton_23.clicked.connect(button_clicked9)
    form.pushButton_24.clicked.connect(button_clicked10)
    form.pushButton_25.clicked.connect(button_clicked11)
    form.pushButton_26.clicked.connect(button_clicked12)


def pause1():
    global pause_flag_1, start_flag_1
    if pause_flag_1 == 1:
        return
    pause_flag_1 = 1
    start_flag_1 = 0
    sheet['S2'].value = pause_flag_1
    sheet['T2'].value = start_flag_1
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(1, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(0, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_15.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189);color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.pushButton_3.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_7.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_1.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def pause2():
    global pause_flag_2, start_flag_2
    if pause_flag_2 == 1:
        return
    pause_flag_2 = 1
    start_flag_2 = 0
    sheet['S3'].value = pause_flag_2
    sheet['T3'].value = start_flag_2
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(3, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(2, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_16.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189);color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.pushButton_4.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_8.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_2.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def pause3():
    global pause_flag_3, start_flag_3
    if pause_flag_3 == 1:
        return
    pause_flag_3 = 1
    start_flag_3 = 0
    sheet['S4'].value = pause_flag_3
    sheet['T4'].value = start_flag_3
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(5, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(4, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_17.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189);color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255)}")
    form.pushButton_11.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_9.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_3.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def pause4():
    global pause_flag_4, start_flag_4
    if pause_flag_4 == 1:
        return
    pause_flag_4 = 1
    start_flag_4 = 0
    sheet['S5'].value = pause_flag_4
    sheet['T5'].value = start_flag_4
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(7, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(6, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_18.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189);color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255)}")
    form.pushButton_12.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_10.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_4.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def pause5():
    global pause_flag_5, start_flag_5
    if pause_flag_5 == 1:
        return
    pause_flag_5 = 1
    start_flag_5 = 0
    sheet['S6'].value = pause_flag_5
    sheet['T6'].value = start_flag_5
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(9, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(8, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_19.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189);color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255)}")
    form.pushButton_13.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_11.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_5.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def pause6():
    global pause_flag_6, start_flag_6
    if pause_flag_6 == 1:
        return
    pause_flag_6 = 1
    start_flag_6 = 0
    sheet['S7'].value = pause_flag_6
    sheet['T7'].value = start_flag_6
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(11, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(10, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_20.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189);color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255)}")
    form.pushButton_14.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_12.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_6.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def button_pause():
    form.pushButton_3.clicked.connect(pause1)
    form.pushButton_4.clicked.connect(pause2)
    form.pushButton_11.clicked.connect(pause3)
    form.pushButton_12.clicked.connect(pause4)
    form.pushButton_13.clicked.connect(pause5)
    form.pushButton_14.clicked.connect(pause6)


def start1():
    global start_flag_1, pause_flag_1
    if start_flag_1 == 1:
        return
    start_flag_1 = 1
    pause_flag_1 = 0
    sheet['S2'].value = pause_flag_1
    sheet['T2'].value = start_flag_1
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(0, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(1, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_3.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189); color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.pushButton_15.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_1.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_7.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def start2():
    global start_flag_2, pause_flag_2
    if start_flag_2 == 1:
        return
    start_flag_2 = 1
    pause_flag_2 = 0
    sheet['S3'].value = pause_flag_2
    sheet['T3'].value = start_flag_2
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(2, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(3, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_4.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189); color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.pushButton_16.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_2.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_8.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def start3():
    global start_flag_3, pause_flag_3
    if start_flag_3 == 1:
        return
    start_flag_3 = 1
    pause_flag_3 = 0
    sheet['S4'].value = pause_flag_3
    sheet['T4'].value = start_flag_3
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(4, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(5, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_11.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189); color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.pushButton_17.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_3.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_9.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def start4():
    global start_flag_4, pause_flag_4
    if start_flag_4 == 1:
        return
    start_flag_4 = 1
    pause_flag_4 = 0
    sheet['S5'].value = pause_flag_4
    sheet['T5'].value = start_flag_4
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(6, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(7, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_12.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189); color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.pushButton_18.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_4.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_10.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def start5():
    global start_flag_5, pause_flag_5
    if start_flag_5 == 1:
        return
    start_flag_5 = 1
    pause_flag_5 = 0
    sheet['S6'].value = pause_flag_5
    sheet['T6'].value = start_flag_5
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(8, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(9, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_13.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189); color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.pushButton_19.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_5.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_11.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def start6():
    global start_flag_6, pause_flag_6
    if start_flag_6 == 1:
        return
    start_flag_6 = 1
    pause_flag_6 = 0
    sheet['S7'].value = pause_flag_6
    sheet['T7'].value = start_flag_6
    wb.save("lokofile.xlsx")
    for i in range(4):
        form.tableWidget.item(10, i).setBackground(QtGui.QColor(85, 255, 127))
        form.tableWidget.item(11, i).setBackground(QtGui.QColor(255, 161, 148))
    form.pushButton_14.setStyleSheet(
        "QPushButton{background-color: rgb(202, 201, 189); color: red;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.pushButton_20.setStyleSheet(
        "QPushButton{background-color: rgb(240, 240, 240); color: green;border: 1px solid black;border-radius: 4px} "
        "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
    form.spinBox_6.setStyleSheet(
        "QSpinBox {background-color: rgb(85,255,127);font-size: 12pt; font-family: Times New Roman;}")
    form.spinBox_12.setStyleSheet(
        "QSpinBox {background-color: rgb(255,161,148);font-size: 12pt; font-family: Times New Roman;}")


def button_start():
    form.pushButton_15.clicked.connect(start1)
    form.pushButton_16.clicked.connect(start2)
    form.pushButton_17.clicked.connect(start3)
    form.pushButton_18.clicked.connect(start4)
    form.pushButton_19.clicked.connect(start5)
    form.pushButton_20.clicked.connect(start6)


def norma():
    form.spinBox_1.setProperty("value", sheet['E2'].value)
    form.spinBox_2.setProperty("value", sheet['E3'].value)
    form.spinBox_3.setProperty("value", sheet['E4'].value)
    form.spinBox_4.setProperty("value", sheet['E5'].value)
    form.spinBox_5.setProperty("value", sheet['E6'].value)
    form.spinBox_6.setProperty("value", sheet['E7'].value)
    form.spinBox_7.setProperty("value", sheet['E8'].value)
    form.spinBox_8.setProperty("value", sheet['E9'].value)
    form.spinBox_9.setProperty("value", sheet['E10'].value)
    form.spinBox_10.setProperty("value", sheet['E11'].value)
    form.spinBox_11.setProperty("value", sheet['E12'].value)
    form.spinBox_12.setProperty("value", sheet['E13'].value)


def Save():
    sheet['D2'].value = form.tableWidget.item(0, 0).text()
    sheet['D3'].value = form.tableWidget.item(2, 0).text()
    sheet['D4'].value = form.tableWidget.item(4, 0).text()
    sheet['D5'].value = form.tableWidget.item(6, 0).text()
    sheet['D6'].value = form.tableWidget.item(8, 0).text()
    sheet['D7'].value = form.tableWidget.item(10, 0).text()
    sheet['D8'].value = form.tableWidget.item(1, 0).text()
    sheet['D9'].value = form.tableWidget.item(3, 0).text()
    sheet['D10'].value = form.tableWidget.item(5, 0).text()
    sheet['D11'].value = form.tableWidget.item(7, 0).text()
    sheet['D12'].value = form.tableWidget.item(9, 0).text()
    sheet['D13'].value = form.tableWidget.item(11, 0).text()
    sheet['E2'].value = form.spinBox_1.text().replace(" hour", "")
    sheet['E3'].value = form.spinBox_2.text().replace(" hour", "")
    sheet['E4'].value = form.spinBox_3.text().replace(" hour", "")
    sheet['E5'].value = form.spinBox_4.text().replace(" hour", "")
    sheet['E6'].value = form.spinBox_5.text().replace(" hour", "")
    sheet['E7'].value = form.spinBox_6.text().replace(" hour", "")
    sheet['E8'].value = form.spinBox_7.text().replace(" hour", "")
    sheet['E9'].value = form.spinBox_8.text().replace(" hour", "")
    sheet['E10'].value = form.spinBox_9.text().replace(" hour", "")
    sheet['E10'].value = form.spinBox_9.text().replace(" hour", "")
    sheet['E11'].value = form.spinBox_10.text().replace(" hour", "")
    sheet['E12'].value = form.spinBox_11.text().replace(" hour", "")
    sheet['E13'].value = form.spinBox_12.text().replace(" hour", "")
    sheet["N4"].value = form.spinBox_13.text()
    sheet["N6"].value = form.spinBox_14.text()
    wb.save("lokofile.xlsx")


def initialTrains():
    form.tableWidget.item(10, 0).setText(str(wb['Sheet1']['D7'].value))
    form.tableWidget.item(8, 0).setText(str(wb['Sheet1']['D6'].value))
    form.tableWidget.item(6, 0).setText(str(wb['Sheet1']['D5'].value))
    form.tableWidget.item(4, 0).setText(str(wb['Sheet1']['D4'].value))
    form.tableWidget.item(2, 0).setText(str(wb['Sheet1']['D3'].value))
    form.tableWidget.item(0, 0).setText(str(wb['Sheet1']['D2'].value))
    form.tableWidget.item(11, 0).setText(str(wb['Sheet1']['D13'].value))
    form.tableWidget.item(9, 0).setText(str(wb['Sheet1']['D12'].value))
    form.tableWidget.item(7, 0).setText(str(wb['Sheet1']['D11'].value))
    form.tableWidget.item(5, 0).setText(str(wb['Sheet1']['D10'].value))
    form.tableWidget.item(3, 0).setText(str(wb['Sheet1']['D9'].value))
    form.tableWidget.item(1, 0).setText(str(wb['Sheet1']['D8'].value))


def progress_bar():
    lol2 = wb['Sheet1']['E2'].value
    lol3 = wb['Sheet1']['E3'].value
    lol4 = wb['Sheet1']['E4'].value
    lol5 = wb['Sheet1']['E5'].value
    lol6 = wb['Sheet1']['E6'].value
    lol7 = wb['Sheet1']['E7'].value
    lol8 = wb['Sheet1']['E8'].value
    lol9 = wb['Sheet1']['E9'].value
    lol10 = wb['Sheet1']['E10'].value
    lol11 = wb['Sheet1']['E11'].value
    lol12 = wb['Sheet1']['E12'].value
    lol13 = wb['Sheet1']['E13'].value
    if int(lol2) <= int(((int(sheet['R2'].value) + int(sheet['P2'].value)) * go // 3600)):
        form.progressBar.setProperty("value", 100)
        form.progressBar.setStyleSheet("QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                                       "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                                       "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol2) - int(((int(sheet['R2'].value) + int(sheet['P2'].value)) * go // 3600))) / int(
                lol2)) * 100) < 60:
            form.progressBar.setProperty("value",
                                         int((1 - int(int(lol2) - int(
                                             ((int(sheet['R2'].value) + int(sheet['P2'].value)) * go // 3600))) / int(
                                             lol2)) * 100))
        else:
            form.progressBar.setProperty("value",
                                         int((1 - int(int(lol2) - int(
                                             ((int(sheet['R2'].value) + int(sheet['P2'].value)) * go // 3600))) / int(
                                             lol2)) * 100))
            form.progressBar.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
    if int(lol3) <= int(((int(sheet['R3'].value) + int(sheet['P3'].value)) * go // 3600)):
        form.progressBar_2.setProperty("value", 100)
        form.progressBar_2.setStyleSheet("QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                                         "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                                         "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol3) - int(((int(sheet['R3'].value) + int(sheet['P3'].value)) * go // 3600))) / int(
                lol3)) * 100) < 60:
            form.progressBar_2.setProperty("value",
                                           int((1 - int(int(lol3) - int(
                                               ((int(sheet['R3'].value) + int(sheet['P3'].value)) * go // 3600))) / int(
                                               lol3)) * 100))
        else:
            form.progressBar_2.setProperty("value",
                                           int((1 - int(int(lol3) - int(
                                               ((int(sheet['R3'].value) + int(sheet['P3'].value)) * go // 3600))) / int(
                                               lol3)) * 100))
            form.progressBar_2.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
    if int(lol4) <= int(((int(sheet['R4'].value) + int(sheet['P4'].value)) * go // 3600)):
        form.progressBar_3.setProperty("value", 100)
        form.progressBar_3.setStyleSheet("QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                                         "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                                         "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol4) - int(((int(sheet['R4'].value) + int(sheet['P4'].value)) * go // 3600))) / int(
                lol4)) * 100) < 60:
            form.progressBar_3.setProperty("value",
                                           int((1 - int(int(lol4) - int(
                                               ((int(sheet['R4'].value) + int(sheet['P4'].value)) * go // 3600))) / int(
                                               lol4)) * 100))
        else:
            form.progressBar_3.setProperty("value",
                                           int((1 - int(int(lol4) - int(
                                               ((int(sheet['R4'].value) + int(sheet['P4'].value)) * go // 3600))) / int(
                                               lol4)) * 100))
            form.progressBar_3.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
    if int(lol5) <= int(((int(sheet['R5'].value) + int(sheet['P5'].value)) * go // 3600)):
        form.progressBar_4.setProperty("value", 100)
        form.progressBar_4.setStyleSheet("QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                                         "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                                         "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol5) - int(((int(sheet['R5'].value) + int(sheet['P5'].value)) * go // 3600))) / int(
                lol5)) * 100) < 60:
            form.progressBar_4.setProperty("value",
                                           int((1 - int(int(lol5) - int(
                                               ((int(sheet['R5'].value) + int(sheet['P5'].value)) * go // 3600))) / int(
                                               lol5)) * 100))
        else:
            form.progressBar_4.setProperty("value",
                                           int((1 - int(int(lol5) - int(
                                               ((int(sheet['R5'].value) + int(sheet['P5'].value)) * go // 3600))) / int(
                                               lol5)) * 100))
            form.progressBar_4.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
    if int(lol6) <= int(((int(sheet['R6'].value) + int(sheet['P6'].value)) * go // 3600)):
        form.progressBar_5.setProperty("value", 100)
        form.progressBar_5.setStyleSheet("QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                                         "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                                         "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol6) - int(((int(sheet['R6'].value) + int(sheet['P6'].value)) * go // 3600))) / int(
                lol6)) * 100) < 60:
            form.progressBar_5.setProperty("value",
                                           int((1 - int(int(lol6) - int(
                                               ((int(sheet['R6'].value) + int(sheet['P6'].value)) * go // 3600))) / int(
                                               lol6)) * 100))
        else:
            form.progressBar_5.setProperty("value",
                                           int((1 - int(int(lol6) - int(
                                               ((int(sheet['R6'].value) + int(sheet['P6'].value)) * go // 3600))) / int(
                                               lol6)) * 100))
            form.progressBar_5.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
    if int(lol7) <= int(((int(sheet['R7'].value) + int(sheet['P7'].value)) * go // 3600)):
        form.progressBar_6.setProperty("value", 100)
        form.progressBar_6.setStyleSheet("QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                                         "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                                         "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol7) - int(((int(sheet['R7'].value) + int(sheet['P7'].value)) * go // 3600))) / int(
                lol7)) * 100) < 60:
            form.progressBar_6.setProperty("value",
                                           int((1 - int(int(lol7) - int(
                                               ((int(sheet['R7'].value) + int(sheet['P7'].value)) * go // 3600))) / int(
                                               lol7)) * 100))
        else:
            form.progressBar_6.setProperty("value",
                                           int((1 - int(int(lol7) - int(
                                               ((int(sheet['R7'].value) + int(sheet['P7'].value)) * go // 3600))) / int(
                                               lol7)) * 100))
            form.progressBar_6.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
    if int(lol8) <= int(((int(sheet['Q2'].value) + int(sheet['O2'].value)) * go // 3600)):
        form.progressBar_7.setProperty("value", 100)
        form.progressBar_7.setStyleSheet("QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                                         "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                                         "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol8) - int(((int(sheet['Q2'].value) + int(sheet['O2'].value)) * go // 3600))) / int(
                lol8)) * 100) < 60:
            form.progressBar_7.setProperty("value",
                                           int((1 - int(int(lol8) - int(
                                               ((int(sheet['Q2'].value) + int(sheet['O2'].value)) * go // 3600))) / int(
                                               lol8)) * 100))
        else:
            form.progressBar_7.setProperty("value",
                                           int((1 - int(int(lol8) - int(
                                               ((int(sheet['Q2'].value) + int(sheet['O2'].value)) * go // 3600))) / int(
                                               lol8)) * 100))
            form.progressBar_7.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
    if int(lol9) <= int(((int(sheet['Q3'].value) + int(sheet['O3'].value)) * go // 3600)):
        form.progressBar_8.setProperty("value", 100)
        form.progressBar_8.setStyleSheet("QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                                         "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                                         "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol9) - int(((int(sheet['Q3'].value) + int(sheet['O3'].value)) * go // 3600))) / int(
                lol9)) * 100) < 60:
            form.progressBar_8.setProperty("value",
                                           int((1 - int(int(lol9) - int(
                                               ((int(sheet['Q3'].value) + int(sheet['O3'].value)) * go // 3600))) / int(
                                               lol9)) * 100))
        else:
            form.progressBar_8.setProperty("value",
                                           int((1 - int(int(lol9) - int(
                                               ((int(sheet['Q3'].value) + int(sheet['O3'].value)) * go // 3600))) / int(
                                               lol9)) * 100))
            form.progressBar_8.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
    if int(lol10) <= int(((int(sheet['Q4'].value) + int(sheet['O4'].value)) * go // 3600)):
        form.progressBar_9.setProperty("value", 100)
        form.progressBar_9.setStyleSheet("QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                                         "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                                         "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol10) - int(((int(sheet['Q4'].value) + int(sheet['O4'].value)) * go // 3600))) / int(
                lol10)) * 100) < 60:
            form.progressBar_9.setProperty("value",
                                           int((1 - int(int(lol10) - int(
                                               ((int(sheet['Q4'].value) + int(sheet['O4'].value)) * go // 3600))) / int(
                                               lol10)) * 100))
        else:
            form.progressBar_9.setProperty("value",
                                           int((1 - int(int(lol10) - int(
                                               ((int(sheet['Q4'].value) + int(sheet['O4'].value)) * go // 3600))) / int(
                                               lol10)) * 100))
            form.progressBar_9.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
    if int(lol11) <= int(((int(sheet['Q5'].value) + int(sheet['O5'].value)) * go // 3600)):
        form.progressBar_10.setProperty("value", 100)
        form.progressBar_10.setStyleSheet(
            "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
            "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
            "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol11) - int(((int(sheet['Q5'].value) + int(sheet['O5'].value)) * go // 3600))) / int(
                lol11)) * 100) < 60:
            form.progressBar_10.setProperty("value",
                                            int((1 - int(int(lol11) - int(((int(sheet['Q5'].value) + int(
                                                sheet['O5'].value)) * go // 3600))) / int(lol11)) * 100))
        else:
            form.progressBar_10.setProperty("value",
                                            int((1 - int(int(lol11) - int(((int(sheet['Q5'].value) + int(
                                                sheet['O5'].value)) * go // 3600))) / int(lol11)) * 100))
            form.progressBar_10.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
    if int(lol12) <= int(((int(sheet['Q6'].value) + int(sheet['O6'].value)) * go // 3600)):
        form.progressBar_11.setProperty("value", 100)
        form.progressBar_11.setStyleSheet(
            "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
            "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
            "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol12) - int(((int(sheet['Q6'].value) + int(sheet['O6'].value)) * go // 3600))) / int(
                lol12)) * 100) < 60:
            form.progressBar_11.setProperty("value",
                                            int((1 - int(int(lol12) - int(((int(sheet['Q6'].value) + int(
                                                sheet['O6'].value)) * go // 3600))) / int(lol12)) * 100))
        else:
            form.progressBar_11.setProperty("value",
                                            int((1 - int(int(lol12) - int(((int(sheet['Q6'].value) + int(
                                                sheet['O6'].value)) * go // 3600))) / int(lol12)) * 100))
            form.progressBar_11.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
    if int(lol13) <= int(((int(sheet['Q7'].value) + int(sheet['O7'].value)) * go // 3600)):
        form.progressBar_12.setProperty("value", 100)
        form.progressBar_12.setStyleSheet(
            "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
            "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
            "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
    else:
        if int((1 - int(int(lol13) - int(((int(sheet['Q7'].value) + int(sheet['O7'].value)) * go // 3600))) / int(
                lol13)) * 100) < 60:
            form.progressBar_12.setProperty("value",
                                            int((1 - int(int(lol13) - int(((int(sheet['Q7'].value) + int(
                                                sheet['O7'].value)) * go // 3600))) / int(lol13)) * 100))
        else:
            form.progressBar_12.setProperty("value",
                                            int((1 - int(int(lol13) - int(((int(sheet['Q7'].value) + int(
                                                sheet['O7'].value)) * go // 3600))) / int(lol13)) * 100))
            form.progressBar_12.setStyleSheet(
                "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")


def progress_bar_update():
    while True:
        global lol2, lol3, lol4, lol5, lol6, lol7, lol8, lol9, lol10, lol11, lol12, lol13
        time.sleep(300)
        lol2 = wb['Sheet1']['E2'].value
        lol3 = wb['Sheet1']['E3'].value
        lol4 = wb['Sheet1']['E4'].value
        lol5 = wb['Sheet1']['E5'].value
        lol6 = wb['Sheet1']['E6'].value
        lol7 = wb['Sheet1']['E7'].value
        lol8 = wb['Sheet1']['E8'].value
        lol9 = wb['Sheet1']['E9'].value
        lol10 = wb['Sheet1']['E10'].value
        lol11 = wb['Sheet1']['E11'].value
        lol12 = wb['Sheet1']['E12'].value
        lol13 = wb['Sheet1']['E13'].value
        if int(lol2) <= int(((int(sheet['R2'].value) + int(sheet['P2'].value)) * go // 3600)):
            form.progressBar.setProperty("value", 100)
            form.progressBar.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol2) - int(((int(sheet['R2'].value) + int(sheet['P2'].value)) * go // 3600))) / int(
                    lol2)) * 100) < 60:
                form.progressBar.setProperty("value",
                                             int((1 - int(int(lol2) - int(((int(sheet['R2'].value) + int(
                                                 sheet['P2'].value)) * go // 3600))) / int(lol2)) * 100))
                form.progressBar.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                               "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                               "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar.setProperty("value",
                                             int((1 - int(int(lol2) - int(((int(sheet['R2'].value) + int(
                                                 sheet['P2'].value)) * go // 3600))) / int(lol2)) * 100))
                form.progressBar.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
        time.sleep(1)
        if int(lol3) <= int(((int(sheet['R3'].value) + int(sheet['P3'].value)) * go // 3600)):
            form.progressBar_2.setProperty("value", 100)
            form.progressBar_2.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol3) - int(((int(sheet['R3'].value) + int(sheet['P3'].value)) * go // 3600))) / int(
                    lol3)) * 100) < 60:
                form.progressBar_2.setProperty("value",
                                               int((1 - int(int(lol3) - int(((int(sheet['R3'].value) + int(
                                                   sheet['P3'].value)) * go // 3600))) / int(lol3)) * 100))
                form.progressBar_2.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                                 "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                                 "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar_2.setProperty("value",
                                               int((1 - int(int(lol3) - int(((int(sheet['R3'].value) + int(
                                                   sheet['P3'].value)) * go // 3600))) / int(lol3)) * 100))
                form.progressBar_2.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
        time.sleep(1)
        if int(lol4) <= int(((int(sheet['R4'].value) + int(sheet['P4'].value)) * go // 3600)):
            form.progressBar_3.setProperty("value", 100)
            form.progressBar_3.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol4) - int(((int(sheet['R4'].value) + int(sheet['P4'].value)) * go // 3600))) / int(
                    lol4)) * 100) < 60:
                form.progressBar_3.setProperty("value",
                                               int((1 - int(int(lol4) - int(((int(sheet['R4'].value) + int(
                                                   sheet['P4'].value)) * go // 3600))) / int(lol4)) * 100))
                form.progressBar_3.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                                 "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                                 "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar_3.setProperty("value",
                                               int((1 - int(int(lol4) - int(((int(sheet['R4'].value) + int(
                                                   sheet['P4'].value)) * go // 3600))) / int(lol4)) * 100))
                form.progressBar_3.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
        time.sleep(1)
        if int(lol5) <= int(((int(sheet['R5'].value) + int(sheet['P5'].value)) * go // 3600)):
            form.progressBar_4.setProperty("value", 100)
            form.progressBar_4.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol5) - int(((int(sheet['R5'].value) + int(sheet['P5'].value)) * go // 3600))) / int(
                    lol5)) * 100) < 60:
                form.progressBar_4.setProperty("value",
                                               int((1 - int(int(lol5) - int(((int(sheet['R5'].value) + int(
                                                   sheet['P5'].value)) * go // 3600))) / int(lol5)) * 100))
                form.progressBar_4.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                                 "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                                 "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar_4.setProperty("value",
                                               int((1 - int(int(lol5) - int(((int(sheet['R5'].value) + int(
                                                   sheet['P5'].value)) * go // 3600))) / int(lol5)) * 100))
                form.progressBar_4.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
        time.sleep(1)
        if int(lol6) <= int(((int(sheet['R6'].value) + int(sheet['P6'].value)) * go // 3600)):
            form.progressBar_5.setProperty("value", 100)
            form.progressBar_5.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol6) - int(((int(sheet['R6'].value) + int(sheet['P6'].value)) * go // 3600))) / int(
                    lol6)) * 100) < 60:
                form.progressBar_5.setProperty("value",
                                               int((1 - int(int(lol6) - int(((int(sheet['R6'].value) + int(
                                                   sheet['P6'].value)) * go // 3600))) / int(lol6)) * 100))
                form.progressBar_5.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                                 "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                                 "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar_5.setProperty("value",
                                               int((1 - int(int(lol6) - int(((int(sheet['R6'].value) + int(
                                                   sheet['P6'].value)) * go // 3600))) / int(lol6)) * 100))
                form.progressBar_5.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
        time.sleep(1)
        if int(lol7) <= int(((int(sheet['R7'].value) + int(sheet['P7'].value)) * go // 3600)):
            form.progressBar_6.setProperty("value", 100)
            form.progressBar_6.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol7) - int(((int(sheet['R7'].value) + int(sheet['P7'].value)) * go // 3600))) / int(
                    lol7)) * 100) < 60:
                form.progressBar_6.setProperty("value",
                                               int((1 - int(int(lol7) - int(((int(sheet['R7'].value) + int(
                                                   sheet['P7'].value)) * go // 3600))) / int(lol7)) * 100))
                form.progressBar_6.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                                 "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                                 "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar_6.setProperty("value",
                                               int((1 - int(int(lol7) - int(((int(sheet['R7'].value) + int(
                                                   sheet['P7'].value)) * go // 3600))) / int(lol7)) * 100))
                form.progressBar_6.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
        time.sleep(1)
        if int(lol8) <= int(((int(sheet['Q2'].value) + int(sheet['O2'].value)) * go // 3600)):
            form.progressBar_7.setProperty("value", 100)
            form.progressBar_7.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol8) - int(((int(sheet['Q2'].value) + int(sheet['O2'].value)) * go // 3600))) / int(
                    lol8)) * 100) < 60:
                form.progressBar_7.setProperty("value",
                                               int((1 - int(int(lol8) - int(((int(sheet['Q2'].value) + int(
                                                   sheet['O2'].value)) * go // 3600))) / int(lol8)) * 100))
                form.progressBar_7.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                                 "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                                 "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar_7.setProperty("value",
                                               int((1 - int(int(lol8) - int(((int(sheet['Q2'].value) + int(
                                                   sheet['O2'].value)) * go // 3600))) / int(lol8)) * 100))
                form.progressBar_7.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
        time.sleep(1)
        if int(lol9) <= int(((int(sheet['Q3'].value) + int(sheet['O3'].value)) * go // 3600)):
            form.progressBar_8.setProperty("value", 100)
            form.progressBar_8.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol9) - int(((int(sheet['Q3'].value) + int(sheet['O3'].value)) * go // 3600))) / int(
                    lol9)) * 100) < 60:
                form.progressBar_8.setProperty("value",
                                               int((1 - int(int(lol9) - int(((int(sheet['Q3'].value) + int(
                                                   sheet['O3'].value)) * go // 3600))) / int(lol9)) * 100))
                form.progressBar_8.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                                 "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                                 "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar_8.setProperty("value",
                                               int((1 - int(int(lol9) - int(((int(sheet['Q3'].value) + int(
                                                   sheet['O3'].value)) * go // 3600))) / int(lol9)) * 100))
                form.progressBar_8.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
        time.sleep(1)
        if int(lol10) <= int(((int(sheet['Q4'].value) + int(sheet['O4'].value)) * go // 3600)):
            form.progressBar_9.setProperty("value", 100)
            form.progressBar_9.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol10) - int(((int(sheet['Q4'].value) + int(sheet['O4'].value)) * go // 3600))) / int(
                    lol10)) * 100) < 60:
                form.progressBar_9.setProperty("value",
                                               int((1 - int(int(lol10) - int(((int(sheet['Q4'].value) + int(
                                                   sheet['O4'].value)) * go // 3600))) / int(lol10)) * 100))
                form.progressBar_9.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                                 "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                                 "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar_9.setProperty("value",
                                               int((1 - int(int(lol10) - int(((int(sheet['Q4'].value) + int(
                                                   sheet['O4'].value)) * go // 3600))) / int(lol10)) * 100))
                form.progressBar_9.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
        time.sleep(1)
        if int(lol11) <= int(((int(sheet['Q5'].value) + int(sheet['O5'].value)) * go // 3600)):
            form.progressBar_10.setProperty("value", 100)
            form.progressBar_10.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol11) - int(((int(sheet['Q5'].value) + int(sheet['O5'].value)) * go // 3600))) / int(
                    lol11)) * 100) < 60:
                form.progressBar_10.setProperty("value",
                                                int((1 - int(int(lol11) - int(((int(sheet['Q5'].value) + int(
                                                    sheet['O5'].value)) * go // 3600))) / int(lol11)) * 100))
                form.progressBar_10.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                                  "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                                  "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar_10.setProperty("value",
                                                int((1 - int(int(lol11) - int(((int(sheet['Q5'].value) + int(
                                                    sheet['O5'].value)) * go // 3600))) / int(lol11)) * 100))
                form.progressBar_10.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
        time.sleep(1)
        if int(lol12) <= int(((int(sheet['Q6'].value) + int(sheet['O6'].value)) * go // 3600)):
            form.progressBar_11.setProperty("value", 100)
            form.progressBar_11.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol12) - int(((int(sheet['Q6'].value) + int(sheet['O6'].value)) * go // 3600))) / int(
                    lol12)) * 100) < 60:
                form.progressBar_11.setProperty("value",
                                                int((1 - int(int(lol12) - int(((int(sheet['Q6'].value) + int(
                                                    sheet['O6'].value)) * go // 3600))) / int(lol12)) * 100))
                form.progressBar_11.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                                  "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                                  "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar_11.setProperty("value",
                                                int((1 - int(int(lol12) - int(((int(sheet['Q6'].value) + int(
                                                    sheet['O6'].value)) * go // 3600))) / int(lol12)) * 100))
                form.progressBar_11.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")
        time.sleep(1)
        if int(lol13) <= int(((int(sheet['Q7'].value) + int(sheet['O7'].value)) * go // 3600)):
            form.progressBar_12.setProperty("value", 100)
            form.progressBar_12.setStyleSheet(
                "QProgressBar::chunk {background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0,"
                "stop: 0 #FF0350,stop: 0.4999 #FF0020,stop: 0.5 #FF0019,stop: 1 #FF0000 );border"
                "-bottom-right-radius: 1px;border-bottom-left-radius: 1px;border: 1px solid black;}")
        else:
            if int((1 - int(int(lol13) - int(((int(sheet['Q7'].value) + int(sheet['O7'].value)) * go // 3600))) / int(
                    lol13)) * 100) < 60:
                form.progressBar_12.setProperty("value",
                                                int((1 - int(int(lol13) - int(((int(sheet['Q7'].value) + int(
                                                    sheet['O7'].value)) * go // 3600))) / int(lol13)) * 100))
                form.progressBar_12.setStyleSheet("QProgressBar::chunk {background-color: qlineargradient(spread:pad, "
                                                  "x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), "
                                                  "stop:1 rgba(221, 221, 221, 255));border: 1px solid}")
            else:
                form.progressBar_12.setProperty("value",
                                                int((1 - int(int(lol13) - int(((int(sheet['Q7'].value) + int(
                                                    sheet['O7'].value)) * go // 3600))) / int(lol13)) * 100))
                form.progressBar_12.setStyleSheet(
                    "QProgressBar::chunk {background-color: rgb(250, 232, 25); border: 1px solid;}")


def minuts():
    global min
    min = int(sheet["N4"].value) * 60


def pool():
    form.spinBox_13.setProperty("value", sheet["N4"].value)
    form.spinBox_14.setProperty("value", sheet["N6"].value)


def seconds():
    global sec, go
    sec = int(sheet["N6"].value)
    go = min + sec


def interval():
    minuts()
    seconds()
    time_now = f"{((int(sheet['R2'].value) + int(sheet['P2'].value)) * go // 3600)}h {((int(sheet['R2'].value) + int(sheet['P2'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(0, 2).setText(time_now)
    time_now1 = f"{((int(sheet['Q2'].value) + int(sheet['O2'].value)) * go // 3600)}h {((int(sheet['Q2'].value) + int(sheet['O2'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(1, 2).setText(time_now1)
    time_now2 = f"{((int(sheet['R3'].value) + int(sheet['P3'].value)) * go // 3600)}h {((int(sheet['R3'].value) + int(sheet['P3'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(2, 2).setText(time_now2)
    time_now3 = f"{((int(sheet['Q3'].value) + int(sheet['O3'].value)) * go // 3600)}h {((int(sheet['Q3'].value) + int(sheet['O3'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(3, 2).setText(time_now3)
    time_now4 = f"{((int(sheet['R4'].value) + int(sheet['P4'].value)) * go // 3600)}h {((int(sheet['R4'].value) + int(sheet['P4'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(4, 2).setText(time_now4)
    time_now5 = f"{((int(sheet['Q4'].value) + int(sheet['O4'].value)) * go // 3600)}h {((int(sheet['Q4'].value) + int(sheet['O4'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(5, 2).setText(time_now5)
    time_now6 = f"{((int(sheet['R5'].value) + int(sheet['P5'].value)) * go // 3600)}h {((int(sheet['R5'].value) + int(sheet['P5'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(6, 2).setText(time_now6)
    time_now7 = f"{((int(sheet['Q5'].value) + int(sheet['O5'].value)) * go // 3600)}h {((int(sheet['Q5'].value) + int(sheet['O5'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(7, 2).setText(time_now7)
    time_now8 = f"{((int(sheet['R6'].value) + int(sheet['P6'].value)) * go // 3600)}h {((int(sheet['R6'].value) + int(sheet['P6'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(8, 2).setText(time_now8)
    time_now9 = f"{((int(sheet['Q6'].value) + int(sheet['O6'].value)) * go // 3600)}h {((int(sheet['Q6'].value) + int(sheet['O6'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(9, 2).setText(time_now9)
    time_now10 = f"{((int(sheet['R7'].value) + int(sheet['P7'].value)) * go // 3600)}h {((int(sheet['R7'].value) + int(sheet['P7'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(10, 2).setText(time_now10)
    time_now11 = f"{((int(sheet['Q7'].value) + int(sheet['O7'].value)) * go // 3600)}h {((int(sheet['Q7'].value) + int(sheet['O7'].value)) * go % 3600 // 60)}m"
    form.tableWidget.item(11, 2).setText(time_now11)


def interval_now():
    while True:
        minuts()
        seconds()
        time_now = f"{((int(sheet['R2'].value) + int(sheet['P2'].value)) * go // 3600)}h {((int(sheet['R2'].value) + int(sheet['P2'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(0, 2).setText(time_now)
        time_now1 = f"{((int(sheet['Q2'].value) + int(sheet['O2'].value)) * go // 3600)}h {((int(sheet['Q2'].value) + int(sheet['O2'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(1, 2).setText(time_now1)
        time_now2 = f"{((int(sheet['R3'].value) + int(sheet['P3'].value)) * go // 3600)}h {((int(sheet['R3'].value) + int(sheet['P3'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(2, 2).setText(time_now2)
        time_now3 = f"{((int(sheet['Q3'].value) + int(sheet['O3'].value)) * go // 3600)}h {((int(sheet['Q3'].value) + int(sheet['O3'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(3, 2).setText(time_now3)
        time_now4 = f"{((int(sheet['R4'].value) + int(sheet['P4'].value)) * go // 3600)}h {((int(sheet['R4'].value) + int(sheet['P4'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(4, 2).setText(time_now4)
        time_now5 = f"{((int(sheet['Q4'].value) + int(sheet['O4'].value)) * go // 3600)}h {((int(sheet['Q4'].value) + int(sheet['O4'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(5, 2).setText(time_now5)
        time_now6 = f"{((int(sheet['R5'].value) + int(sheet['P5'].value)) * go // 3600)}h {((int(sheet['R5'].value) + int(sheet['P5'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(6, 2).setText(time_now6)
        time_now7 = f"{((int(sheet['Q5'].value) + int(sheet['O5'].value)) * go // 3600)}h {((int(sheet['Q5'].value) + int(sheet['O5'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(7, 2).setText(time_now7)
        time_now8 = f"{((int(sheet['R6'].value) + int(sheet['P6'].value)) * go // 3600)}h {((int(sheet['R6'].value) + int(sheet['P6'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(8, 2).setText(time_now8)
        time_now9 = f"{((int(sheet['Q6'].value) + int(sheet['O6'].value)) * go // 3600)}h {((int(sheet['Q6'].value) + int(sheet['O6'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(9, 2).setText(time_now9)
        time_now10 = f"{((int(sheet['R7'].value) + int(sheet['P7'].value)) * go // 3600)}h {((int(sheet['R7'].value) + int(sheet['P7'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(10, 2).setText(time_now10)
        time_now11 = f"{((int(sheet['Q7'].value) + int(sheet['O7'].value)) * go // 3600)}h {((int(sheet['Q7'].value) + int(sheet['O7'].value)) * go % 3600 // 60)}m"
        form.tableWidget.item(11, 2).setText(time_now11)
        time.sleep(60)


def today_z():
    if str(sheet["N2"].value[0:10]) <= str(date.today()):
        sheet["N2"].value = str(date.today() + timedelta(days=1))
        sheet["P2"].value = str(int(sheet["R2"].value) + int(sheet["P2"].value))
        sheet["P3"].value = str(int(sheet["R3"].value) + int(sheet["P3"].value))
        sheet["P4"].value = str(int(sheet["R4"].value) + int(sheet["P4"].value))
        sheet["P5"].value = str(int(sheet["R5"].value) + int(sheet["P5"].value))
        sheet["P6"].value = str(int(sheet["R6"].value) + int(sheet["P6"].value))
        sheet["P7"].value = str(int(sheet["R7"].value) + int(sheet["P7"].value))
        sheet["O2"].value = str(int(sheet["Q2"].value) + int(sheet["O2"].value))
        sheet["O3"].value = str(int(sheet["Q3"].value) + int(sheet["O3"].value))
        sheet["O4"].value = str(int(sheet["Q4"].value) + int(sheet["O4"].value))
        sheet["O5"].value = str(int(sheet["Q5"].value) + int(sheet["O5"].value))
        sheet["O6"].value = str(int(sheet["Q6"].value) + int(sheet["O6"].value))
        sheet["O7"].value = str(int(sheet["Q7"].value) + int(sheet["O7"].value))
        sheet["R2"].value = "0"
        sheet["R3"].value = "0"
        sheet["R4"].value = "0"
        sheet["R5"].value = "0"
        sheet["R6"].value = "0"
        sheet["R7"].value = "0"
        sheet["Q2"].value = "0"
        sheet["Q3"].value = "0"
        sheet["Q4"].value = "0"
        sheet["Q5"].value = "0"
        sheet["Q6"].value = "0"
        sheet["Q7"].value = "0"
        sheet["M2"].value = "0"
        sheet["M3"].value = "0"
        sheet["M4"].value = "0"
        sheet["M5"].value = "0"
        sheet["M6"].value = "0"
        sheet["M7"].value = "0"
        sheet["L2"].value = "0"
        sheet["L3"].value = "0"
        sheet["L4"].value = "0"
        sheet["L5"].value = "0"
        sheet["L6"].value = "0"
        sheet["L7"].value = "0"
        wb.save("lokofile.xlsx")


def puff():
    if pause_flag_1 == 1:
        form.pushButton_15.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189);color: red;border:1px solid black;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255)}")
        form.pushButton_3.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240); color: green;border:1px solid black;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.tableWidget.item(1, 0).setBackground(QtGui.QColor(255, 255, 255))
        form.tableWidget.item(0, 0).setBackground(QtGui.QColor(255, 161, 148))
        form.spinBox_7.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_1.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(1, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(0, i).setBackground(QtGui.QColor(255, 161, 148))
    if start_flag_1 == 1:
        form.pushButton_3.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189);color: red;border:1px solid black;border-radius: 4px}"
            "QPushButton:hover{background-color:  rgb(225, 252, 255); }")
        form.pushButton_15.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240);color: green;border:1px solid black;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.spinBox_1.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_7.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(0, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(1, i).setBackground(QtGui.QColor(255, 161, 148))
    if pause_flag_2 == 1:
        form.pushButton_16.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189);color: red;border:1px solid black;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255)}")
        form.pushButton_4.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240); color: green;border:1px solid black;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.spinBox_8.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_2.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(3, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(2, i).setBackground(QtGui.QColor(255, 161, 148))
    if start_flag_2 == 1:
        form.pushButton_4.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189); color: red;border:1px solid black;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.pushButton_16.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240);color: green;border:1px solid black;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.spinBox_2.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_8.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(2, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(3, i).setBackground(QtGui.QColor(255, 161, 148))
    if pause_flag_3 == 1:
        form.pushButton_17.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189);color: red; border: 1px solid black;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255)}")
        form.pushButton_11.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240); border: 1px solid black;color: green;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.spinBox_9.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_3.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(5, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(4, i).setBackground(QtGui.QColor(255, 161, 148))
    if start_flag_3 == 1:
        form.pushButton_11.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189); border: 1px solid black;color: red;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.pushButton_17.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240); border: 1px solid black;color: green;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.spinBox_3.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_9.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(4, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(5, i).setBackground(QtGui.QColor(255, 161, 148))
    if pause_flag_4 == 1:
        form.pushButton_18.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189);color: red; border: 1px solid black;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255)}")
        form.pushButton_12.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240); border: 1px solid black;color: green;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.spinBox_10.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_4.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(7, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(6, i).setBackground(QtGui.QColor(255, 161, 148))
    if start_flag_4 == 1:
        form.pushButton_12.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189); border: 1px solid black;color: red;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.pushButton_18.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240); border: 1px solid black;color: green;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.spinBox_4.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_10.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(6, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(7, i).setBackground(QtGui.QColor(255, 161, 148))
    if pause_flag_5 == 1:
        form.pushButton_19.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189);color: red; border: 1px solid black;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255)}")
        form.pushButton_13.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240); border: 1px solid black;color: green;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.spinBox_11.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_5.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(9, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(8, i).setBackground(QtGui.QColor(255, 161, 148))
    if start_flag_5 == 1:
        form.pushButton_13.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189); border: 1px solid black;color: red;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.pushButton_19.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240); border: 1px solid black;color: green;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.spinBox_5.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_11.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(8, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(9, i).setBackground(QtGui.QColor(255, 161, 148))
    if pause_flag_6 == 1:
        form.pushButton_20.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189);color: red; border: 1px solid black;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255)}")
        form.pushButton_14.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240); border: 1px solid black;color: green;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.spinBox_12.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_6.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(11, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(10, i).setBackground(QtGui.QColor(255, 161, 148))
    if start_flag_6 == 1:
        form.pushButton_14.setStyleSheet(
            "QPushButton{background-color: rgb(202, 201, 189);border: 1px solid black;color: red;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.pushButton_20.setStyleSheet(
            "QPushButton{background-color: rgb(240, 240, 240); border: 1px solid black;color: green;border-radius: 4px} "
            "QPushButton:hover{background-color:  rgb(225, 252, 255);}")
        form.spinBox_6.setStyleSheet("QSpinBox {background-color: rgb(85,255,127)}")
        form.spinBox_12.setStyleSheet("QSpinBox {background-color: rgb(255,161,148)}")
        for i in range(4):
            form.tableWidget.item(10, i).setBackground(QtGui.QColor(85, 255, 127))
            form.tableWidget.item(11, i).setBackground(QtGui.QColor(255, 161, 148))


def run():
    global pool
    i = 0
    road = today.date().strftime("%d.%m.%Y")
    try:
        open(fr'\\PULT1\RailControl\log\{road}.txt')
        pool = 1
    except:
        pool = 0
        time.sleep(5)
        return run()
    with open(fr'\\PULT1\RailControl\log\{road}.txt') as t:
        while True:
            time.sleep(3)
            for line in t:
                if "total" in line:
                    i += 1
            if i != str(
                    int(sheet["R2"].value) + int(sheet["Q2"].value) + int(sheet["M2"].value) + int(sheet["L2"].value)):
                if pause_flag_1 == 0 and start_flag_1 == 1:
                    sheet["R2"].value = str(
                        (int(i) - int(sheet["Q2"].value) - int(sheet["M2"].value) - int(sheet["L2"].value)))
                    wb.save("lokofile.xlsx")
                else:
                    sheet["Q2"].value = str(
                        int(i) - int(sheet["R2"].value) - int(sheet["L2"].value) - int(sheet["M2"].value))
                    wb.save("lokofile.xlsx")
            time.sleep(1)


def run1():
    global pool
    o = 0
    road = today.date().strftime("%d.%m.%Y")
    try:
        open(fr'\\PULT2\RailControl\log\{road}.txt')
        pool = 1
    except:
        pool = 0
        time.sleep(5)
        return run1()
    with open(fr'\\PULT2\RailControl\log\{road}.txt') as t1:
        while True:
            time.sleep(5)
            for line1 in t1:
                if "total" in line1:
                    o += 1
            if o != str(
                    int(sheet["R3"].value) + int(sheet["Q3"].value) + int(sheet["M3"].value) + int(sheet["L3"].value)):
                if pause_flag_2 == 0 and start_flag_2 == 1:
                    sheet["R3"].value = str(
                        int(o) - int(sheet["Q3"].value) - int(sheet["M3"].value) - int(sheet["L3"].value))
                    wb.save("lokofile.xlsx")
                else:
                    sheet["Q3"].value = str(
                        int(o) - int(sheet["R3"].value) - int(sheet["L3"].value) - int(sheet["M3"].value))
                    wb.save("lokofile.xlsx")
            time.sleep(1)


def run2():
    global pool
    p = 0
    road = today.date().strftime("%d.%m.%Y")
    try:
        open(fr'\\PULT3\RailControl\log\{road}.txt')
        pool = 1
    except:
        pool = 0
        time.sleep(5)
        return run2()
    with open(fr'\\PULT3\RailControl\log\{road}.txt') as t2:
        while True:
            time.sleep(7)
            for line2 in t2:
                if "total" in line2:
                    p += 1
            if p != str(
                    int(sheet["R4"].value) + int(sheet["Q4"].value) + int(sheet["M4"].value) + int(sheet["L4"].value)):
                if pause_flag_3 == 0 and start_flag_3 == 1:
                    sheet["R4"].value = str(
                        int(p) - int(sheet["Q4"].value) - int(sheet["M4"].value) - int(sheet["L4"].value))
                    wb.save("lokofile.xlsx")
                else:
                    sheet["Q4"].value = str(
                        int(p) - int(sheet["R4"].value) - int(sheet["L4"].value) - int(sheet["M4"].value))
                    wb.save("lokofile.xlsx")
            time.sleep(1)


def run3():
    global pool
    j = 0
    road = today.date().strftime("%d.%m.%Y")
    try:
        open(fr'\\PULT4\RailControl\log\{road}.txt')
        pool = 1
    except:
        pool = 0
        time.sleep(5)
        return run3()
    with open(fr'\\PULT4\RailControl\log\{road}.txt') as t3:
        while True:
            time.sleep(9)
            for line3 in t3:
                if "total" in line3:
                    j += 1
            if j != str(
                    int(sheet["R5"].value) + int(sheet["Q5"].value) + int(sheet["M5"].value) + int(sheet["L5"].value)):
                if pause_flag_4 == 0 and start_flag_4 == 1:
                    sheet["R5"].value = str(
                        int(j) - int(sheet["Q5"].value) - int(sheet["M5"].value) - int(sheet["L5"].value))
                    wb.save("lokofile.xlsx")
                else:
                    sheet["Q5"].value = str(
                        int(j) - int(sheet["R5"].value) - int(sheet["L5"].value) - int(sheet["M5"].value))
                    wb.save("lokofile.xlsx")
            time.sleep(1)


def run4():
    global pool
    k = 0
    road = today.date().strftime("%d.%m.%Y")
    try:
        open(fr'\\PULT5\RailControl\log\{road}.txt')
        pool = 1
    except:
        pool = 0
        time.sleep(5)
        return run4()
    with open(fr'\\PULT5\RailControl\log\{road}.txt') as t4:
        while True:
            time.sleep(11)
            for line4 in t4:
                if "total" in line4:
                    k += 1
            if k != str(
                    int(sheet["R6"].value) + int(sheet["Q6"].value) + int(sheet["M6"].value) + int(sheet["L6"].value)):
                if pause_flag_5 == 0 and start_flag_5 == 1:
                    sheet["R6"].value = str(
                        int(k) - int(sheet["Q6"].value) - int(sheet["M6"].value) - int(sheet["L6"].value))
                    wb.save("lokofile.xlsx")
                else:
                    sheet["Q6"].value = str(
                        int(k) - int(sheet["R6"].value) - int(sheet["L6"].value) - int(sheet["M6"].value))
                    wb.save("lokofile.xlsx")
            time.sleep(1)


def run5():
    l = 0
    global pool
    road = today.date().strftime("%d.%m.%Y")
    try:
        open(fr'\\PULT6\RailControl\log\{road}.txt')
        pool = 1
    except:
        pool = 0
        time.sleep(5)
        return run5()
    with open(fr'\\PULT6\RailControl\log\{road}.txt') as t5:
        while True:
            time.sleep(13)
            for line5 in t5:
                if "total" in line5:
                    l += 1
            if l != str(
                    int(sheet["R7"].value) + int(sheet["Q7"].value) + int(sheet["M7"].value) + int(sheet["L7"].value)):
                if pause_flag_6 == 0 and start_flag_6 == 1:
                    sheet["R7"].value = str(
                        int(l) - int(sheet["Q7"].value) - int(sheet["M7"].value) - int(sheet["L7"].value))
                    wb.save("lokofile.xlsx")
                else:
                    sheet["Q7"].value = str(
                        int(l) - int(sheet["R7"].value) - int(sheet["L7"].value) - int(sheet["M7"].value))
                    wb.save("lokofile.xlsx")
            time.sleep(1)


def connect():
    if pool == 1:
        form.checkBox.setStyleSheet("QCheckBox::indicator {background-color: qlineargradient"
                                    "(spread:pad, x1:0.447, y1:0.341, x2:1, y2:0, stop:0 rgba(66, 200, 77, 255), stop:1 rgba(221, 221, 221, 255));}")
    else:
        form.checkBox.setStyleSheet("QCheckBox::indicator {background: qlineargradient"
                                    "(spread:pad, x1:0, y1:0.29, x2:1, y2:0, stop:0.253731 rgba(164, 0, 0, 255), stop:1 rgba(255, 0, 40, 255));}")


def button_clicked12():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D13'].value + " ?"))

    def fact():
        Reset12()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


def button_clicked11():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D12'].value + " ?"))

    def fact():
        Reset11()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


def button_clicked10():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D11'].value + " ?"))

    def fact():
        Reset10()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


def button_clicked9():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D10'].value + " ?"))

    def fact():
        Reset9()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


def button_clicked8():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D9'].value + " ?"))

    def fact():
        Reset8()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


def button_clicked7():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D8'].value + " ?"))

    def fact():
        Reset7()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


def button_clicked6():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D7'].value + " ?"))

    def fact():
        Reset6()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


def button_clicked5():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D6'].value + " ?"))

    def fact():
        Reset5()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


def button_clicked4():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D5'].value + " ?"))

    def fact():
        Reset4()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


def button_clicked3():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D4'].value + " ?"))

    def fact():
        Reset3()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


def button_clicked2():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D3'].value + " ?"))

    def fact():
        Reset2()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


def button_clicked1():
    Form, App = uic.loadUiType("dialog.ui")
    app = QDialog()
    form = Form()
    form.setupUi(app)
    app.show()
    form.label_2.setText(str(wb['Sheet1']['D2'].value + " ?"))

    def fact():
        Reset1()
        app.close()

    def close():
        app.close()

    form.pushButton.clicked.connect(fact)
    form.pushButton_2.clicked.connect(close)
    app.exec()


today_z()
pool()
button()
button_pause()
button_start()
norma()
interval()
initialTrains()
puff()
minuts()
seconds()
progress_bar()
form.pushButton.clicked.connect(Save)
thr1 = threading.Thread(target=run, daemon=True).start()
thr2 = threading.Thread(target=run1, daemon=True).start()
thr3 = threading.Thread(target=run2, daemon=True).start()
thr4 = threading.Thread(target=run3, daemon=True).start()
thr5 = threading.Thread(target=run4, daemon=True).start()
thr6 = threading.Thread(target=run5, daemon=True).start()
thr7 = threading.Thread(target=progress_bar_update, daemon=True).start()
thr8 = threading.Thread(target=interval_now, daemon=True).start()
timer = QTimer(interval=30000, timeout=connect)
timer.start()
time.sleep(2)
app.exec()
time.sleep(5)
if thr1 == True:
    daemon = False
