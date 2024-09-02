import importlib
import PyQt5.QtWidgets
from openpyxl.styles import PatternFill
from PyQt5 import QtCore, QtGui, QtWidgets
from front import Ui_MainWindow
from PyQt5 import QtCore, QtGui, QtWidgets, QtTest
from PyQt5.QtWidgets import QMessageBox, QFileDialog
import re
from openpyxl.utils import get_column_letter
import time
import csv
import shutil
import datetime
from PyQt5.QtGui import QTextCursor
import PyQt5.QtWidgets
from PyQt5.QtWidgets import QLineEdit
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5 import QtCore, QtGui, QtWidgets, QtTest
import sys
# import os
# import re
# # from pyexcel.cookbook import merge_all_to_a_book
# import glob
# from PyQt5.QtWidgets import QMessageBox, QFileDialog
# # import win32com.client as win32
# from PyQt5.QtGui import QTextCursor
# # import numpy as np
# # from tabulate import tabulate
# import datetime
# import winsound
# import pandas as pd
# from openpyxl import workbook, load_workbook
# import openpyxl
# import traceback
# from openpyxl.worksheet.datavalidation import DataValidation
# import threading
import openpyxl



class BackEndClass(QtWidgets.QWidget, Ui_MainWindow):

    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        self.setupUi(MainWindow)
        self.tabWidget.tabBar().setVisible(False)
        self.user_btn.clicked.connect(self.user_mode)
        self.audit_btn.clicked.connect(self.audit_mode)
        self.back_btn_user.clicked.connect(self.back_menu)
        self.back_btn_audit.clicked.connect(self.back_menu)


        # Audit Mode Buttons:
        self.insert_btn_audit.clicked.connect(self.insert_audit)
        self.pushButton_browse_audit.clicked.connect(self.browse_audit)
        self.Init_View()

        #Init Variables:
        self.Excel_Name = ""


    # Function called in __init__(), Fixes the view of user
    def Init_View(self):
        self.insert_btn_audit.setEnabled(False)
    def user_mode(self):
        self.tabWidget.setCurrentIndex(1)
    def audit_mode(self):
        self.tabWidget.setCurrentIndex(2)
    def back_menu(self):
        self.tabWidget.setCurrentIndex(0)

    def insert_audit(self):
        AssetTag = str(self.lineEdit_audit.text())
        print("Assert Tag: " + AssetTag)
        try:
            wb = openpyxl.load_workbook(self.Excel_Name)
            sheet = wb.active
            for i in range(1, sheet.max_row + 1):
                if(str(sheet.cell(row=i, column=2).value) == AssetTag):
                    sheet.cell(row=i, column=6).value = "True"
                    wb.save(self.Excel_Name)
                    self.lineEdit_audit.setText("")
                    print("here")
        except:
            QMessageBox.about(self, "Message", "Error ")


    #
    def browse_audit(self):
        Excel_File = QFileDialog.getOpenFileName(self, 'Open File', 'Select File', '(*.xlsx)')
        Excel_File = Excel_File[0]
        print(Excel_File)

        if Excel_File.endswith(".xlsx"):
            self.Excel_Name = Excel_File
            self.insert_btn_audit.setEnabled(True)
            self.lineEdit_audit.setReadOnly(False)
        else:
            QMessageBox.about(self, "Message", "Please select excel file type")
            self.insert_btn_audit.setEnabled(False)
            self.lineEdit_audit.setReadOnly(True)











if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = BackEndClass()
    MainWindow.show()
    sys.exit(app.exec_())



