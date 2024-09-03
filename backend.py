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
                    self.insert_asset_tag(AssetTag, i)
        except:
            QMessageBox.about(self, "Message", "Error ")

    from PyQt5.QtWidgets import QLabel, QMessageBox
    import openpyxl


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

    def display_true_row(self):
        try:
            # Ensure the workbook is reloaded to reflect any recent changes
            wb = openpyxl.load_workbook(self.Excel_Name, data_only=True)
            sheet = wb.active
            found_true = False  # Flag to check if any "True" value is found

            for i in range(1, sheet.max_row + 1):
                # Check if the row is marked as "True" in column 6
                if str(sheet.cell(row=i, column=6).value).strip() == "True":
                    # Read the row and convert it to a string for display
                    row_data = [str(sheet.cell(row=i, column=col).value) for col in range(1, sheet.max_column + 1)]
                    self.display_label.setText(" | ".join(row_data))
                    found_true = True
                    break

            if not found_true:
                self.display_label.setText("No rows marked as 'True' found.")

        except Exception as e:
            QMessageBox.about(self, "Message", "Error: " + str(e))

    def insert_asset_tag(self, asset_tag, row_number):
        try:
            wb = openpyxl.load_workbook(self.Excel_Name)
            sheet = wb.active

            # Insert asset tag in the specified row
            sheet.cell(row=row_number, column=1).value = asset_tag
            # Mark this row as "True" in column 6
            sheet.cell(row=row_number, column=6).value = "True"

            # Save the workbook to persist changes
            wb.save(self.Excel_Name)

            # Now update the display label
            self.display_true_row()

        except Exception as e:
            QMessageBox.about(self, "Message", "Error: " + str(e))


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = BackEndClass()
    MainWindow.show()
    sys.exit(app.exec_())



