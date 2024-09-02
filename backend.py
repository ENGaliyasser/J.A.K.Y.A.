import importlib
import PyQt5.QtWidgets
from lxml import etree
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
import os
import re
from pyexcel.cookbook import merge_all_to_a_book
import glob
from PyQt5.QtWidgets import QMessageBox, QFileDialog
# import win32com.client as win32
from PyQt5.QtGui import QTextCursor
import numpy as np
from tabulate import tabulate
import datetime
import winsound
import pandas as pd
from openpyxl import workbook, load_workbook
import openpyxl
import traceback
from openpyxl.worksheet.datavalidation import DataValidation
import threading


class BackEndClass(QtWidgets.QWidget, Ui_MainWindow):

    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        self.setupUi(MainWindow)
        self.tabWidget.tabBar().setVisible(False)
        self.user_btn.clicked.connect(self.user_mode)
        self.audit_btn.clicked.connect(self.audit_mode)
        self.back_btn_user.clicked.connect(self.back_menu)
        self.back_btn_audit.clicked.connect(self.back_menu)
    def user_mode(self):
        self.tabWidget.setCurrentIndex(1)
    def audit_mode(self):
        self.tabWidget.setCurrentIndex(2)
    def back_menu(self):
        self.tabWidget.setCurrentIndex(0)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = BackEndClass()
    MainWindow.show()
    sys.exit(app.exec_())
