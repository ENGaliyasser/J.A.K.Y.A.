import sys
import openpyxl
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox, QFileDialog

# Assuming Ui_MainWindow is the generated class from Qt Designer
from front import Ui_MainWindow
# Password Check
password = "JAKYA2024" # Audit Password
def check_password(entered_pass):
    if entered_pass == password:
        return True
    else:
        return False
class BackEndClass(QtWidgets.QWidget, Ui_MainWindow):

    def __init__(self):
        super().__init__()
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

        # Init Variables:
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

    def browse_audit(self):
        entered_password = self.lineEdit_password_audit.text()
        if check_password(entered_password) == True:
            Excel_File = QFileDialog.getOpenFileName(self, 'Open File', 'Select File', '(*.xlsx)')
            Excel_File = Excel_File[0]
            print(Excel_File)
            if Excel_File.endswith(".xlsx"):
                self.Excel_Name = Excel_File
                self.insert_btn_audit.setEnabled(True)
                self.lineEdit_audit.setReadOnly(False)
            else:
                QMessageBox.about(self, "Message", "Please select an Excel file type")
                self.insert_btn_audit.setEnabled(False)
                self.lineEdit_audit.setReadOnly(True)
        else:
            QMessageBox.about(self, "Message", "Please Enter the correct Password")

    def insert_audit(self):
        entered_password = self.lineEdit_password_audit.text()
        if check_password(entered_password) == True:
            AssetTag = str(self.lineEdit_audit.text())
            print("Asset Tag: " + AssetTag)
            try:
                wb = openpyxl.load_workbook(self.Excel_Name)
                sheet = wb.active
                row_to_update = None

                for i in range(1, sheet.max_row + 1):
                    if str(sheet.cell(row=i, column=2).value) == AssetTag:
                        # Mark this row as "True" in column 6
                        sheet.cell(row=i, column=6).value = "True"
                        row_to_update = i
                        break

                if row_to_update:
                    # Save the workbook to persist changes
                    wb.save(self.Excel_Name)
                    # Clear the input field
                    self.lineEdit_audit.setText("")
                    print(f"Row {row_to_update} updated with 'True'")
                    # Update the display label with the updated row
                    self.display_row(row_to_update)
                else:
                    QMessageBox.about(self, "Message", "Asset Tag not found in the Excel sheet")

            except Exception as e:
                QMessageBox.about(self, "Message", "Error: " + str(e))
        else:
            QMessageBox.about(self, "Message", "Please Enter the correct Password")

    def display_row(self, row_number):
        try:
            # Ensure the workbook is reloaded to reflect any recent changes
            wb = openpyxl.load_workbook(self.Excel_Name, data_only=True)
            sheet = wb.active

            # Read the specific row and convert it to a string for display
            row_data = [str(sheet.cell(row=row_number, column=col).value) for col in range(1, sheet.max_column + 1)]
            self.textEdit_user.setText(" | ".join(row_data))

        except Exception as e:
            QMessageBox.about(self, "Message", "Error: " + str(e))


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = BackEndClass()
    MainWindow.show()
    sys.exit(app.exec_())
