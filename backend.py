from front import Ui_MainWindow
from PyQt5 import QtCore, QtGui, QtWidgets, QtTest
import sys
from PyQt5.QtWidgets import QMessageBox, QFileDialog
import openpyxl
from PyQt5.QtCore import QTimer

# Global Variable to store the password
password = "jakya2024"
# Functions
def check_password(entered_pass):
    """Checks if the entered password is correct.

    Args:
        entered_pass: string of the password to check.

    Returns:
        True if the password is correct, False otherwise.
    """
    if entered_pass == password:
        return True
    else:
        return False

def check_excel_format(excel_file):
    """Checks if the excel file has the correct format.

    Args:
        excel_file: Path to the Excel file.

    Returns:
        True if the sheet format is correct, False otherwise.
    """
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    # defining the excpected columns
    expected_columns = ["Asset Name", "Asset Tag", "Serial", "Model", "Assigne"]
    # Get the actual number of columns in the first row
    actual_columns = [sheet.cell(row=1, column=i).value for i in range(1, 6)]
    print(actual_columns)
    # Check if the number of columns is correct
    if len(actual_columns) != len(expected_columns):
        return False

    # Check if the column titles match the expected titles
    if actual_columns != expected_columns:
        return False
    # All correct Return True
    return True

class BackEndClass(QtWidgets.QWidget, Ui_MainWindow):

    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        self.setupUi(MainWindow)
        self.tabWidget.tabBar().setVisible(False)
        self.user_btn.clicked.connect(self.user_mode)
        self.audit_btn.clicked.connect(self.audit_mode)
        self.back_btn_user.clicked.connect(self.back_menu)
        self.back_btn_audit.clicked.connect(self.back_menu)
        self.back_btn_user_2.clicked.connect(self.back_menu)
        self.pushButton_browse_audit.clicked.connect(self.browse_excel)
        self.pushButton_login.clicked.connect(self.login)
        self.insert_btn_audit.clicked.connect(self.insert_audit)
        self.new_login()

    def new_login(self):
        """
        Prepares the UI for a new login attempt by resetting fields and adjusting widget visibility.
        - Clears the password and Excel path fields.
        - Hides Excel-related widgets.
        - Shows password-related widgets.
        - Enables the password field and login button.
        """
        # Clear line edits
        self.lineEdit_password_audit_2.clear()
        self.lineEdit_excel_audit.clear()
        # Hide Excel related widgets
        self.label_excel.hide()
        self.lineEdit_excel_audit.hide()
        self.pushButton_browse_audit.hide()
        # Show password related widgets
        self.lineEdit_password_audit_2.show()
        self.pushButton_login.show()
        self.password_label_2.show()
        # enable login button and write password
        self.lineEdit_password_audit_2.setReadOnly(False)
        self.pushButton_login.setDisabled(False)
        self.Excel_Name = None
        self.pushButton_Scan.clicked.connect(self.Scan_user)




    def user_mode(self):
        """
        switches to user mode tab
        """
        self.tabWidget.setCurrentIndex(1)
    def audit_mode(self):
        """
        Switches to audit mode login tab
        """
        self.tabWidget.setCurrentIndex(2)

    def back_menu(self):
        """
        Returns back to the main menu
        """
        self.tabWidget.setCurrentIndex(0)
        self.new_login()


    def login(self):
        """
        Handles user login by checking the entered password and updating the UI accordingly.

        - If the password is correct:
            - Disables further password input.
            - Shows Excel-related widgets.
            - Displays a success message.
        - If the password is incorrect:
            - Displays an error message.
        """
        entered_password = self.lineEdit_password_audit_2.text()
        # Password Check
        if check_password(entered_password):
            # self.lineEdit_password_audit_2.clear()
            self.lineEdit_password_audit_2.setReadOnly(True)
            self.pushButton_login.setDisabled(True)
            # self.password_label_2.hide()
            self.label_excel.show()
            self.lineEdit_excel_audit.show()
            self.pushButton_browse_audit.show()
            QMessageBox.about(self, "Message", "Correct Password! Select an Excel file.")
        else:
            QMessageBox.about(self, "Message", "Invalid password! Please Enter the Correct Password.")

    def browse_excel(self):
        """
        Opens a file dialog for selecting an Excel file and updates the UI based on the selection.
        - Prompts the user to select an Excel (.xlsx) file.
        - If a valid file is selected:
            - Stores the file path.
            - Displays the path in the UI.
            - If the file format is correct, shows a success message and starts the audit.
            - If the format is incorrect, shows an error message.
        - If no file is selected, shows an error message.
        """
        Excel_File = QFileDialog.getOpenFileName(self, 'Open File', 'Select File', '(*.xlsx)')
        Excel_File = Excel_File[0]
        if Excel_File.endswith(".xlsx"):
            self.Excel_Name = Excel_File
            self.lineEdit_excel_audit.setText(Excel_File)
            if check_excel_format(Excel_File):
                self.msg_box = QMessageBox(self)
                self.msg_box.setWindowTitle("Success")
                self.msg_box.setText("Let's start the audit")
                self.msg_box.show()
                # Create a QTimer to close the message box after 2 seconds
                QTimer.singleShot(2000, self.enter_the_audit)
            else:
                QMessageBox.about(self, "Message", "Please select a valid Excel file.")
        else:
            QMessageBox.about(self, "Message", "No File selected!")

    def enter_the_audit(self):
        """
        Closes the message box and switches to the audit tab.
        """
        self.msg_box.close()
        self.tabWidget.setCurrentIndex(3)
        self.Display_audit()


    def insert_audit(self):
        AssetTag = str(self.lineEdit_audit.text())
        print("Asset Tag: " + AssetTag)
        try:
            try:
                wb = openpyxl.load_workbook("Audit_Output.xlsx")
            except Exception as e:
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
                wb.save("Audit_Output.xlsx")
                self.Display_audit()
            else:
                QMessageBox.about(self, "Message", "Asset Tag not found in the Excel sheet")

        except Exception as e:
            QMessageBox.about(self, "Message", "Error: " + str(e))

    def Display_audit(self):
        try:
            wb = openpyxl.load_workbook("Audit_Output.xlsx")
        except Exception as e:
            wb = openpyxl.load_workbook(self.Excel_Name)
        sheet = wb.active
        self.textBrowser_6.setText(str(sheet.max_row))
        self.textBrowser_13.setText("0")
        self.textBrowser_14.setText("0")
        k = 0
        for i in range(1, sheet.max_row + 1):
            if sheet.cell(row=i, column=6).value == "True":
                k += 1
        self.textBrowser_14.setText(str(k))
        self.textBrowser_15.setText(str(sheet.max_row - k))
        self.textBrowser_16.setText(str(k))


    #TODO: Use the live excel sheet
    def Scan_user(self):
        AssetTag = str(self.lineEdit.text())
        print("Asset Tag: " + AssetTag)
        try:
            #Change this
            wb = openpyxl.load_workbook("Audit_Output.xlsx")
            sheet = wb.active
            row_to_update = None

            for i in range(1, sheet.max_row + 1):
                if str(sheet.cell(row=i, column=2).value) == AssetTag:
                    # Mark this row as "True" in column 6
                    self.textBrowser.setText(str(sheet.cell(row = i, column = 1).value))
                    self.textBrowser_2.setText(str(sheet.cell(row = i, column = 4).value))
                    self.textBrowser_3.setText(str(sheet.cell(row = i, column = 5).value))
                    row_to_update = i
                    break

        except Exception as e:
            QMessageBox.about(self, "Message", "Error: " + str(e))



if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = BackEndClass()
    MainWindow.show()
    sys.exit(app.exec_())
