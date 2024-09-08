from front import Ui_MainWindow
from PyQt5 import QtCore, QtGui, QtWidgets, QtTest
import sys
from PyQt5.QtWidgets import QMessageBox, QFileDialog, QApplication, QMainWindow, QLineEdit, QVBoxLayout, QWidget, QTabWidget
import openpyxl
from PyQt5.QtCore import QTimer, pyqtSignal, QObject, QThread
import os
import sys
from PyQt5 import QtWidgets
from updater import Updater
import time
from datetime import date

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


class FocusRedirector(QObject):
    """
    FocusRedirector is a class that ensures a QLineEdit always 
    receives focus whenever certain events occur within the 
    parent widget it monitors.

    Attributes
    ----------
    target : QWidget
        The target widget that should always receive focus.

    Methods
    -------
    eventFilter(obj, event)
        Monitors and handles events, redirecting focus to the target widget 
        when specific events occur.
    """
    def __init__(self, target):
        super().__init__()
        self.target = target

    def eventFilter(self, obj, event):
        """
        Monitors and handles events, redirecting focus to the target widget 
        when specific events occur.

        Parameters
        ----------
        obj : QObject
            The object for which the event is being filtered.
        event : QEvent
            The event to be filtered and possibly handled.

        Returns
        -------
        bool
            True if the event should be filtered out, False otherwise.
        """
        if event.type() in (event.FocusIn, event.MouseButtonPress):
            self.target.setFocus()
        return super().eventFilter(obj, event)

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
        #self.insert_btn_audit.clicked.connect(self.insert_audit)
        self.new_login()
        self.lineEdit_audit.returnPressed.connect(self.insert_audit)
        self.scan_user.returnPressed.connect(self.Scan_user)
        self.scan_user.setFocus()

        # Initialize Updater with the current version and repository details
        self.updater = Updater(
            current_version="v2.00",  # Replace with your tool's version
            repo_owner="ENGaliyasser",
            repo_name="JAKYA",
            progress_bar=self.progressBar
        )

        # Connect the update button to the updater's update function
        self.update.clicked.connect(self.updater.update_application)

        # Hide update-related UI elements initially
        self.progressBar.setVisible(False)
        self.update.setVisible(False)
        self.ask.setVisible(False)

        # Connect the "Check" button to the check_update function
        self.check.clicked.connect(lambda: self.updater.check_update(self.update, self.ask))

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
        #self.pushButton_scan.clicked.connect(self.Scan_user)
        self.right.hide()


    def user_mode(self):
        """
        switches to user mode tab
        """
        self.tabWidget.setCurrentIndex(1)
        self.scan_user.setFocus()
        self.scan_user.setCursor(QtCore.Qt.BlankCursor)
        self.scan_user.clear()
        # Remove other event filters
        try:
            self.tabWidget.removeEventFilter(self.focus_redirector_audit)
        except:
            pass
        # Install the event filter on the tab widget in user mode
        self.focus_redirector_user = FocusRedirector(self.scan_user)
        self.tabWidget.installEventFilter(self.focus_redirector_user)

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
            # QMessageBox.about(self, "Message", "Correct Password! Select an Excel file.")
            self.right.show()
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
            if self.check_excel_format(Excel_File):
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
        self.textBrowser_5.setText(str(date.today()))
        self.textBrowser_4.setText(os.path.basename(str(self.Excel_Name)))
        #self.lineEdit_audit.setVisible(False)
        self.lineEdit_audit.setFocus()
        self.lineEdit_audit.setCursor(QtCore.Qt.BlankCursor)
        # Remove othe event filters
        try:
            self.tabWidget.removeEventFilter(self.focus_redirector_user)
        except:
            pass
        # Install the event filter on the tab widget in audit mode
        self.focus_redirector_audit = FocusRedirector(self.lineEdit_audit)
        self.tabWidget.installEventFilter(self.focus_redirector_audit)
        

# TODO:
# on scanning if the item is already scanned tell the user in a label. 
    def insert_audit(self):
        """
           - Retrieves the asset tag entered by the user from the input field.
           - Attempts to open 'Audit_Output.xlsx' for updates; if not available, opens the original Excel file.
           - Searches for the asset tag in column 2 of the sheet.
           - If the asset tag is found, marks "True" in column 6 of the corresponding row.
           - Saves the updated data to 'Audit_Output.xlsx' and updates the display.
           - If the asset tag is not found, shows an error message.
           - Catches and handles any errors that occur during the process.
           """
        self.start_scan_label.clear()
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
                if str(sheet.cell(row=i, column=self.asset_tag_col).value) == AssetTag:
                    # Mark this row as "True" in column 6
                    if sheet.cell(row=i, column=self.status_col).value == "True":
                        self.start_scan_label.setText("Item Already Scanned!")
                    sheet.cell(row=i, column=self.status_col).value = "True"
                    row_to_update = i
                    break

            if row_to_update:
                wb.save("Audit_Output.xlsx")
                self.Display_audit()
            else:
                if not self.lineEdit_audit.text() == "":
                    QMessageBox.about(self, "Message", "Asset Tag not found in the Excel sheet")

        except Exception as e:
            QMessageBox.about(self, "Message", "Error: " + str(e))
        self.lineEdit_audit.clear()
        self.lineEdit_audit.setFocus()

    def Display_audit(self):
        """
            - Attempts to open 'Audit_Output.xlsx'; if not found, opens the original Excel file.
            - Displays the total number of rows in the sheet in 'textBrowser_6'.
            - Initializes counters for marked ("True") and unmarked rows.
            - Iterates through the rows to count how many rows are marked as "True" in column 6.
            - Updates 'textBrowser_14' and 'textBrowser_16' with the count of marked rows.
            - Updates 'textBrowser_15' with the count of unmarked rows.
            """
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
            if sheet.cell(row=i, column=self.status_col).value == "True":
                k += 1
        self.textBrowser_14.setText(str(k))
        self.textBrowser_15.setText(str(sheet.max_row - k))
        self.textBrowser_16.setText(str(k))


    #TODO: Use the live excel sheet
    # Auto focus - error handling - reset labels
    def Scan_user(self):
        """
            - Retrieves the asset tag entered by the user from the input field.
            - Opens 'Audit_Output.xlsx' to check for the asset tag in column 2.
            - If a matching asset tag is found, it updates several text browsers with values from that row.
            - Displays the asset name,model, and assigned user in the text browsers.
            - Catches and handles any errors that occur during the process.
            """
        # Reset Labels on each new scan
        self.asset_name.clear()
        self.asset_tag.clear()
        self.serial.clear()
        self.checked_out_to.clear()
        self.owner.clear()

        AssetTag = str(self.scan_user.text())
        print("Asset Tag: " + AssetTag)
        try:
            wb = openpyxl.load_workbook("Audit_Output.xlsx")
            sheet = wb.active
            row_to_update = None
            for i in range(1, sheet.max_column + 1):
                if (sheet.cell(row = 1, column = i).value).lower() == "asset tag":
                    self.asset_tag_col = i
                if (sheet.cell(row = 1, column = i).value).lower() == "serial":
                    self.serial_col = i
                    print(self.serial_col)
                if (sheet.cell(row = 1, column = i).value).lower() == "checked out to":
                    self.check_out_to_col = i
                if (sheet.cell(row = 1, column = i).value).lower() == "owner":
                    self.owner_col = i
                if (sheet.cell(row = 1, column = i).value).lower() == "asset name":
                    self.asset_name_col = i

            for i in range(1, sheet.max_row + 1):
                if str(sheet.cell(row=i, column=self.asset_tag_col).value) == AssetTag:
                    # Mark this row as "True" in column 6
                    self.asset_name.setText(str(sheet.cell(row = i, column = self.asset_name_col).value))
                    self.asset_tag.setText(str(sheet.cell(row = i, column = self.asset_tag_col).value))
                    self.serial.setText(str(sheet.cell(row = i, column = self.serial_col).value))
                    self.checked_out_to.setText(str(sheet.cell(row = i, column = self.check_out_to_col).value))
                    self.owner.setText(str(sheet.cell(row = i, column = self.owner_col).value))
                    row_to_update = i
                    break
            if row_to_update == None:
                QMessageBox.about(self, "Message", "Asset Tag not found in the Excel sheet")

        except Exception as e:
            QMessageBox.about(self, "Message", "Error: " + str(e))
        self.scan_user.clear()

    def check_excel_format(self, excel_file):
        """Checks if the excel file has the correct format.

        Args:
            excel_file: Path to the Excel file.

        Returns:
            True if the sheet format is correct, False otherwise.
        """
        if(excel_file != os.getcwd().replace("\\", "/") + "/Audit_Output.xlsx"):
            try:
                workbook = openpyxl.load_workbook("Audit_Output.xlsx")
                QMessageBox.about(self, "Error", "Please move the Audit_Output.xlsx file to a different folder, or choose it as the Excel file.")
                return False
            except Exception as e:
                workbook = openpyxl.load_workbook(excel_file)
        else:
            workbook = openpyxl.load_workbook("Audit_Output.xlsx")

        sheet = workbook.active
        # defining the excpected columns
        # TODO:
        # Add Model, Category, Status, Checked Out To, Location, Purchase Cost, HS-Code, Owner
        # Important (ERROR): Asset Name, Asset Tag, Serial, Checked Out To, Owner (USER MODE)
        # Warning: The rest
        expected_columns = set(["asset name", "asset tag", "serial", "checked out to", "model", "category", "status",
                            "location", "purchase cost", "hs-code", "owner"])
        important_columns = set(["asset name", "asset tag", "serial", "checked out to", "owner"])

        # Get the actual number of columns in the first row
        actual_columns = set([sheet.cell(row=1, column=i).value.lower() for i in range(1, sheet.max_column + 1)])
        print(actual_columns)
        print(important_columns)

        # Check if the column titles match the expected titles
        if not important_columns.issubset(actual_columns):
            return False
        # All correct Return True
        flag_status = False
        for i in range(1, sheet.max_column + 1):
            if (sheet.cell(row = 1, column = i).value).lower() == "asset tag":
                self.asset_tag_col = i
            if (sheet.cell(row = 1, column = i).value).lower() == "status":
                self.status_col = i
                flag_status = True
            if (sheet.cell(row = 1, column = i).value).lower() == "serial":
                self.serial_col = i
                print(self.serial_col)
            if (sheet.cell(row = 1, column = i).value).lower() == "check out to":
                self.check_out_to_col = i
            if (sheet.cell(row = 1, column = i).value).lower() == "owner":
                self.owner_col = i
            if (sheet.cell(row = 1, column = i).value).lower() == "asset name":
                self.asset_name_col = i
            

        print("here2")

        if not flag_status:
            self.status_col = sheet.max_row + 1
            sheet.cell(row = 1, column = self.status_col).value = "Status"
        workbook.save("Audit_Output.xlsx")
        return True


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    Updater.delete_old_versions(current_version="v2.00")  # Replace with your tool's version
    MainWindow = QtWidgets.QMainWindow()
    ui = BackEndClass()
    MainWindow.show()
    sys.exit(app.exec_())
