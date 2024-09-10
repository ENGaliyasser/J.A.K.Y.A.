# Author: Khaled Waleed & Ali Yasser
# Date: 9 September 2024
# Description: This script is designed to interface with a USB serial device, specifically a scanner. 
#              It uses PyQt5 to handle threading and signal communication. The script includes functionality 
#              to automatically find a connected USB serial device and read data from it. The data is emitted 
#              via a PyQt5 signal for further processing or display in a GUI application.
import serial
import serial.tools.list_ports
from PyQt5.QtCore import QCoreApplication, QThread, pyqtSignal, QObject

def find_usb_serial_device():
    ports = serial.tools.list_ports.comports()
    for port in ports:
        if 'USB' in port.description:
            return port.device
    return None

class Scanner(QObject):
    data_received = pyqtSignal(str)

    def __init__(self, ser):
        super().__init__()
        self.ser = ser
        self._is_running = True

    def run(self):
        while self._is_running:
            if self.ser.in_waiting > 0:
                data = self.ser.readline().decode('utf-8').rstrip()
                print("Received data: {0}".format(data))
                self.data_received.emit(data)

    def stop(self):
        self._is_running = False
