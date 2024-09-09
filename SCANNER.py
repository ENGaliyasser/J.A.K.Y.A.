import serial
import serial.tools.list_ports

def find_usb_serial_device():
    ports = serial.tools.list_ports.comports()
    for port in ports:
        if 'USB' in port.description:
            return port.device
    return None

def read_data(ser):
    while True:
        if ser.in_waiting > 0:
            data = ser.readline().decode('utf-8').rstrip()
            print("Received data: {0}".format(data))

# Find the USB serial device
usb_serial_port = find_usb_serial_device()

if usb_serial_port:
    # Configure the serial port
    ser = serial.Serial(
        port=usb_serial_port,
        baudrate=9600,
        timeout=1
    )

    if ser.is_open:
        print(f"Connected to {usb_serial_port}. Waiting for data... Press Ctrl+C to exit.")
        try:
            read_data(ser)
        except KeyboardInterrupt:
            print("\nExiting...")
        finally:
            ser.close()
    else:
        print("Failed to open the selected port.")
else:
    print("No USB serial device found.")
