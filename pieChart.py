import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from front import Ui_MainWindow  # Importing your generated UI
import numpy as np  # Add this import statement at the beginning of your file


class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=8, height=6, dpi=100):
        fig, self.ax = plt.subplots(figsize=(width, height), dpi=dpi)
        super(MplCanvas, self).__init__(fig)

    def plot_pie_chart(self, data, labels, colors=None, label_positions=None):
        self.ax.clear()
        wedges, texts, autotexts = self.ax.pie(
            data, colors=colors, autopct='%1.1f%%', startangle=90,
            pctdistance=0.5,  # Adjust distance of percentage text from center
        )

        # Customize the positions of the labels if provided
        if label_positions:
            for i, pos in enumerate(label_positions):
                angle = pos[0]
                distance = pos[1] if len(pos) > 1 else 1.05

                # Calculate position based on angle and distance
                x = np.cos(np.deg2rad(angle)) * distance
                y = np.sin(np.deg2rad(angle)) * distance

                # Place the label
                self.ax.text(x, y, labels[i], ha='center', va='center', fontsize=10, color='black')

        self.ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        self.draw()


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        # # Load the UI from the generated class
        # self.ui = Ui_MainWindow()
        # self.ui.setupUi(self)

        # Create Matplotlib canvases for four pie charts
        self.canvas_1 = MplCanvas(self, width=8, height=6, dpi=100)
        self.canvas_2 = MplCanvas(self, width=8, height=6, dpi=100)
        self.canvas_3 = MplCanvas(self, width=8, height=6, dpi=100)
        self.canvas_4 = MplCanvas(self, width=8, height=6, dpi=100)

        # Set up layouts to add the canvases directly into the widgets
        self.layout_1 = QVBoxLayout(self.ui.widget)
        self.layout_2 = QVBoxLayout(self.ui.widget_2)
        self.layout_3 = QVBoxLayout(self.ui.widget_3)
        self.layout_4 = QVBoxLayout(self.ui.widget_4)

        # Add the canvases to the respective layouts
        self.layout_1.addWidget(self.canvas_1)
        self.layout_2.addWidget(self.canvas_2)
        self.layout_3.addWidget(self.canvas_3)
        self.layout_4.addWidget(self.canvas_4)

        # Manually set the data (replace this with your data)
        labels_1 = ['Failed Trials ', 'Success Trials']
        data_1 = [50, 50]
        colors_1 = ['lightseagreen', 'red']  # Custom colors by name
        label_positions_1 = [(45, 1.5), (225, 1.5)]  # Example positions

        labels_2 = ['Scanned Assets', 'remaining Assets']
        data_2 = [45, 30]
        colors_2 = ['lightseagreen', 'red']  # Custom colors by name
        label_positions_2 = [(45, 1.5), (225, 1.5)]  # Example positions

        labels_3 = ['Mahle Scanned Assets', 'Mahle remaining Assets']
        data_3 = [50, 20]
        colors_3 = ['papayawhip', 'grey']  # Custom colors by name
        label_positions_3 = [(45, 1.5), (225, 1.5)]  # Example positions

        labels_4 = ['BS Scanned Assets', 'BS remaining Assets']
        data_4 = [40, 35]
        colors_4 = ['papayawhip', 'grey']  # Custom colors by name
        label_positions_4 = [(45, 1.5), (225, 1.5)]  # Example positions

        # Plot pie charts with custom colors
        self.canvas_1.plot_pie_chart(data_1, labels_1, colors_1, label_positions_1)
        self.canvas_2.plot_pie_chart(data_2, labels_2, colors_2, label_positions_2)
        self.canvas_3.plot_pie_chart(data_3, labels_3, colors_3, label_positions_3)
        self.canvas_4.plot_pie_chart(data_4, labels_4, colors_4, label_positions_4)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
