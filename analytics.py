from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QMessageBox
from PyQt5.QtCore import Qt, pyqtSignal, QThread
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import os


class DataLoader(QThread):
    data_loaded = pyqtSignal(pd.DataFrame)
    error_occurred = pyqtSignal(str)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        try:
            df = pd.read_excel(self.file_path)
            self.data_loaded.emit(df)
        except Exception as e:
            self.error_occurred.emit(f"Error loading Excel file: {e}")

class AnalyticsTab(QWidget):
    def __init__(self, directory, parent=None):
        super(AnalyticsTab, self).__init__(parent)

        self.layout = QVBoxLayout(self)
        self.title_label = QLabel('Parcel Analytics Dashboard', self)
        self.title_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.title_label)

        self.figure = plt.figure(figsize=(10, 8))
        self.canvas = FigureCanvas(self.figure)
        self.layout.addWidget(self.canvas)

        self.setMinimumSize(900, 700)
        self.setWindowTitle('Parcel Analytics Dashboard')
        self.load_data(directory)

    def load_data(self, directory=None):
        file_path = os.path.join(directory or os.path.join(os.environ['USERPROFILE'], 'Desktop'), 'final.xlsx')
        self.data_loader = DataLoader(file_path)
        self.data_loader.data_loaded.connect(self.plot_pie_chart)
        self.data_loader.error_occurred.connect(self.show_error_message)
        self.data_loader.start()

    def plot_pie_chart(self, df):
        status_mapping = {
            'Return': ['Returned to shipper', 'Being Return', 'Pickup Request Sent', 'Ready for Return'],
            'Pending': ['Pending'],
            'Delivered': ['Delivered'],
            'In Transit': ['Arrived at Station', 'Dispatched', 'Assign to Courier']
        }

        df['Status'] = df['Recent Location'].apply(
            lambda x: next((status for status, values in status_mapping.items() if x in values), 'In Transit')
        )

        status_counts = df['Status'].value_counts()
        
        if status_counts.empty:
            QMessageBox.warning(self, "No Data", "No statuses found in the data.")
            return

        labels = status_counts.index
        sizes = status_counts.values

        explode = [0.1 if size < 0.05 * sizes.sum() else 0.05 for size in sizes]
        
        if 'Return' in labels:
            explode[labels.tolist().index('Return')] = 0.2

        self.figure.clear()
        
        ax = self.figure.add_subplot(111)
        
        colors = plt.get_cmap('Set3').colors[:len(labels)]

        wedges, texts, autotexts = ax.pie(
            sizes,
            labels=labels,
            autopct='%1.1f%%',
            explode=explode,
            startangle=90,
            colors=colors,
            pctdistance=0.85,
            wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
        )

        for text, autotext in zip(texts, autotexts):
            text.set_fontsize(12)
            autotext.set_fontsize(12)
            text.set_color('black')
            autotext.set_color('darkblue')

        legend_labels = [f"{label} ({count})" for label, count in zip(labels, sizes)]
        
        ax.legend(wedges, legend_labels, title="Statuses", loc="center left", bbox_to_anchor=(0.85, 0.5), fontsize=12)

        ax.axis('equal')

        self.canvas.draw()

    def show_error_message(self, message):
        QMessageBox.critical(self, "Error", message)
