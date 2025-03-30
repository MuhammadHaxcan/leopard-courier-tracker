from PyQt5.QtWidgets import QMessageBox
import os
import sys
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QLabel
                             , QFileDialog, QMessageBox, QLineEdit, 
                             QHBoxLayout, QProgressBar, QVBoxLayout, QGroupBox, QGridLayout)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
from analytics import AnalyticsTab
from thread import WorkerThread
import threading
import pandas as pd
from leopard import LeopardCourierAPI
from config import load_config, save_config
from utils import extract_data_from_html, rename_file_extension, delete_temporary_files, is_connected, open_excel_file
from excel_operations import customize_excel, save_data_to_excel, add_columns, append_to_final, calculate_payments, sort_by_booking_date


class FileConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CaseDrip Leopard")
        self.setAcceptDrops(True)
        self.setMinimumWidth(900)
        self.setStyleSheet("""
            QWidget {
                font-family: 'Segoe UI', Helvetica, Arial, sans-serif;
                font-size: 14px;
                background-color: #f5f5f5;
            }
            QGroupBox {
                border: 1px solid #d6d6d6;
                border-radius: 5px;
                font: 18px;
                font-weight: bold;
                padding-bottom:2px;
                margin-top: 0.3em;
                padding: 12px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                color: #37352f;
                padding: 0 8px;
                font-weight: 500;
            }
            QPushButton {
                background-color: #ffffff;
                color: #37352f;
                border: 1px solid #d6d6d6;
                padding: 8px 16px;
                margin-top: 1px;
                border-radius: 5px;
                font-weight: 500;
                box-shadow: inset 0px 0px 0px rgba(0, 0, 0, 0); /* Neutral initial shadow */
            }

            QPushButton:hover {
                background-color: qlineargradient(
                    spread:pad, x1:0, y1:0, x2:1, y2:1, 
                    stop:0 #f7f7f7, stop:1 #eaeaea
                ); /* Gradient for depth */
                border: 1px solid #bcbcbc; /* Slightly darker border */
                box-shadow: 3px 3px 6px rgba(0, 0, 0, 0.2); /* Outer shadow for 3D look */
            }

            QPushButton:pressed {
                background-color: #e0e0e0; /* Slightly darker for press effect */
                box-shadow: inset 2px 2px 4px rgba(0, 0, 0, 0.2); /* Inner shadow to create a "pressed" look */
                transform: translateY(2px); /* Move button down slightly */
            }


            QPushButton:disabled {
                color: #a3a3a3;
                border-color: #e0e0e0;
                background-color: #f5f5f5;
            }
            QLineEdit {
                border: 1px solid #d6d6d6;
                padding: 8px 12px;
                margin-top:4px;
                border-radius: 4px;
                background-color: #ffffff;
            }
            QLabel {
                color: #37352f;
                background-color: None;
                padding: 0px;
                margin:0px;
            }
            QProgressBar {
                border: 1px solid  #cccccc;
                background-color: #f1f1f1;
                height: 8px;
                border-radius: 3px;
            }
            QProgressBar::chunk {
                background-color: #2eaadc;
                border-radius: 3px;
            }
        """)

        # Initialize variables
        self.file_path = None
        self.api_key, self.api_password, self.final_xlsx_directory = load_config()

        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(30, 30, 30, 30)

        # Add sections
        main_layout.addWidget(self.create_header_section())
        main_layout.addWidget(self.create_api_section())
        main_layout.addWidget(self.create_operations_section())
        main_layout.addWidget(self.create_payment_section())
        
        # Drag-and-Drop area expands to fill available space
        main_layout.addStretch()
        main_layout.addWidget(self.create_status_section(), stretch=1)

        # Hook up button actions
        self.upload_button.clicked.connect(self.upload_file)
        self.convert_button.clicked.connect(self.convert_file)
        self.track_button.clicked.connect(self.track_existing_parcels)
        self.directory_button.clicked.connect(self.select_directory)
        self.track_payment_button.clicked.connect(self.track_existing_payments)
        self.show_analytics_button.clicked.connect(self.show_analytics)
        self.calculate_payments_button.clicked.connect(self.calculate_and_update_payments)


    def create_header_section(self):
        container = QWidget()
        layout = QHBoxLayout(container)  # Use QHBoxLayout for horizontal alignment
        layout.setContentsMargins(0, 0, 0, 20)

        # Title Label
        title = QLabel("📦 CASEDRIP COD TRACKING")
        title.setStyleSheet("""
            QLabel {
                font-family: 'Montserrat', sans-serif; /* Expressive font families */
                font-size: 22px; /* Slightly larger size for emphasis */
                font-weight: bold; /* Strong emphasis */
                color: #2a2a2a; /* Darker shade for better visibility */
                letter-spacing: 1px; /* Slight spacing between letters for a modern look */
            }
        """)

        # Open File Button
        open_file_button = QPushButton("Open Excel Sheet")
        open_file_button.setCursor(Qt.PointingHandCursor)
        open_file_button.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                color: #37352f;
                border: 1px solid #d6d6d6;
                padding: 8px 16px;
                border-radius: 5px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #f7f7f7;
            }
        """)
        open_file_button.clicked.connect(self.open_final_excel)  # Link to open_final_excel method

        # API Strength Box
        api_strength_container = QWidget()
        api_strength_layout = QHBoxLayout(api_strength_container)
        api_strength_layout.setContentsMargins(0, 0, 0, 0)
        api_strength_layout.setSpacing(5)

        self.api_strength_label = QLabel("📡 Checking API Strength...")
        self.api_strength_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                color: #37352f;
                font-weight: bold;
            }
        """)

        refresh_button = QPushButton("🔄")
        refresh_button.setCursor(Qt.PointingHandCursor)
        refresh_button.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                color: #37352f;
                border: 1px solid #d6d6d6;
                padding: 4px 10px;
                border-radius: 5px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #f7f7f7;
            }
        """)
        refresh_button.clicked.connect(self.update_api_strength)

        api_strength_layout.addWidget(self.api_strength_label)
        api_strength_layout.addWidget(refresh_button)

        # Add widgets to main layout
        layout.addWidget(title)
        layout.addStretch()  # Pushes other elements to the right
        layout.addWidget(api_strength_container)
        layout.addWidget(open_file_button)

        return container

    def create_api_section(self):
        group = QGroupBox("🔑 API Configuration")
        layout = QHBoxLayout()
        layout.setSpacing(15)

        self.api_key_input = QLineEdit()
        self.api_key_input.setPlaceholderText("Enter API Key")
        self.api_key_input.setText(self.api_key)

        self.api_password_input = QLineEdit()
        self.api_password_input.setPlaceholderText("Enter API Password")
        self.api_password_input.setEchoMode(QLineEdit.Password)
        self.api_password_input.setText(self.api_password)

        layout.addWidget(self.api_key_input)
        layout.addWidget(self.api_password_input)
        group.setLayout(layout)
        return group

    def create_operations_section(self):
        group = QGroupBox("📋 File Operations")
        layout = QVBoxLayout()
        layout.setSpacing(15)

        buttons_layout = QHBoxLayout()
        buttons_layout.setContentsMargins(0, 4, 0, 0)
        buttons_layout.setSpacing(10)

        self.upload_button = self.create_button("Upload LoadSheet", "📄")
        self.convert_button = self.create_button("Convert & Extract", "⚡")
        self.track_button = self.create_button("Track Existing", "🔍")
        self.directory_button = self.create_button("Select Directory", "📁")

        self.convert_button.setEnabled(False)

        buttons_layout.addWidget(self.upload_button)
        buttons_layout.addWidget(self.convert_button)
        buttons_layout.addWidget(self.track_button)
        buttons_layout.addWidget(self.directory_button)

        # Create QLabel for displaying the selected directory
        self.directory_label = QLabel(f"📂 Selected Directory: {self.final_xlsx_directory or 'None'}")
        self.directory_label.setStyleSheet("color: #787774;")

        layout.addLayout(buttons_layout)
        layout.addWidget(self.directory_label)  # Add QLabel to the layout
        group.setLayout(layout)
        return group

    def create_payment_section(self):
        group = QGroupBox("💰 Payment Overview")
        group_layout = QVBoxLayout()

        #add style to the group layout
        # Payment Information Grid
        info_grid_layout = QGridLayout()
        info_grid_layout.setSpacing(20)
        info_grid_layout.setContentsMargins(0, 0, 0, 10)

        # Pending Payments
        pending_payment_title = QLabel("Pending Payments")
        pending_payment_title.setStyleSheet("font-size: 16px; color: #787774;")
        self.pending_payment_label = QLabel("Rs. 0")
        self.pending_payment_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #37352f;")
        info_grid_layout.addWidget(pending_payment_title, 0, 0)
        info_grid_layout.addWidget(self.pending_payment_label, 1, 0)

        # Total Payments
        total_payment_title = QLabel("Total Payments")
        total_payment_title.setStyleSheet("font-size: 16px; color: #787774;")
        self.total_payment_label = QLabel("Rs. 0")
        self.total_payment_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #37352f;")
        info_grid_layout.addWidget(total_payment_title, 0, 2)
        info_grid_layout.addWidget(self.total_payment_label, 1, 2)

        # Delivered Pending
        delivered_pending_title = QLabel("Delivered Pending")
        delivered_pending_title.setStyleSheet("font-size: 16px; color: #787774;")
        self.delivered_pending_label = QLabel("Rs. 0")
        self.delivered_pending_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #37352f;")
        info_grid_layout.addWidget(delivered_pending_title, 0, 1)
        info_grid_layout.addWidget(self.delivered_pending_label, 1, 1)

        # Action Buttons
        button_layout = QHBoxLayout()
        button_layout.setSpacing(15)

        self.track_payment_button = self.create_button("Track Payments", "💳")
        self.calculate_payments_button = self.create_button("Calculate Payments", "🧾")
        self.show_analytics_button = self.create_button("Show Analytics", "📊")

        button_layout.addWidget(self.track_payment_button)
        button_layout.addWidget(self.calculate_payments_button)
        button_layout.addWidget(self.show_analytics_button)

        # Add everything to the main group layout
        group_layout.addLayout(info_grid_layout)
        group_layout.addSpacing(15)
        group_layout.addLayout(button_layout)

        group.setLayout(group_layout)
        return group


    def create_status_section(self):
        group = QGroupBox()
        layout = QVBoxLayout()

        # Drag-and-drop Instructions
        self.status_label = QLabel("Drag and drop a file here or click Upload")
        self.status_label.setObjectName("status_label")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("""
            QLabel {
                color: #787774;
                padding: 65px; /* Large padding for the drag-and-drop area */
                border: 1px dashed #cccccc;
                border-radius: 10px;
                background-color: #f9f9f9;
            }
        """)

        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)

        layout.addWidget(self.status_label)
        layout.addWidget(self.progress_bar)

        group.setLayout(layout)
        return group

    def create_button(self, text, emoji):
        button = QPushButton(f"{emoji} {text}")
        button.setCursor(Qt.PointingHandCursor)
        return button

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if len(urls) > 0 and urls[0].toLocalFile().endswith('.xls'):
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        dropped_file = event.mimeData().urls()[0].toLocalFile()
        self.file_path = dropped_file
        self.status_label.setText(f"Selected file: {os.path.basename(dropped_file)}")
        self.convert_button.setEnabled(True)

    def show_analytics(self):
        if self.final_xlsx_directory:
            file_path = os.path.join(self.final_xlsx_directory, 'final.xlsx')
            if not os.path.exists(file_path):
                QMessageBox.warning(self, "File Not Found", "The final.xlsx file could not be found in the selected directory.")
                return
            self.analytics_tab = AnalyticsTab(self.final_xlsx_directory)
            self.analytics_tab.show()
        else:
            self.status_label.setText("Please select a directory first")

    def calculate_and_update_payments(self):
        try:
            # Determine the file path
            file_path = os.path.join(self.final_xlsx_directory, 'final.xlsx') if self.final_xlsx_directory else \
                os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'final.xlsx')

            # Check if the file exists
            if not os.path.exists(file_path):
                raise FileNotFoundError("The final.xlsx file could not be found.")

            # Proceed with payment calculation
            total_cod_amount, pending_payment, delivered_pending = calculate_payments(file_path)

            if total_cod_amount is None or pending_payment is None:
                raise ValueError("Invalid payment values")

            # Update labels with calculated values
            self.pending_payment_label.setText(f"Rs. {round(pending_payment, -2):,.2f}")
            self.total_payment_label.setText(f"Rs. {round(total_cod_amount, -2):,.2f}")
            self.delivered_pending_label.setText(f"Rs. {round(delivered_pending, -2):,.2f}")

        except (FileNotFoundError, ValueError) as e:
            QMessageBox.warning(self, "Error", str(e))
            self.pending_payment_label.setText("Pending Payment: Error")
            self.total_payment_label.setText("Total Payment: Error")

        except Exception as e:
            QMessageBox.warning(self, "Error", f"An unexpected error occurred: {e}")
            self.pending_payment_label.setText("Pending Payment: Error")
            self.total_payment_label.setText("Total Payment: Error")


    def upload_file(self):
        if self.final_xlsx_directory:
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            file_path, _ = QFileDialog.getOpenFileName(self, "Select an XLS file", self.final_xlsx_directory, "Excel Files (*.xls);;All Files (*)", options=options)
            if file_path:
                self.file_path = file_path
                self.status_label.setText(f"Selected file: {os.path.basename(file_path)}")
                self.convert_button.setEnabled(True)
        else:
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            file_path, _ = QFileDialog.getOpenFileName(self, "Select an XLS file", os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'), "Excel Files (*.xls);;All Files (*)", options=options)
            if file_path:
                self.file_path = file_path
                self.status_label.setText(f"Selected file: {os.path.basename(file_path)}")
                self.convert_button.setEnabled(True)
            else:
                QMessageBox.warning(self, "Error", "Please select a file.")


    def convert_file(self):
        if not is_connected():
            QMessageBox.warning(self, "Network Error", "No internet connection. Please check your network and try again.")
            return
        
        if self.file_path:
            if not self.file_path.endswith('.xls'):
                QMessageBox.warning(self, "Invalid File", "Please select an .xls file.")
                return
            
            self.convert_button.setEnabled(False)
            self.track_button.setEnabled(False)
            self.track_payment_button.setEnabled(False)

            api_key = self.api_key_input.text()
            api_password = self.api_password_input.text()
            save_config(api_key, api_password, self.final_xlsx_directory)

            api = LeopardCourierAPI(api_key, api_password)
            result_message, html_file_path = rename_file_extension(self.file_path)

            self.file_path = None

            if html_file_path:
                try:
                    df = extract_data_from_html(html_file_path)
                except Exception as e:
                    print(f"Error occurred: {e}")
                    return

                directory = os.path.dirname(html_file_path)
                output_file_path = os.path.join(directory, 'temporary.xlsx')

                try:
                    save_data_to_excel(df, output_file_path)
                except Exception as e:
                    print(f"Error occurred: {e}")
                    return

                if self.final_xlsx_directory:
                    final_file_path = os.path.join(self.final_xlsx_directory, 'final.xlsx')
                else:
                    final_file_path = os.path.join(directory, 'final.xlsx')

                if append_to_final(output_file_path, final_file_path):
                    add_columns(final_file_path)
                    sort_by_booking_date(final_file_path)
                    self.worker_thread = WorkerThread(api, final_file_path, mode= "tracking")

                    self.worker_thread.progress.connect(self.update_progress)
                    self.worker_thread.result.connect(self.tracking_completed)
                    self.worker_thread.error.connect(self.tracking_failed)
                    self.worker_thread.finished.connect(self.cleanup_thread)

                    self.worker_thread.start()

                    self.status_label.setText("Tracking packets in progress...")


                else:
                    delete_temporary_files(html_file_path, output_file_path)
            else:
                QMessageBox.warning(self, "Error", result_message)

            self.status_label.setText("Drag and drop a file or click 'Upload'")
            self.convert_button.setEnabled(False)
        else:
            QMessageBox.warning(self, "No File", "Please upload a file first.")

    def track_existing_payments(self):
        if not is_connected():
            QMessageBox.warning(self, "Network Error", "No internet connection. Please check your network and try again.")
            return

        try:
            api_key = self.api_key_input.text()
            api_password = self.api_password_input.text()

            save_config(api_key, api_password, self.final_xlsx_directory)

            api = LeopardCourierAPI(api_key, api_password)

            if not self.final_xlsx_directory:
                QMessageBox.warning(self, "Error", "Please select a directory first.")
                return

            final_file_path = os.path.join(self.final_xlsx_directory, 'final.xlsx')
            if not os.path.exists(final_file_path):
                QMessageBox.warning(self, "Error", "The file 'final.xlsx' does not exist in the selected directory.")
                return

            # Initialize WorkerThread
            self.worker_thread = WorkerThread(api, final_file_path, mode="payment")

            # Connect signals
            self.worker_thread.progress.connect(self.update_progress)
            self.worker_thread.result.connect(self.tracking_completed)
            self.worker_thread.error.connect(self.tracking_failed)
            self.worker_thread.finished.connect(self.cleanup_thread)

            # Disable UI elements during tracking
            self.track_payment_button.setEnabled(False)
            self.track_button.setEnabled(False)
            self.convert_button.setEnabled(False)

            # Start the worker thread
            self.worker_thread.start()
            self.status_label.setText("Tracking payments in progress...")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred: {e}")

    def track_existing_parcels(self):
        if not is_connected():
            QMessageBox.warning(self, "Network Error", "No internet connection. Please check your network and try again.")
            return

        try:
            api_key = self.api_key_input.text()
            api_password = self.api_password_input.text()

            save_config(api_key, api_password, self.final_xlsx_directory)

            api = LeopardCourierAPI(api_key, api_password)

            if not self.final_xlsx_directory:
                QMessageBox.warning(self, "Error", "Please select a directory first.")
                return

            final_file_path = os.path.join(self.final_xlsx_directory, 'final.xlsx')
            if not os.path.exists(final_file_path):
                QMessageBox.warning(self, "Error", "The file 'final.xlsx' does not exist in the selected directory.")
                return

            # Initialize WorkerThread
            self.worker_thread = WorkerThread(api, final_file_path, mode="tracking")

            # Connect signals
            self.worker_thread.progress.connect(self.update_progress)
            self.worker_thread.result.connect(self.tracking_completed)
            self.worker_thread.error.connect(self.tracking_failed)
            self.worker_thread.finished.connect(self.cleanup_thread)

            # Disable UI elements during tracking
            self.convert_button.setEnabled(False)
            self.track_payment_button.setEnabled(False)
            self.track_button.setEnabled(False)

            # Start the worker thread
            self.worker_thread.start()
            self.status_label.setText("Tracking existing parcels in progress...")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred: {e}")

    def update_progress(self, value):
        self.progress_bar.setValue(value)
        if value == 100:
            self.upload_button.setEnabled(True)
            self.setAcceptDrops(True)

    def tracking_completed(self):
        self.convert_button.setEnabled(False if not self.file_path else True)
        self.track_button.setEnabled(True)
        self.track_payment_button.setEnabled(True)
        self.upload_button.setText("📄 Upload LoadSheet")  # Reset text with icon
        self.upload_button.setEnabled(True)
        self.setAcceptDrops(True)

        # Post-processing on Excel (optional)
        if self.final_xlsx_directory:
            try:
                customize_excel(self, self.final_xlsx_directory)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to customize Excel: {e}")

    def tracking_failed(self):
        QMessageBox.warning(self, "Error", "An error occurred while tracking parcels.")
        self.convert_button.setEnabled(False if not self.file_path else True)
        self.track_button.setEnabled(True)
        self.track_payment_button.setEnabled(True)
        self.upload_button.setText("📄 Upload LoadSheet")  # Reset text with icon
        self.upload_button.setEnabled(True)
        self.setAcceptDrops(True)


    def cleanup_thread(self):
        if self.worker_thread:
            self.worker_thread.wait()
            self.worker_thread = None

        # Re-enable UI elements
        self.convert_button.setEnabled(False if not self.file_path else True)
        self.track_button.setEnabled(True)
        self.track_payment_button.setEnabled(True)
        self.upload_button.setText("📄 Upload LoadSheet")  # Reset text with icon
        self.upload_button.setEnabled(True)
        self.setAcceptDrops(True)
        self.status_label.setText("Ready for the next operation.")


    def select_directory(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        directory = QFileDialog.getExistingDirectory(self, "Select Directory", "", options=options)
        if directory:
            self.final_xlsx_directory = directory
            # Update the QLabel text
            self.directory_label.setText(f"📂 Selected Directory: {directory}")
            # Save the new directory to the configuration
            save_config(self.api_key_input.text(), self.api_password_input.text(), directory)
        else:
            QMessageBox.warning(self, "Error", "Please select a directory.")


    def open_final_excel(self):
        file_path = os.path.join(self.final_xlsx_directory, 'final.xlsx') if self.final_xlsx_directory else \
            os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'final.xlsx')
        file_path = os.path.normpath(file_path)
        open_excel_file(file_path)
    
    def update_api_strength(self):
        def check_api_strength():
            try:
                # Ensure API credentials are configured
                if not self.api_key or not self.api_password:
                    self.api_strength_label.setText("❌ API Not Configured")
                    return

                # Locate final.xlsx file and extract the last tracking number
                final_file_path = os.path.join(self.final_xlsx_directory, 'final.xlsx') if self.final_xlsx_directory else \
                    os.path.join(os.environ['USERPROFILE'], 'Desktop', 'final.xlsx')

                if not os.path.exists(final_file_path):
                    self.api_strength_label.setText("❌ final.xlsx Not Found")
                    return

                try:
                    df = pd.read_excel(final_file_path)
                    if 'CN #' not in df.columns or df.empty:
                        self.api_strength_label.setText("❌ No Tracking Data Available")
                        return
                    last_tracking_number = df['CN #'].dropna().iloc[-1]
                except Exception:
                    self.api_strength_label.setText("❌ Error Reading final.xlsx")
                    return

                # API instance
                api = LeopardCourierAPI(self.api_key, self.api_password)

                # Measure response time (3 attempts, calculate average)
                total_time = 0
                attempts = 2
                for _ in range(attempts):
                    try:
                        total_time += api.check_api_strength(last_tracking_number)
                    except Exception:
                        total_time += 15  # Assign max timeout for unresponsive API

                avg_response_time = total_time / attempts
                avg_response_time += 2;

                # Determine speed classification
                if avg_response_time > 10:
                    classification = "Bad 🟥📶"
                elif avg_response_time > 7:
                    classification = "Medium 🟧📶"
                elif avg_response_time > 3:
                    classification = "Good 🟨📶"
                else:
                    classification = "Excellent 🟩📶"

                # Update the label
                self.api_strength_label.setText(f"API: {classification} ({avg_response_time:.2f}s)")

            except Exception:
                self.api_strength_label.setText("❌ API Check Failed")

        # Run the API strength check in a separate thread to prevent blocking the UI
        thread = threading.Thread(target=check_api_strength)
        thread.daemon = True  # Correct way to set the thread as a daemon
        thread.start()


def main():
    app = QApplication(sys.argv)
    window = FileConverterApp()
    window.resize(500,600)
    window.show()
    sys.exit(app.exec_())

if __name__:
    main()
