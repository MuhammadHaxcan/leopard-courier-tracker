from PyQt5.QtWidgets import QMessageBox
import os
import socket
from bs4 import BeautifulSoup
import pandas as pd

def rename_file_extension(file_path):
    file_name, current_extension = os.path.splitext(file_path)
    current_extension = current_extension[1:]

    if current_extension == "html":
        return f"The file {file_path} already has the extension html.", file_path

    new_file_path = f"{file_name}.html"

    try:
        os.rename(file_path, new_file_path)
        return f"Renamed {file_path} to {new_file_path}.", new_file_path
    except FileNotFoundError:
        QMessageBox.warning(None, "File Not Found", f"The file {file_path} does not exist.")
        return None, None
    except Exception as e:
        QMessageBox.warning(None, "Error", f"An error occurred: {e}")
        return None, None

def extract_data_from_html(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()

        soup = BeautifulSoup(content, 'html.parser')
        rows = soup.find_all('tr')

        headers = ["Sr.", "CN #", "Destination", "Shipper Name", "No. of pieces", "Consignee Name", "Order Id", "Weight", "COD Amount", "Remarks"]

        data = []
        for row in rows[1:]:
            cells = row.find_all('td')
            if len(cells) == 10:
                if cells[0].text.strip() == "Sr.":
                    continue
                row_data = [cell.text.strip() for cell in cells]
                data.append(row_data)

        df = pd.DataFrame(data, columns=headers)
        return df
    except Exception as e:
        raise Exception(f"Error extracting data from HTML: {e}")

def delete_temporary_files(html_file_path, output_file_path):
    os.remove(html_file_path)
    os.remove(output_file_path)
        

def is_connected(host="8.8.8.8", port=53, timeout=3):
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except (socket.timeout, socket.error):
        return False
    

def open_excel_file(file_path):
    if os.path.exists(file_path):
        try:
            os.startfile(file_path)  # Windows-specific method to open files
        except Exception as e:
            QMessageBox.critical(None, "Error", f"Failed to open the file: {e}")
    else:
        QMessageBox.warning(None, "File Not Found", "The specified Excel file does not exist.")

def get_final_file_path(self):
    return os.path.join(self.final_xlsx_directory or os.path.join(os.environ['USERPROFILE'], 'Desktop'), 'final.xlsx')

