import requests
from datetime import datetime
from PyQt5.QtWidgets import QMessageBox

def show_error_message(error_message):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setText(error_message)
    msg.setWindowTitle("Error")
    msg.exec_()

class LeopardCourierAPI:
    def __init__(self, api_key, api_password):
        self.api_key = api_key
        self.api_password = api_password
        self.track_endpoint = 'https://merchantapi.leopardscourier.com/api/trackBookedPacket/format/json/'
        self.payment_endpoint = 'https://merchantapi.leopardscourier.com/api/getPaymentDetails/format/json/'
        self.session = requests.Session()
        self.session.headers.update({'Content-Type': 'application/json'})

    def create_payload(self, identifier, is_tracking=True):
        return {
            'api_key': self.api_key,
            'api_password': self.api_password,
            'track_numbers' if is_tracking else 'cn_numbers': identifier,
        }


    def send_request(self, payload):
        try:
            response = self.session.get(self.track_endpoint, json=payload)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            show_error_message(f'Error occurred: {e}')
    
    def track_booked_packet(self, track_number):
        try:
            payload = self.create_payload(track_number, is_tracking=True)
            data = self.send_request(payload)
            return self.parse_response(data)
        except Exception as e:
            raise Exception(f"Tracking packet failed: {e}")
        
    def track_payment_status(self, cn_numbers_str):
        try:
            payload = self.create_payload(cn_numbers_str, is_tracking=False)
            data = self.send_payment_request(payload)
            return self.parse_payment_response(data)
        except Exception as e:
            raise Exception(f"Error occurred: {e}")

    def send_payment_request(self, payload):
        try:
            response = self.session.get(self.payment_endpoint, params=payload)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            raise Exception(f'Error occurred while getting response: {e}')

    def parse_payment_response(self, data):
        results = {}
        if data['status'] == 1 and data['payment_list']:
            for payment_info in data['payment_list']:
                track_id = payment_info.get('booked_packet_cn')
                payment_status = payment_info.get('status', '-')
                payment_date = payment_info.get('invoice_cheque_date', '')

                if track_id:
                    results[track_id] = (payment_status, payment_date)

        return results

    def parse_response(self, data):
        try:
            if data['status'] == 1 and data['packet_list']:
                packet_info = data['packet_list'][0]
                booked_packet_status = packet_info['booked_packet_status']
                booking_date = packet_info['booking_date']
                tracking_details = packet_info.get('Tracking Detail', [])
                recent_status = tracking_details[-1]['Status'] if tracking_details else None
                
                return booked_packet_status, recent_status, booking_date
            else:
                return None, None, None
            
        except (KeyError, IndexError, ValueError) as e:
            raise Exception(f"Error parsing response: {e}")

    def check_api_strength(self, track_number):
        try:
            start = datetime.now()
            self.track_booked_packet(track_number)
            end = datetime.now()
            time_taken = (end - start).total_seconds()
            return time_taken
        except Exception as e:
            raise Exception(f"API strength check failed: {e}")
