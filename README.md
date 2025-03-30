# Leopard Courier Tracker – CaseDrip COD Analytics App

A simple but powerful desktop app built with Python and PyQt5 that connects to the Leopard Courier Pakistan API.

It helps you manage parcel tracking, update delivery and payment statuses, convert daily loadsheets, and view parcel analytics — all in one place.

---

## Key Features

- Convert `.xls` loadsheets into organized Excel files
- Track delivery status of parcels using Leopard’s API
- Monitor COD payment status (paid, pending, etc.)
- View live parcel stats with pie chart analytics
- Auto-highlight important rows (e.g. high COD, pending payments)
- Check API response speed (Fast/Slow)

---

## Tech Stack

Built with:

- Python 3.8+
- PyQt5 (UI)
- Pandas & OpenPyXL (Excel handling)
- Matplotlib (Charts)
- BeautifulSoup (HTML parsing)
- Requests (API calls)

---

## Getting Started

```bash
git clone https://github.com/muhammadhaxcan/leopard-courier-tracker.git
cd leopard-courier-tracker

python -m venv venv
venv\Scripts\activate

pip install pyqt5 pandas openpyxl matplotlib beautifulsoup4 requests

{
  "api_key": "your_leopard_api_key",
  "api_password": "your_leopard_api_password",
  "directory": "C:/Path/Where/You/Want/To/Save/Excel"
}


