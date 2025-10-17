import cv2
from pyzbar.pyzbar import decode, ZBarSymbol
import openpyxl
from openpyxl import Workbook
import time
import numpy as np
import os

# --- Configuration ---
# The name of the Excel file where data will be stored.
EXCEL_FILENAME = 'scanned_entries.xlsx'

# --- Camera Setup ---
# Option 1: Laptop Webcam -> Set to 0
# Option 2: Phone Link App -> Set to 1 (or 2, 3, etc. if 1 doesn't work)
# Option 3: IP Webcam App -> Set to 'http://YOUR_PHONE_IP:8080/video'
CAMERA_SOURCE = 0 # <-- CHANGE THIS if using an external camera

# --- Data Parsing ---
def parse_qr_data(raw_data):
    """Parses the multi-line QR code string into a dictionary."""
    data = {}
    key_map = {
        'Name': 'Name',
        'Email': 'Email',
        'Phone': 'Phone',
        'College': 'College',
        'Ticket': 'Ticket',
        'Ticket No': 'Ticket'
    }
    lines = raw_data.strip().split('\n')
    for line in lines:
        if ':' in line:
            key, value = line.split(':', 1)
            key = key.strip()
            value = value.strip()
            standard_key = key_map.get(key)
            if standard_key:
                data[standard_key] = value
    return data

# --- Excel File Setup ---
def setup_excel_file():
    """Checks if the Excel file exists. If not, creates it with a header row."""
    if not os.path.exists(EXCEL_FILENAME):
        print(f"'{EXCEL_FILENAME}' not found. Creating a new file...")
        workbook = Workbook()
        sheet = workbook.active
        headers = ['Name', 'Email', 'Phone', 'College', 'Ticket', 'Scan Timestamp']
        sheet.append(headers)
        workbook.save(EXCEL_FILENAME)
        print("Excel file created successfully.")

# --- Main Application Logic ---
def main():
    """Main function to run the QR code scanner."""
    setup_excel_file()

    print(f"Attempting to open camera source: {CAMERA_SOURCE}")
    cap = cv2.VideoCapture(CAMERA_SOURCE)
    if not cap.isOpened():
        print("Error: Could not open video stream.")
        print("Check if the camera is connected and the CAMERA_SOURCE value is correct.")
        return

    print(f"Starting QR code scanner. Data will be saved to '{EXCEL_FILENAME}' in real-time.")
    print("IMPORTANT: Keep the Excel file closed to avoid permission errors.")
    print("Press 'q' to quit.")
    last_scan_time = 0
    scan_cooldown = 2 # 2-second cooldown between scans to prevent accidental multiple entries

    while True:
        success, frame = cap.read()
        if not success:
            print("Failed to capture frame.")
            break

        detections = decode(frame, symbols=[ZBarSymbol.QRCODE])

        if not detections:
            cv2.putText(frame, "Point camera at a QR Code", (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (3, 252, 23), 2)
        else:
            for barcode in detections:
                current_time = time.time()
                # Apply cooldown to prevent the same code from being scanned multiple times in rapid succession
                if current_time - last_scan_time < scan_cooldown:
                    continue

                last_scan_time = current_time
                qr_data_raw = barcode.data.decode('utf-8')
                parsed_data = parse_qr_data(qr_data_raw)
                pts = np.array(barcode.polygon, np.int32).reshape((-1, 1, 2))

                # Since duplicates are allowed, we just process and save every scan.
                print(f"SUCCESS: Adding to Excel file...")
                cv2.polylines(frame, [pts], True, (0, 255, 0), 3)
                cv2.putText(frame, "SAVED", (barcode.rect.left, barcode.rect.top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 0), 2)

                try:
                    # --- REAL-TIME SAVE LOGIC ---
                    # Load the workbook, append the row, and save it for each scan.
                    workbook = openpyxl.load_workbook(EXCEL_FILENAME)
                    sheet = workbook.active

                    name = parsed_data.get('Name', '')
                    email = parsed_data.get('Email', '')
                    phone = parsed_data.get('Phone', '')
                    college = parsed_data.get('College', '')
                    ticket_id = parsed_data.get('Ticket', 'N/A') # Use N/A if no ticket is found
                    timestamp = time.strftime('%Y-%m-%d %H:%M:%S')

                    row_to_add = [name, email, phone, college, ticket_id, timestamp]
                    sheet.append(row_to_add)
                    workbook.save(EXCEL_FILENAME)
                    print(f"Entry for '{name}' added to Excel.")

                except Exception as e:
                    # This error will happen if the excel file is open.
                    print(f"ERROR: Could not write to Excel file. Reason: {e}")
                    cv2.putText(frame, "FILE ERROR", (barcode.rect.left, barcode.rect.top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 0, 255), 2)

        cv2.imshow('QR Code Scanner', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()
    print("Scanner stopped.")
    print(f"All scanned data has been saved in '{EXCEL_FILENAME}'.")

if __name__ == '__main__':
    main()

