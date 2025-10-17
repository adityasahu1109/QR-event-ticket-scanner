import cv2
from pyzbar.pyzbar import decode, ZBarSymbol
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import numpy as np


# Google Sheets API setup
SCOPE = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

# IMPORTANT: Paste your service account JSON key contents here.
SERVICE_ACCOUNT_INFO = {}



# IMPORTANT: The name of your Google Sheet file.
SHEET_NAME = 'attendance-flagship' # <-- CHANGE THIS
# IMPORTANT: The name of the worksheet (the tab at the bottom).
WORKSHEET_NAME = 'Sheet1' # <-- CHANGE THIS (if different)

# --- Camera Setup ---
# Option 1: Laptop Webcam -> Set to 0
# Option 2: Phone Link App -> Set to 1 (or 2, 3, etc. if 1 doesn't work)
# Option 3: IP Webcam App -> Set to 'http://YOUR_PHONE_IP:8080/video'
CAMERA_SOURCE = 0 # <-- CHANGE THIS TO 1 TO TRY PHONE LINK

# The column where the UNIQUE TICKET ID will be stored. This is used for duplicate checking.
# 'A'=1, 'B'=2, 'C'=3, 'D'=4, 'E'=5. We will put Ticket in column E.
DATA_COLUMN = 5

# --- Data Parsing ---
def parse_qr_data(raw_data):
    """Parses the multi-line QR code string into a dictionary."""
    data = {}
    # Define the keys we are looking for to handle variations in naming (e.g., 'Ticket' vs 'Ticket No')
    key_map = {
        'Name': 'Name',
        'Email': 'Email',
        'Phone': 'Phone',
        'College': 'College',
        'Ticket': 'Ticket',
        'Ticket No': 'Ticket' # Map 'Ticket No' to 'Ticket'
    }
    lines = raw_data.strip().split('\n')
    for line in lines:
        if ':' in line:
            key, value = line.split(':', 1)
            key = key.strip()
            value = value.strip()
            # Use the key_map to standardize the keys
            standard_key = key_map.get(key)
            if standard_key:
                data[standard_key] = value
    return data

# --- Google Sheets Connection ---
def setup_google_sheets_client():
    """Initializes and returns the gspread client."""
    try:
        if SERVICE_ACCOUNT_INFO.get("project_id") == "your-project-id":
             print("ERROR: Paste your service account credentials into the SERVICE_ACCOUNT_INFO dictionary.")
             return None
        creds = ServiceAccountCredentials.from_json_keyfile_dict(SERVICE_ACCOUNT_INFO, SCOPE)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        print(f"An error occurred during Google Sheets authentication: {e}")
        return None

def get_worksheet(client):
    """Opens and returns the specified worksheet."""
    try:
        sheet = client.open(SHEET_NAME).worksheet(WORKSHEET_NAME)
        return sheet
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"ERROR: Spreadsheet '{SHEET_NAME}' not found.")
        return None
    except gspread.exceptions.WorksheetNotFound:
        print(f"ERROR: Worksheet '{WORKSHEET_NAME}' not found.")
        return None
    except Exception as e:
        print(f"An error occurred while opening the worksheet: {e}")
        return None

# --- Main Application Logic ---
def main():
    """Main function to run the QR code scanner."""
    client = setup_google_sheets_client()
    if not client: return
    worksheet = get_worksheet(client)
    if not worksheet: return

    print("Successfully connected to Google Sheets.")
    print("\nIMPORTANT: Make sure your sheet has the following headers in the first row:")
    print("--> Name | Email | Phone | College | Ticket | Scan Timestamp <--\n")
    
    print(f"Attempting to open camera source: {CAMERA_SOURCE}")
    cap = cv2.VideoCapture(CAMERA_SOURCE)
    if not cap.isOpened():
        print("Error: Could not open video stream.")
        print("Check if the camera is connected and the CAMERA_SOURCE value is correct.")
        return

    print("Starting QR code scanner... Press 'q' to quit.")
    last_scan_time = 0
    scan_cooldown = 2

    while True:
        success, frame = cap.read()
        if not success:
            print("Failed to capture frame.")
            break

        detections = decode(frame, symbols=[ZBarSymbol.QRCODE])

        if not detections:
            cv2.putText(frame, "Point camera at a QR Code", (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 0), 2)
        else:
            for barcode in detections:
                current_time = time.time()
                if current_time - last_scan_time < scan_cooldown: continue

                qr_data_raw = barcode.data.decode('utf-8')
                parsed_data = parse_qr_data(qr_data_raw)
                ticket_id = parsed_data.get('Ticket')
                pts = np.array(barcode.polygon, np.int32).reshape((-1, 1, 2))

                # If there's no ticket number, we can't check for duplicates. Skip this QR code.
                if not ticket_id:
                    print("Warning: Scanned QR code has invalid format or is missing 'Ticket'.")
                    cv2.polylines(frame, [pts], True, (0, 165, 255), 3) # Orange box for invalid format
                    cv2.putText(frame, "INVALID FORMAT", (barcode.rect.left, barcode.rect.top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 165, 255), 2)
                    continue

                last_scan_time = current_time
                
                # --- REAL-TIME DUPLICATE CHECK using Ticket ID ---
                print(f"Checking for Ticket ID '{ticket_id}' in the sheet...")
                try:
                    cell = worksheet.find(ticket_id, in_column=DATA_COLUMN)
                except gspread.exceptions.APIError as e:
                    print(f"ERROR: Google Sheets API error during check: {e}")
                    cv2.putText(frame, "API ERROR", (barcode.rect.left, barcode.rect.top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 0, 255), 2)
                    continue # Skip this frame and try again

                if cell is not None:
                    # If a cell is found, the data already exists. It's a duplicate.
                    print(f"DUPLICATE: Already in sheet at row {cell.row}.")
                    cv2.polylines(frame, [pts], True, (0, 0, 255), 3)
                    cv2.putText(frame, "DUPLICATE", (barcode.rect.left, barcode.rect.top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 0, 255), 2)
                else:
                    # If no cell is found, this is a new, unique entry.
                    print(f"SUCCESS: Adding to Google Sheets...")
                    cv2.polylines(frame, [pts], True, (0, 255, 0), 3)
                    cv2.putText(frame, "SUCCESS", (barcode.rect.left, barcode.rect.top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 0), 2)
                    try:
                        # Get data from parsed dictionary, with fallbacks for missing fields
                        name = parsed_data.get('Name', '')
                        email = parsed_data.get('Email', '')
                        phone = parsed_data.get('Phone', '')
                        college = parsed_data.get('College', '')
                        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')

                        # Prepare row in the correct order for the sheet
                        row_to_add = [name, email, phone, college, ticket_id, timestamp]
                        worksheet.append_row(row_to_add)
                        print("Successfully added to sheet.")
                    except Exception as e:
                        print(f"ERROR: Could not write to Google Sheet. Reason: {e}")
                        cv2.putText(frame, "SHEET ERROR", (barcode.rect.left, barcode.rect.top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 0, 255), 2)

        cv2.imshow('QR Code Scanner', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'): break

    cap.release()
    cv2.destroyAllWindows()
    print("Scanner stopped.")

if __name__ == '__main__':
    main()

