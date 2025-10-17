# **QR Code Scanner to Microsoft Excel \- Setup Guide**

This guide provides all the necessary instructions to set up and run the Python QR code scanner, which saves data directly to a local Microsoft Excel file.

## **Features**

* **Real-Time Scanning**: Uses your webcam or a connected phone camera to scan QR codes.  
* **Direct to Excel**: Instantly saves scanned data into a local .xlsx file.  
* **Data Parsing**: Automatically separates data fields (Name, Email, Phone, etc.) into different columns.  
* **Simple Setup**: Requires no cloud services or complex authentication.  
* **Visual Feedback**: Displays a live camera feed with a green box to confirm when a code has been successfully saved.

## **Prerequisites**

* Python 3.7 or newer installed on your computer.  
* A webcam, or a smartphone that can be used as a webcam (e.g., via the Phone Link app).

## **Step 1: Set Up Your Project Folder**

1. Create a new folder for your project on your computer.  
2. Save the qr\_scanner.py script inside this folder.

## **Step 2: Create and Activate a Python Virtual Environment**

Using a virtual environment is highly recommended to keep your project's dependencies separate from your system's Python installation.

1. **Open a terminal or command prompt** inside your new project folder.  
2. **Create the virtual environment** by running this command:  
   python \-m venv venv

3. **Activate the environment**:  
   * **On Windows**:  
     venv\\Scripts\\activate

   * **On macOS / Linux**:  
     source venv/bin/activate

You will know it's active when you see (venv) at the beginning of your command prompt line.

## **Step 3: Install the Required Python Libraries**

With your virtual environment active, install all the necessary libraries with a single command:

pip install opencv-python pyzbar numpy openpyxl

## **Step 4: Configure the Scanner (Optional)**

You may need to change the camera source if you are not using your computer's default webcam.

1. Open the qr\_scanner.py file in a text editor.  
2. Find the CAMERA\_SOURCE variable:  
   \# \--- Camera Setup \---  
   \# Option 1: Laptop Webcam \-\> Set to 0  
   \# Option 2: Phone Link App \-\> Set to 1 (or 2, 3, etc. if 1 doesn't work)  
   \# Option 3: IP Webcam App \-\> Set to 'http://YOUR\_PHONE\_IP:8080/video'  
   CAMERA\_SOURCE \= 0 \# \<-- CHANGE THIS if using an external camera

3. Change the value 0 to 1 or 2 if you are using a secondary camera like one connected through the Phone Link app.

## **Step 5: How to Run the Scanner**

1. Make sure your virtual environment is still active. If not, activate it again (see Step 2).  
2. Ensure you are in the correct project folder in your terminal.  
3. Run the script with the following command:  
   python qr\_scanner.py

4. The script will automatically create the Excel file (e.g., scanned\_entries.xlsx) if it doesn't exist.  
5. A window will appear showing your camera feed.  
6. Point your camera at a QR code. A green box will confirm that the data has been saved to the Excel file.  
7. To stop the scanner, click on the camera window and press the **'q'** key on your keyboard.

## **Known Limitations**

* **Excel File Must Be Closed**: You **must** keep the Excel file (scanned\_entries.xlsx) closed while the scanner is running. Because the script saves data in real-time, it needs exclusive write access to the file. If you have the file open in Excel, the script will be blocked, and you will see a Permission denied error in the terminal.