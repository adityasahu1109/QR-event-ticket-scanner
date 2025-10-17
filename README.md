# **📷 Real-Time QR Code Scanner**

This Python script uses your webcam to scan QR codes in real-time and logs the scanned data directly into a Google Sheet. It includes a feature to prevent duplicate entries by checking against already scanned data.

## **Features**

* **Real-Time Scanning**: Uses OpenCV to capture video from your webcam.  
* **QR Code Detection**: Employs the pyzbar library to accurately detect and decode QR codes.  
* **Google Sheets Integration**: Connects to the Google Sheets API using gspread to log data.  
* **Duplicate Prevention**: Keeps track of scanned QR codes to prevent logging the same code twice.  
* **Visual Feedback**: Displays a live camera feed, drawing boxes around detected QR codes (green for success, red for duplicates).

## **Prerequisites**

* Python 3.7+  
* A webcam connected to your computer.  
* A Google account.

## **Setup Instructions**

Follow these steps carefully to get the scanner up and running.
### **Step 1: Install System Dependencies (ZBar)**

The pyzbar library requires a system-level dependency called ZBar. You must install it first.

* **Windows**:  
  1. The easiest way is often to use a pre-compiled wheel. When you install pyzbar with pip, it should handle this. If it fails, you may need to find specific ZBar DLLs online. For most users, the pip install in the next step will suffice.  
* **macOS** (using [Homebrew](https://brew.sh/)):  
  brew install zbar

* **Linux** (Debian/Ubuntu):  
  ```
  sudo apt-get update  
  sudo apt-get install libzbar0
  ```
### **Step 2: Set Up a Python Environment**

It is highly recommended to use a virtual environment to manage project dependencies.

1. **Create a virtual environment**:  
   ```
   python \-m venv venv
    ```
2. **Activate it**:  
   * On Windows:
     ```
     venv\\Scripts\\activate
     ```

   * On macOS/Linux:
     ```  
     source venv/bin/activate
      ```
3. **Install the required Python libraries**:  
   ```
   pip install opencv-python pyzbar gspread oauth2client numpy
   ```
  
### **Step 3: Configure Google Cloud & Service Account**

This is the most critical part. This process allows the script to securely access your Google Sheet without you having to log in manually.

1. **Go to the [Google Cloud Console](https://console.cloud.google.com/)**.  
2. **Create a new project** (or select an existing one).  
3. **Enable APIs**:  
   * In the search bar at the top, find and **enable** the **Google Drive API**.  
   * Search again and **enable** the **Google Sheets API**.  
   * *You must enable both APIs for the script to work.*  
4. **Create a Service Account**:  
   * In the top search bar, navigate to "Service Accounts".  
   * Click **"+ CREATE SERVICE ACCOUNT"**.  
   * Give it a name (e.g., "qr-scanner-bot") and a description. Click **"CREATE AND CONTINUE"**.  
   * In the "Grant this service account access to project" step, select the role **"Editor"**. Click **"CONTINUE"**, then **"DONE"**.  
5. **Generate JSON Key**:  
   * You will be returned to the service accounts list. Find the account you just created and click on the email address.  
   * Go to the **"KEYS"** tab.  
   * Click **"ADD KEY"** \-\> **"Create new key"**.  
   * Select **JSON** as the key type and click **"CREATE"**.  
   * A JSON file will be downloaded to your computer.  
6. **Final Step for Key**:  
   * **Rename the downloaded file to service\_account.json**.  
   * **Move this file into the same folder as your qr\_scanner.py script.**

### **Step 4: Configure Your Google Sheet**

1. **Create a new Google Sheet**.  
2. **Share the sheet**:  
   * Open your service\_account.json file in a text editor. Find the value associated with the "client\_email" key (it will look like an email address).  
   * In your Google Sheet, click the **"Share"** button.  
   * Paste the client\_email into the sharing dialog and give it **"Editor"** permissions.  
3. **Set up the headers**:  
   * In the first row of your worksheet, set the headers as follows:  
     * Cell A1: Scanned Data  
     * Cell B1: Scan Timestamp

### **Step 5: Configure the Python Script**

1. Open qr\_scanner.py in an editor.  
2. Find the line SHEET\_NAME \= 'YourGoogleSheetName' and change 'YourGoogleSheetName' to the exact name of your Google Sheet file.  
3. If you are not using the default worksheet named "Sheet1", change WORKSHEET\_NAME \= 'Sheet1' accordingly.

## **How to Run the Scanner**

1. Make sure your virtual environment is activated.  
2. Run the script from your terminal:  
  ```
   python qr\_scanner.py
  ```
4. A window showing your webcam feed will appear. Point the camera at a QR code.  
5. To stop the scanner, make sure the webcam window is active and press the **'q'** key.


## **📄 Version 2: Microsoft Excel (Offline)**

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
  ```
 python \-m venv venv
  ```
3. **Activate the environment**:  
   * **On Windows**:  
    ```   venv\\Scripts\\activate ```

   * **On macOS / Linux**:  
     ``` source venv/bin/activate ```

You will know it's active when you see (venv) at the beginning of your command prompt line.

## **Step 3: Install the Required Python Libraries**

With your virtual environment active, install all the necessary libraries with a single command:

``` pip install opencv-python pyzbar numpy openpyxl  ```

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
