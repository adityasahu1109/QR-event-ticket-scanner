# **QR Code Scanner to Google Sheets**

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
  sudo apt-get update  
  sudo apt-get install libzbar0

### **Step 2: Set Up a Python Environment**

It is highly recommended to use a virtual environment to manage project dependencies.

1. **Create a virtual environment**:  
   python \-m venv venv

2. **Activate it**:  
   * On Windows:  
     venv\\Scripts\\activate

   * On macOS/Linux:  
     source venv/bin/activate

3. **Install the required Python libraries**:  
   pip install opencv-python pyzbar gspread oauth2client numpy

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
   python qr\_scanner.py

3. A window showing your webcam feed will appear. Point the camera at a QR code.  
4. To stop the scanner, make sure the webcam window is active and press the **'q'** key.