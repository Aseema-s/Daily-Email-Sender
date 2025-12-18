# Daily Excel Email Sender V2 📧

A robust desktop application built with Python (Tkinter) to automate the process of sending daily HTML reports based on Excel data. This tool allows you to schedule emails, automatically converts Excel tables to HTML, and manages recipient lists efficiently.

# 🚀 Features

User-Friendly GUI: Simple dashboard to manage credentials and files.

Multi-Provider Support: Built-in support for Gmail, Yahoo, and Outlook.

Excel to HTML Automation: Automatically reads an Excel file and converts the data into a beautifully formatted HTML table inside the email body.

Bulk Sending: Reads recipient emails from a dedicated Excel list.

Daily Scheduler: Set a specific time (24-hour format) for the email to be sent automatically every day.

Smart Settings: Automatically saves your credentials, file paths, and schedule time to settings.json. You only need to configure it once!

Live Logging: View real-time status, success counts, and error logs directly within the app.

# 🛠️ Installation & Requirements

Prerequisites

Python 3.x installed on your system.

Install Dependencies

Open your terminal or command prompt and install the required libraries:

pip install pandas schedule openpyxl


(Note: tkinter, smtplib, json, and threading are standard Python libraries included with Python.)

# ⚙️ Setup Guide

1. Prepare Your Files

Content File (.xlsx): The Excel sheet containing the data report you want to send.

Recipient File (.xlsx): An Excel sheet listing the email addresses. Ensure there is a column header named Email (or the emails are in the first column).

2. Get App Passwords (Security)

To use this app, you must use an App Password instead of your regular email password.

Gmail Users: Go to Google Account Security > 2-Step Verification > App Passwords. Create a new one and use that code.

Outlook/Yahoo: Check your security settings to generate an app-specific password.

# 📖 Usage Instructions

Step 1: Launch the Application

Run the script:

python daily_email_sender_v2.py


Step 2: Configuration (One-Time Setup)

Email Configuration: Select your provider, enter your email, and paste your App Password.

File Selection: Browse and select your Content Excel and Recipient Excel files.

Scheduler: Enter the time you want the email to go out (e.g., 09:00 for 9 AM).

Save: These settings are saved automatically when you start the scheduler. Next time you open the app, they will be pre-filled!

Step 3: Start the Automation

Click the "Start Scheduler" button. The button will turn Red, indicating the timer is active.

# ⚠️ IMPORTANT: How to keep it running

The "Watchman" Rule

Think of this app as a Watchman. For the watchman to open the gate (send the email) at the correct time, he must be present at the gate.

✅ MINIMIZE the App: You can click the _ (Minimize) button. The app will stay open in your taskbar and will send the email at the scheduled time.

❌ DO NOT CLOSE the App: If you click the X (Close) button, the app stops running. The scheduler will die, and the email will not be sent.

# Daily Routine:

Turn on your PC.

Open the App (Settings will auto-load).

Click "Start Scheduler".

Minimize the window and continue your work.

# 📂 Files in Repository

daily_email_sender_v2.py: The main source code.

settings.json: Stores your configuration locally (created automatically after first run).

README.md: Project documentation.

# 🐛 Troubleshooting

"Authentication Failed": You are likely using your normal password. Please generate an App Password from your email provider's security settings.

"File not found": Ensure the Excel files haven't been moved or renamed.

"Email not sent": Check if the computer was asleep or if the app was closed (X) instead of minimized (_) at the scheduled time.

# 📄 License

MIT License
