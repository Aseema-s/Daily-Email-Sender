import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pandas as pd
import smtplib
import schedule
import time
import threading
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import os
import json

CONFIG_FILE = "settings.json"

class DailyEmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Daily Excel Email Sender V2")
        self.root.geometry("600x680")
        
        # State variables
        self.is_running = False
        self.scheduler_thread = None

        # --- UI ELEMENTS ---

        # 1. Credentials Section
        frame_creds = tk.LabelFrame(root, text="Email Configuration", padx=10, pady=10)
        frame_creds.pack(fill="x", padx=10, pady=5)

        # Provider Selection
        tk.Label(frame_creds, text="Email Service:").grid(row=0, column=0, sticky="w")
        self.combo_provider = ttk.Combobox(frame_creds, values=["Gmail", "Yahoo", "Outlook"], state="readonly", width=37)
        self.combo_provider.grid(row=0, column=1, padx=5, pady=2)
        self.combo_provider.current(0) # Default to Gmail

        tk.Label(frame_creds, text="Sender Email:").grid(row=1, column=0, sticky="w")
        self.entry_email = tk.Entry(frame_creds, width=40)
        self.entry_email.grid(row=1, column=1, padx=5, pady=2)

        tk.Label(frame_creds, text="App Password:").grid(row=2, column=0, sticky="w")
        self.entry_password = tk.Entry(frame_creds, width=40, show="*")
        self.entry_password.grid(row=2, column=1, padx=5, pady=2)

        # 2. File Selection Section
        frame_files = tk.LabelFrame(root, text="File Selection", padx=10, pady=10)
        frame_files.pack(fill="x", padx=10, pady=5)

        # Content File
        tk.Label(frame_files, text="Content Excel File:").grid(row=0, column=0, sticky="w")
        self.entry_content_file = tk.Entry(frame_files, width=40)
        self.entry_content_file.grid(row=0, column=1, padx=5, pady=2)
        tk.Button(frame_files, text="Browse", command=self.browse_content).grid(row=0, column=2, padx=5)

        # Recipient File
        tk.Label(frame_files, text="Recipient List Excel:").grid(row=1, column=0, sticky="w")
        self.entry_recipient_file = tk.Entry(frame_files, width=40)
        self.entry_recipient_file.grid(row=1, column=1, padx=5, pady=2)
        tk.Button(frame_files, text="Browse", command=self.browse_recipient).grid(row=1, column=2, padx=5)

        # 3. Schedule Section
        frame_sched = tk.LabelFrame(root, text="Scheduler Settings", padx=10, pady=10)
        frame_sched.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_sched, text="Daily Time (24h format):").pack(side="left")
        self.entry_time = tk.Entry(frame_sched, width=10)
        self.entry_time.insert(0, "09:00")
        self.entry_time.pack(side="left", padx=5)

        self.btn_start = tk.Button(frame_sched, text="Start Scheduler", command=self.toggle_scheduler, bg="green", fg="white")
        self.btn_start.pack(side="left", padx=20)
        
        self.btn_send_now = tk.Button(frame_sched, text="Send Now (Test)", command=self.run_job_manual)
        self.btn_send_now.pack(side="left")

        # 4. Logs
        frame_log = tk.LabelFrame(root, text="Activity Log", padx=10, pady=10)
        frame_log.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.log_area = scrolledtext.ScrolledText(frame_log, height=10)
        self.log_area.pack(fill="both", expand=True)

        # Load settings automatically when app starts
        self.load_settings()

    # --- SETTINGS MANAGEMENT ---
    def save_settings(self):
        data = {
            "provider": self.combo_provider.get(),
            "email": self.entry_email.get(),
            "password": self.entry_password.get(),
            "content_path": self.entry_content_file.get(),
            "recipient_path": self.entry_recipient_file.get(),
            "time": self.entry_time.get()
        }
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(data, f)
            self.log("Settings saved successfully.")
        except Exception as e:
            self.log(f"Could not save settings: {e}")

    def load_settings(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    data = json.load(f)
                    
                    # Load Provider
                    saved_provider = data.get("provider", "Gmail")
                    if saved_provider in self.combo_provider['values']:
                        self.combo_provider.set(saved_provider)
                    
                    self.entry_email.insert(0, data.get("email", ""))
                    self.entry_password.insert(0, data.get("password", ""))
                    self.entry_content_file.insert(0, data.get("content_path", ""))
                    self.entry_recipient_file.insert(0, data.get("recipient_path", ""))
                    
                    saved_time = data.get("time", "09:00")
                    self.entry_time.delete(0, tk.END)
                    self.entry_time.insert(0, saved_time)
                    
                self.log("Previous settings loaded.")
            except Exception as e:
                self.log(f"Could not load settings: {e}")

    # --- HELPER FUNCTIONS ---
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_area.see(tk.END)

    def browse_content(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.entry_content_file.delete(0, tk.END)
            self.entry_content_file.insert(0, filename)

    def browse_recipient(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.entry_recipient_file.delete(0, tk.END)
            self.entry_recipient_file.insert(0, filename)

    # --- EMAIL LOGIC ---
    def send_email_task(self):
        provider = self.combo_provider.get()
        sender_email = self.entry_email.get().strip()
        sender_password = self.entry_password.get().strip()
        content_path = self.entry_content_file.get().strip()
        recipient_path = self.entry_recipient_file.get().strip()
        
        if not all([sender_email, sender_password, content_path, recipient_path]):
            self.log("Error: Missing fields. Please fill all inputs.")
            return

        self.log(f"Starting email job via {provider}...")

        # Determine SMTP settings
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        
        if provider == "Yahoo":
            smtp_server = "smtp.mail.yahoo.com"
            smtp_port = 587
        elif provider == "Outlook":
            smtp_server = "smtp.office365.com"
            smtp_port = 587

        try:
            # 1. Read Content (Reads LIVE from the hard drive to get your latest saved changes)
            try:
                df_content = pd.read_excel(content_path)
                html_table = df_content.to_html(index=False, border=1, justify="center")
            except PermissionError:
                self.log("Error: Content Excel is open and locked. Please close the Excel window.")
                return
            except Exception as e:
                self.log(f"Error reading content file: {e}")
                return

            # 2. Read Recipients (Reads LIVE from the hard drive)
            try:
                df_recipients = pd.read_excel(recipient_path)
                if 'Email' in df_recipients.columns:
                    raw_emails = df_recipients['Email'].dropna().astype(str).tolist()
                else:
                    raw_emails = df_recipients.iloc[:, 0].dropna().astype(str).tolist()
                
                # Robust cleaning: handles empty spaces, multiple emails per cell, and drops invalid entries
                recipient_list = []
                for raw in raw_emails:
                    # Split by commas or semicolons if multiple emails are in a single cell
                    for email in raw.replace(';', ',').split(','):
                        clean_email = email.strip()
                        # Check if it looks like a valid email (contains '@' and no spaces)
                        if '@' in clean_email and ' ' not in clean_email:
                            recipient_list.append(clean_email)
                
                # Remove duplicates but preserve order
                recipient_list = list(dict.fromkeys(recipient_list))

            except PermissionError:
                self.log("Error: Recipient Excel is open and locked. Please close the Excel window.")
                return
            except Exception as e:
                self.log(f"Error reading recipient file: {e}")
                return

            if not recipient_list:
                self.log("No valid email addresses found in the recipient file.")
                return

            self.log(f"Found {len(recipient_list)} email(s). Preparing to send...")

            # 3. Connect and Send
            self.log(f"Connecting to {smtp_server}...")
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)

            success_count = 0
            for index, recipient in enumerate(recipient_list, 1):
                try:
                    self.log(f"Sending ({index}/{len(recipient_list)}) to: {recipient}")
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = recipient
                    msg['Subject'] = f"Daily Report - {datetime.now().strftime('%Y-%m-%d')}"

                    body = f"""
                    <html>
                    <body>
                        <h2>Daily Update</h2>
                        <p>Please find the daily data below:</p>
                        <br>
                        {html_table}
                        <br>
                        <p>Best Regards,<br>Automated System</p>
                    </body>
                    </html>
                    """
                    msg.attach(MIMEText(body, 'html'))
                    server.send_message(msg)
                    success_count += 1
                    
                    # Pause for 1.5 seconds between emails to prevent SMTP rate-limiting
                    time.sleep(1.5)
                except Exception as e:
                    self.log(f"Failed to send to {recipient}: {e}")

            server.quit()
            self.log(f"Job Finished. Sent {success_count} out of {len(recipient_list)} emails successfully.")

        except Exception as e:
            self.log(f"Critical Error: {e}")

    # --- SCHEDULING LOGIC ---
    def run_job_manual(self):
        # Save settings before manual run too
        self.save_settings()
        threading.Thread(target=self.send_email_task).start()

    def scheduler_loop(self):
        while self.is_running:
            schedule.run_pending()
            time.sleep(1)

    def toggle_scheduler(self):
        if not self.is_running:
            # Save settings when starting scheduler
            self.save_settings()
            
            # Start
            target_time = self.entry_time.get()
            self.log(f"Scheduler started. Waiting for {target_time}...")
            
            schedule.clear()
            schedule.every().day.at(target_time).do(self.send_email_task)
            
            self.is_running = True
            self.btn_start.config(text="Stop Scheduler", bg="red")
            
            # Start loop in thread
            self.scheduler_thread = threading.Thread(target=self.scheduler_loop, daemon=True)
            self.scheduler_thread.start()
        else:
            # Stop
            self.is_running = False
            self.log("Scheduler stopped.")
            self.btn_start.config(text="Start Scheduler", bg="green")

if __name__ == "__main__":
    root = tk.Tk()
    app = DailyEmailApp(root)
    root.mainloop()
