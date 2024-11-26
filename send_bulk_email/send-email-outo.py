import smtplib
from tkinter import Tk, Label, Entry, Text, Button, filedialog, messagebox, Frame
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import time
from pathlib import Path
import threading
import schedule
from datetime import datetime

# SMTP Configuration
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587

# Global Variables for Excel File
excel_file_path = None

# Select Attachment
def select_attachment():
    filepath = filedialog.askopenfilename()
    if filepath:
        attachment_label.config(text=filepath)

# Select Excel File
def select_excel_file():
    global excel_file_path
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if filepath:
        excel_file_path = filepath
        excel_file_label.config(text=f"File Selected: {filepath}")

# Bulk Email Sending Function
def send_bulk_emails():
    sender_email = sender_entry.get()
    sender_password = password_entry.get()

    if not excel_file_path:
        messagebox.showerror("Error", "No Excel file selected!")
        return

    try:
        df = pd.read_excel(excel_file_path)

        required_columns = ['Name', 'Email', 'DueDate', 'invioce_no', 'amount']
        if not all(col in df.columns for col in required_columns):
            messagebox.showerror("Error", f"Excel file must contain columns: {', '.join(required_columns)}")
            return

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(sender_email, sender_password)

            for _, row in df.iterrows():
                recipient_name = row['Name']
                recipient_email = row['Email']
                due_date = row['DueDate']
                invoice_no = row['invioce_no']
                amount = row['amount']
                custom_message = row.get('Message', '')

                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient_email
                msg['Subject'] = f"Information about price trips {recipient_name}"

                body = f"""
                Dear {recipient_name},\n\n
                I just wanted to drop your price {amount} USD in respect of our invoice {invoice_no} is due 
                for payment on {due_date}. I would be really grateful if you could confirm that everything is on track 
                for payment.\n\n
                {custom_message} \n\n
                Best regards,\n
                Your Name
                """
                msg.attach(MIMEText(body, 'plain'))

                try:
                    server.send_message(msg)
                    print(f"Email sent to {recipient_name} ({recipient_email})")
                except Exception as e:
                    print(f"Failed to send email to {recipient_name} ({recipient_email}): {e}")

        messagebox.showinfo("Success", "Bulk emails sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send bulk emails: {e}")

# Function to run the schedule in a separate thread
def run_schedule():
    while True:
        schedule.run_pending()
        time.sleep(1)
        
# Create GUI
root = Tk()
root.title("Email Automation System")

# Get screen width and height
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Set window size and center it
window_width = 800
window_height = 600
x_coordinate = (screen_width - window_width) // 2
y_coordinate = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

# Create main frame for central alignment
main_frame = Frame(root, padx=20, pady=20, bg="lightblue")
main_frame.place(relx=0.5, rely=0.5, anchor="center")

# Title Label
title_font = ("Arial", 24, "bold")
font_config = ("Arial", 13)

Label(main_frame, text="Email Sending Automation System", font=title_font, bg="lightblue").grid(row=0, column=0, columnspan=3, pady=20)

# Sender Email
Label(main_frame, text="Sender Email:", font=font_config, bg="lightblue").grid(row=1, column=0, padx=10, pady=5, sticky="w")
sender_entry = Entry(main_frame, width=40, font=font_config, bd=1, relief="solid")
sender_entry.grid(row=1, column=1, padx=10, pady=5)

# Sender Password
Label(main_frame, text="Password:", font=font_config, bg="lightblue").grid(row=2, column=0, padx=10, pady=5, sticky="w")
password_entry = Entry(main_frame, width=40, show="*", font=font_config, bd=1, relief="solid")
password_entry.grid(row=2, column=1, padx=10, pady=5)

# Subject
Label(main_frame, text="Subject:", font=font_config, bg="lightblue").grid(row=3, column=0, padx=10, pady=5, sticky="w")
subject_entry = Entry(main_frame, width=40, font=font_config, bd=1, relief="solid")
subject_entry.grid(row=3, column=1, padx=10, pady=5)

# Email Body
# Label(main_frame, text="Email Body (HTML):", font=font_config, bg="lightblue").grid(row=4, column=0, padx=10, pady=5, sticky="w")
# email_body = Text(main_frame, width=60, height=10, font=font_config, bd=2, relief="solid")
# email_body.grid(row=4, column=1, padx=10, pady=5)

# Attachment
Label(main_frame, text="Attachment:", font=font_config, bg="lightblue").grid(row=5, column=0, padx=10, pady=5, sticky="w")
attachment_label = Label(main_frame, text="No file selected", width=40, anchor="w", font=font_config, bg="lightblue")
attachment_label.grid(row=5, column=1, padx=10, pady=5)
Button(main_frame, text="Browse", command=select_attachment, font=font_config, bg="#4CAF50", fg="white").grid(row=5, column=2, padx=10, pady=5)

# Excel File Selection
Label(main_frame, text="Excel File:", font=font_config, bg="lightblue").grid(row=6, column=0, padx=10, pady=5, sticky="w")
excel_file_label = Label(main_frame, text="No file selected", width=40, anchor="w", font=font_config, bg="lightblue")
excel_file_label.grid(row=6, column=1, padx=10, pady=5)
Button(main_frame, text="Browse", command=select_excel_file, font=font_config, bg="#4CAF50", fg="white").grid(row=6, column=2, padx=10, pady=5)

# Buttons for Sending Emails
Button(main_frame, text="Send Emails", command=send_bulk_emails, bg="#007BFF", fg="white", font=font_config, width=20).grid(row=9, column=0, pady=10, columnspan=3)

root.mainloop()
