import smtplib
from tkinter import Tk, Label, Entry, Text, Button, filedialog, messagebox, Frame, Toplevel
from tkcalendar import Calendar
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

# Open Calendar to Select Date
def open_calendar():
    global selected_date

    def select_date():
        global selected_date
        selected_date = calendar.get_date()
        schedule_date_entry.delete(0, "end")
        schedule_date_entry.insert(0, selected_date)
        calendar_window.destroy()

    calendar_window = Toplevel(root)
    calendar_window.title("Select Date")
    calendar = Calendar(calendar_window, date_pattern="yyyy-mm-dd")
    calendar.pack(pady=20)

    Button(calendar_window, text="Select", command=select_date).pack(pady=30)

# Bulk Email Sending Function
def send_bulk_emails():
    sender_email = sender_entry.get()
    sender_password = password_entry.get()
    recipient_email = recipient_entry.get()  # Get recipient email

    if not excel_file_path:
        messagebox.showerror("Error", "No Excel file selected!")
        return

    if not recipient_email:  # Check if recipient email is empty
        messagebox.showerror("Error", "Recipient email is required!")
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
                due_date = row['DueDate']
                invoice_no = row['invioce_no']
                amount = row['amount']
                custom_message = row.get('Message', '')

                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient_email  # Send to the entered recipient email
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

# Schedule Email Function
def schedule_emails():
    schedule_date = schedule_date_entry.get()
    schedule_time = schedule_time_entry.get()
    schedule_datetime_str = f"{schedule_date} {schedule_time}"
    schedule_datetime = datetime.strptime(schedule_datetime_str, "%Y-%m-%d %H:%M")

    schedule.every().day.at(schedule_time).do(send_bulk_emails)
    messagebox.showinfo("Scheduled", f"Bulk emails scheduled for {schedule_datetime}")

    threading.Thread(target=run_schedule, daemon=True).start()

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
password_entry.grid(row=2, column=1, padx=15, pady=13)

# Recipient Email
Label(main_frame, text="Recipient Email:", font=font_config, bg="lightblue").grid(row=3, column=0, padx=10, pady=5, sticky="w")
recipient_entry = Entry(main_frame, width=40, font=font_config, bd=1, relief="solid")
recipient_entry.grid(row=3, column=1, padx=10, pady=5)

# Subject
Label(main_frame, text="Subject:", font=font_config, bg="lightblue").grid(row=4, column=0, padx=10, pady=5, sticky="w")
subject_entry = Entry(main_frame, width=40, font=font_config, bd=1, relief="solid")
subject_entry.grid(row=4, column=1, padx=10, pady=5)

# Attachment
Label(main_frame, text="Attachment:", font=font_config, bg="lightblue").grid(row=5, column=0, padx=10, pady=5, sticky="w")
attachment_label = Label(main_frame, text="No file selected", width=40, anchor="w", font=font_config, bg="white")
attachment_label.grid(row=5, column=1, padx=10, pady=5)
Button(main_frame, text="Browse", command=select_attachment, font=font_config, bg="#4CAF50", fg="white").grid(row=5, column=2, padx=10, pady=5)

# Excel File Selection
Label(main_frame, text="Excel File:", font=font_config, bg="lightblue").grid(row=6, column=0, padx=10, pady=5, sticky="w")
excel_file_label = Label(main_frame, text="No file selected", width=40, anchor="w", font=font_config, bg="white")
excel_file_label.grid(row=6, column=1, padx=10, pady=5)
Button(main_frame, text="Browse", command=select_excel_file, font=font_config, bg="#4CAF50", fg="white").grid(row=6, column=2, padx=10, pady=5)

# Schedule Date and Time
Label(main_frame, text="Schedule Date (YYYY-MM-DD):", font=font_config, bg="lightblue").grid(row=8, column=0, padx=10, pady=5, sticky="w")
schedule_date_entry = Entry(main_frame, width=25, font=font_config, bd=1, relief="solid")
schedule_date_entry.grid(row=8, column=1, padx=10, pady=5)
Button(main_frame, text="Select Date", command=open_calendar, bg="#4CAF50", fg="white").grid(row=8, column=2, padx=10, pady=5)

Label(main_frame, text="Schedule Time (HH:MM):", font=font_config, bg="lightblue").grid(row=9, column=0, padx=10, pady=5, sticky="w")
schedule_time_entry = Entry(main_frame, width=25, font=font_config, bd=1, relief="solid")
schedule_time_entry.grid(row=9, column=1, padx=10, pady=5)

# Buttons for Sending Emails
Button(main_frame, text="Send Emails", command=send_bulk_emails, bg="#007BFF", fg="white", font=font_config, width=20).grid(row=10, column=0, pady=10, columnspan=3)

root.mainloop()