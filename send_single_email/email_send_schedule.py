import smtplib
from tkinter import Tk, Label, Entry, Text, Button, filedialog, messagebox, Frame, Toplevel
from tkcalendar import Calendar
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import threading
import schedule
from datetime import datetime
import time

# SMTP Configuration
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587

# Send email function
def send_single_email():
    sender_email = sender_entry.get()
    sender_password = password_entry.get()
    recipient_email = recipient_entry.get()  # Get recipient email
    subject = subject_entry.get()  # Get subject
    message_body = message_text.get("1.0", "end").strip()  # Get message body from Text widget

    if not recipient_email:
        messagebox.showerror("Error", "Recipient email is required!")
        return

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(sender_email, sender_password)

            # Create the email message
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            msg['Subject'] = subject

            # Attach the message body
            if not message_body:
                message_body = "No message body provided."  # Default text if the message box is empty
            msg.attach(MIMEText(message_body, 'plain'))

            # Send the email
            server.send_message(msg)
            print(f"Email sent to {recipient_email}")

        messagebox.showinfo("Success", "Email sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send email: {e}")

# Open calendar
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

# Function to schedule the email
def schedule_email():
    schedule_date = schedule_date_entry.get()
    schedule_time = schedule_time_entry.get()

    try:
        # Combine date and time and convert to datetime object
        schedule_datetime_str = f"{schedule_date} {schedule_time}"
        schedule_datetime = datetime.strptime(schedule_datetime_str, "%Y-%m-%d %H:%M")

        # Get the current time and check if the scheduled time is in the future
        current_time = datetime.now()

        if schedule_datetime <= current_time:
            messagebox.showerror("Error", "Scheduled time must be in the future!")
            return

        # Schedule the task
        schedule.every().day.at(schedule_time).do(send_single_email)

        # Inform the user about the schedule
        messagebox.showinfo("Scheduled", f"Email will be sent on {schedule_datetime}")

        # Start a background thread to run the schedule
        threading.Thread(target=run_schedule, daemon=True).start()

    except ValueError:
        messagebox.showerror("Error", "Invalid date or time format. Use YYYY-MM-DD for date and HH:MM for time.")

# Function to run the schedule in the background
def run_schedule():
    while True:
        schedule.run_pending()
        time.sleep(1)

# Create GUI
root = Tk()
root.title("Email Sending Automation System")

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

Label(main_frame, text="Single Email Sending System with Scheduling", font=title_font, bg="lightblue").grid(row=0, column=0, columnspan=3, pady=20)

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

# Message Body
Label(main_frame, text="Message:", font=font_config, bg="lightblue").grid(row=5, column=0, padx=10, pady=5, sticky="w")
message_text = Text(main_frame, height=5, width=40, font=font_config)
message_text.grid(row=5, column=1, padx=10, pady=5)

# Schedule Date
Label(main_frame, text="Schedule Date (YYYY-MM-DD):", font=font_config, bg="lightblue").grid(row=6, column=0, padx=10, pady=5, sticky="w")
schedule_date_entry = Entry(main_frame, width=25, font=font_config, bd=1, relief="solid")
schedule_date_entry.grid(row=6, column=1, padx=10, pady=5)
Button(main_frame, text="Select Date", command=open_calendar, bg="#4CAF50", fg="white").grid(row=6, column=2, padx=10, pady=5)

# Schedule Time
Label(main_frame, text="Schedule Time (HH:MM):", font=font_config, bg="lightblue").grid(row=7, column=0, padx=10, pady=5, sticky="w")
schedule_time_entry = Entry(main_frame, width=25, font=font_config, bd=1, relief="solid")
schedule_time_entry.grid(row=7, column=1, padx=10, pady=5)

# Button for Sending Email Immediately
Button(main_frame, text="Send Email", command=send_single_email, bg="#007BFF", fg="white", font=font_config, width=20).grid(row=8, column=0, pady=10, columnspan=3)

root.mainloop()
