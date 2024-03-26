import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import smtplib
from email.message import EmailMessage

# Function to send individual email
def send_email(from_email, password, recipient_email, subject, body, attachment_path=None):
    try:
        # Create a new message
        msg = EmailMessage()
        msg["From"] = from_email
        msg["To"] = recipient_email
        msg["Subject"] = subject
        msg.set_content(body)
        
        # Attach file if provided
        if attachment_path:
            with open(attachment_path, "rb") as file:
                file_data = file.read()
                file_name = attachment_path.split("/")[-1]
                msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

        # Send the email
        with smtplib.SMTP("smtp-mail.outlook.com", 587) as server:
            server.starttls()
            server.login(from_email, password)
            server.send_message(msg)
        
        # Print success message
        print(f"Email sent successfully to {recipient_email}")
        
        # Return success message
        return f"Email sent to {recipient_email} successfully!"
    except Exception as e:
        # Return error message
        return f"Error sending email to {recipient_email}: {str(e)}"

# Function to send mass emails
def send_mass_emails():
    # Get input values
    from_email = from_email_entry.get()
    password = password_entry.get()
    subject = subject_entry.get()
    body = body_text.get("1.0", tk.END).strip()
    from_row = int(from_row_entry.get())
    to_row = int(to_row_entry.get())
    attachment_path = attachment_entry.get()
    recipients_file = recipients_entry.get()
    
    try:
        # Read recipient emails from Excel file
        df = pd.read_excel(recipients_file)
        
        # Iterate over recipients and send emails
        email_reports = []
        for index, row in df.iloc[from_row-1:to_row].iterrows():
            recipient_email = row.iloc[0]
            report = send_email(from_email, password, recipient_email, subject, body, attachment_path)
            email_reports.append(report)

        # Show email reports
        messagebox.showinfo("Email Reports", "\n".join(email_reports))
    except Exception as e:
        # Show error message
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create main application window
window = tk.Tk()
frame = tk.Frame(window)
frame2 = tk.Frame(window, borderwidth=100)

frame2.pack(side=BOTTOM, fill=X)
w = 800 
h = 650 

ws = window.winfo_screenwidth() # width of the screen
hs = window.winfo_screenheight() # height of the screen

# calculate x and y coordinates for the Tk root window
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)

# set the dimensions of the screen and where it is placed
window.geometry('%dx%d+%d+%d' % (w, h, x, y))
window.columnconfigure(0, weight=1)

# Create labels and input fields
from_email_label = ttk.Label(frame, text="From Email:", justify="left").grid(row=0, column=0, padx=10, pady=5)
from_email_entry = ttk.Entry(frame , justify="left").grid(row=0, column=1, padx=10, pady=5,)

password_label = ttk.Label(frame, text="Password:", justify="left").grid(row=1, column=0, padx=10, pady=5)
password_entry = ttk.Entry(frame, show="*", justify="left").grid(row=1, column=1, padx=10, pady=5)

subject_label = ttk.Label(frame, text="Subject:", justify="left").grid(row=2, column=0, padx=10, pady=5)
subject_entry = ttk.Entry(frame , justify="left").grid(row=2, column=1, padx=10, pady=5)

body_label = ttk.Label(frame, text="Body:", justify="left").grid(row=3, column=0, padx=10, pady=5)
body_text = tk.Text(frame, height=5, width=20).grid(row=3, column=1, padx=10, pady=5)

attachment_label = ttk.Label(frame, text="Attachment:").grid(row=4, column=0, padx=10, pady=5)
attachment_entry = ttk.Entry(frame).grid(row=4, column=1, padx=10, pady=5)

recipient_label = ttk.Label(frame, text="Recipient File:").grid(row=5, column=0, padx=10, pady=5)
recipients_entry = ttk.Entry(frame).grid(row=5, column=1, padx=10, pady=5)

from_row_label = ttk.Label(frame, text="From Row:").grid(row=6, column=0, padx=10, pady=5)
from_row_entry = ttk.Entry(frame).grid(row=6, column=1, padx=10, pady=5)

to_row_label = ttk.Label(frame, text="To Row:").grid(row=7, column=0, padx=10, pady=5)
to_row_entry = ttk.Entry(frame).grid(row=7, column=1, padx=10, pady=5)

frame.pack()

# Function to browse and attach file
def attach_file():
    file_path = filedialog.askopenfilename()
    attachment_entry.delete(0, tk.END)
    attachment_entry.insert(0, file_path)

attach_button = ttk.Button(frame, text="Attach File", command=attach_file)
attach_button.grid(row=4, column=2, columnspan=30, padx=5, pady=5, sticky="ew")

# Function to browse and attach recipient file
def attach_recipient_file():
    file_path = filedialog.askopenfilename()
    recipients_entry.delete(0, tk.END)
    recipients_entry.insert(0, file_path)

recipient_button = ttk.Button(frame, text="Attach Recipient File", command=attach_recipient_file).grid(row=5, column=2, padx=5, pady=5)

# Function to send mass emails
send_button = ttk.Button(frame, text="Send", command=send_mass_emails).grid(row=8, column=1, padx=10, pady=5)

# Print designed by Aparimit Singhal
label1 = Label(frame2, text="Designed by Aparimit Singhal (R) 2024 All rights reserved", fg="black",
               bd=1, relief=SUNKEN, pady=10,padx=20)
label1.pack(side=BOTTOM, fill=X)
# Start the application
window.mainloop()