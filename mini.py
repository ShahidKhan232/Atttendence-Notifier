import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
import schedule
import time
import pandas as pd

def send_reminder_email():
    sender_email = sender_email_entry.get()
    sender_password = sender_password_entry.get()
    subject = subject_entry.get()
    message = message_text.get("1.0", tk.END)

    # Load the selected Excel file
    file_path = file_path_entry.get()
    df = pd.read_excel(file_path)
    email_ids = df['Email'].dropna().tolist()
    try:
        wb = load_workbook(file_path)
        sheet = wb.active
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
        return

    # Set up the SMTP server
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    # Create a secure SMTP connection
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)
    for receiver_email in email_ids:
            # Compose the email message
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = receiver_email
            msg['Subject'] = subject
            msg.attach(MIMEText(message, 'plain'))

            # Send the email
            server.send_message(msg)
            messagebox.showinfo("Success", "Reminder emails sent successfully!")
    try:
        # Extract email IDs from specified column (Column A in this example)
        email_ids = [cell.value for cell in sheet['A'] if cell.value is not None]

        # Send email to each recipient
        
        
        messagebox.showinfo("Success", "Reminder emails sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    # Close the connection
    server.quit()
# Function to browse and select an Excel file
# Modify the browse_excel_file function to print the selected file path
# Modify the browse_excel_file function to print the selected file path
def browse_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    print("Selected file path:", file_path)  # Debug statement
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, file_path)

# Create the main window
root = tk.Tk()
root.title("Email Sender")

# Email sender details (unchanged)
sender_email_label = tk.Label(root, text="Sender Email:")
sender_email_label.grid(row=0, column=0, padx=5, pady=5)
sender_email_entry = tk.Entry(root)
sender_email_entry.grid(row=0, column=1, padx=5, pady=5)

sender_password_label = tk.Label(root, text="Password:")
sender_password_label.grid(row=1, column=0, padx=5, pady=5)
sender_password_entry = tk.Entry(root, show="*")
sender_password_entry.grid(row=1, column=1, padx=5, pady=5)


# Email content (unchanged)
subject_label = tk.Label(root, text="Subject:")
subject_label.grid(row=2, column=0, padx=5, pady=5)
subject_entry = tk.Entry(root)
subject_entry.grid(row=2, column=1, padx=5, pady=5)

message_label = tk.Label(root, text="Message:")
message_label.grid(row=3, column=0, padx=5, pady=5)
message_text = tk.Text(root, height=5, width=30)
message_text.grid(row=3, column=1, padx=5, pady=5)

# Browse Excel file button
browse_button = tk.Button(root, text="Browse Excel File", command=browse_excel_file)
browse_button.grid(row=4, columnspan=2, padx=5, pady=5)

# Entry to display selected file path
file_path_entry = tk.Entry(root, state='readonly')
file_path_entry.grid(row=5, columnspan=2, padx=5, pady=5)

# Send button (unchanged)
send_button = tk.Button(root, text="Send Reminder Email", command=send_reminder_email)
send_button.grid(row=6, columnspan=2, padx=5, pady=5)

# Start the GUI event loop
root.mainloop()