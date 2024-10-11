
# Attendance Notifier

This Python application allows users to send reminder emails to multiple recipients by reading email addresses from an Excel file. It's built with `tkinter` for the graphical user interface (GUI) and uses `smtplib` for sending emails via SMTP.

## Features
- Send bulk reminder emails.
- Extract email addresses from an Excel file.
- Simple and user-friendly interface.
- Supports Gmail SMTP for sending emails.
  
## Requirements

Ensure you have the following Python packages installed before running the program:

- `pandas`   (included with standard Python installations)
- `openpyxl` (included with standard Python installations)
- `tkinter` (included with standard Python installations)
- `smtplib` (included with standard Python installations)

You can install missing dependencies using `pip`:

```bash
pip install pandas openpyxl
```

## Usage

1. **Clone the repository**:

    ```bash
    git clone https://github.com/yourusername/email-reminder-sender.git
    cd email-reminder-sender
    ```

2. **Run the script**:

    Run the Python script using the following command:

    ```bash
    python email_reminder.py
    ```

3. **Instructions**:
    - **Sender Email and Password**: Enter the sender's email and password (for now, the app only supports Gmail accounts).
    - **Subject**: Add a subject for the email.
    - **Message**: Enter the message body.
    - **Browse Excel File**: Use the button to select an Excel file (.xlsx) that contains recipient email addresses. The emails should be listed in a column with the header `Email`.
    - **Send Reminder Email**: Once everything is filled out, click the button to send the email to all recipients.

## Excel File Format

The application expects the Excel file to have a column labeled `Email` that contains the email addresses of the recipients. Any other columns or data are ignored.

### Example Excel File:

| Email             |
|-------------------|
| example1@gmail.com|
| example2@gmail.com|
| example3@gmail.com|

## Contributing

If you'd like to contribute to this project, please fork the repository and submit a pull request.

1. Fork the Project.
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`).
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`).
4. Push to the Branch (`git push origin feature/AmazingFeature`).
5. Open a Pull Request.

## Contact

For issues, suggestions, or questions, feel free to open an issue on the repository or contact me via shahidkhan.95173@gmail.com.
