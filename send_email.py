import pandas as pd
import os
import pickle
import base64
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Define the Gmail API scope
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def authenticate_gmail():
    """Authenticate with Gmail API and return the service object."""
    creds = None
    token_path = 'token.json'
    credentials_path = r'C:\Users\sansk\project\credentials.json'  # Ensure correct path

    # Load existing credentials if available
    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)

    # If credentials are invalid or missing, refresh or re-authenticate
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(credentials_path):
                print(f"Error: Credentials file not found at {credentials_path}")
                return None
            
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)

        # Save new credentials
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)

    return build('gmail', 'v1', credentials=creds)

def send_email(service, sender, recipient, subject, body):
    """Send an email using Gmail API."""
    try:
        message = MIMEMultipart()
        message['From'] = sender
        message['To'] = recipient
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain'))

        raw_message = base64.urlsafe_b64encode(message.as_string().encode('utf-8')).decode('utf-8')
        email_body = {'raw': raw_message}

        sent_message = service.users().messages().send(userId="me", body=email_body).execute()
        print(f" Email sent successfully! Message ID: {sent_message['id']}")
    except Exception as e:
        print(f" Error sending email: {e}")

def search_student(file_path, prn):
    """Search for student details in the Excel file using PRN."""
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            print(f" Error: File '{file_path}' not found.")
            return None

        # Read Excel file
        df = pd.read_excel(file_path)

        # Check if PRN exists
        if 'PRN' not in df.columns:
            print("Error: PRN column not found in the Excel file.")
            return None

        student_row = df[df['PRN'] == prn]
        if student_row.empty:
            print(" Student not found.")
            return None

        return student_row.to_dict(orient='records')[0]

    except Exception as e:
        print(f" Error reading Excel file: {e}")
        return None

def main():
    # File path to student data
    excel_file_path = r'C:\Users\sansk\OneDrive\文档\students.xlsx'

    try:
        # Get PRN from user
        prn = int(input(" Enter the PRN number of the student: "))

        # Search for student details
        student = search_student(excel_file_path, prn)
        if student is None:
            return

        # Display student details
        print("\n Student Details:")
        for key, value in student.items():
            print(f"{key}: {value}")

        # Get recipient email
        sender_email = "sanstorm164@gmail.com"  # Update with your email
        recipient_email = input("\n Enter recipient email address: ")

        # Prepare email content
        subject = f" Attendance Details for PRN {prn}"
        body = f" Here are the attendance details for PRN {prn}:\n\n"
        for key, value in student.items():
            body += f" {key}: {value}\n"

        # Authenticate and send email
        service = authenticate_gmail()
        if service:
            send_email(service, sender_email, recipient_email, subject, body)

    except ValueError:
        print(" Error: PRN must be a number.")

if _name_ == '_main_':
    main()