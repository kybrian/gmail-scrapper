import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
import email
from email.mime.text import MIMEText
from xhtml2pdf import pisa
from bs4 import BeautifulSoup

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
EMAIL_ADDRESS = "Insert Gmail Email Here"


def create_message(sender, to, subject, message_text):
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw_message}


def get_gmail_service():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)


def html_to_pdf(html_content, output_path):
    try:
        with open(output_path, "wb") as pdf_file:
           pisa_status = pisa.CreatePDF(html_content, dest=pdf_file)
        if pisa_status.err:
            print(f"Error converting to PDF: {pisa_status.err}")
    except Exception as error:
        print(f"Error converting to PDF: {error}")

def download_emails(service):
    email_address = EMAIL_ADDRESS
    user_id = 'me'
    query = f'from:{email_address}'
    try:
        results = service.users().messages().list(userId=user_id, q=query).execute()
        messages = results.get('messages', [])
        print(f"Found {len(messages)} messages from {email_address}")
        if not messages:
            print(f"No messages from {email_address}")
            return

        for message in messages:
            msg = service.users().messages().get(userId=user_id, id=message['id'], format='raw').execute()
            msg_str = base64.urlsafe_b64decode(msg['raw'].encode('ASCII'))
            mime_msg = email.message_from_bytes(msg_str)

            date_header = mime_msg.get('Date')
            subject_header = mime_msg.get('Subject')
            sender_header = mime_msg.get('From')

            # Extract the body text of email
            body_text = ""
            if mime_msg.is_multipart():
                for part in mime_msg.walk():
                    if part.get_content_type() == 'text/plain':
                        body_text = part.get_payload(decode=True).decode("utf-8", errors='ignore')
                        break #Take the plain text part first
                if not body_text: #Try HTML text
                    for part in mime_msg.walk():
                        if part.get_content_type() == 'text/html':
                            body_text = part.get_payload(decode=True).decode("utf-8", errors='ignore')
                            break
            else:
                body_text = mime_msg.get_payload(decode=True).decode("utf-8", errors='ignore')

            folder_name = 'downloaded_emails_pdf'
            os.makedirs(folder_name, exist_ok=True)
            file_name = f"{date_header}-{sender_header}-{subject_header}".replace(":","").replace("/","").replace("\\","")
            file_path = os.path.join(folder_name, f"{file_name}.pdf")

            # Add basic HTML formatting
            html_content = f"""
            <html>
            <head>
            <meta charset="utf-8">
            <title>{subject_header}</title>
            </head>
            <body>
            <h1>{subject_header}</h1>
            <p><b>From:</b> {sender_header}</p>
            <p><b>Date:</b> {date_header}</p>
            <pre style="white-space: pre-wrap;">{body_text}</pre>
            </body>
            </html>
            """
            html_to_pdf(html_content, file_path)
            print(f"Downloaded message ID: {message['id']} saved to {file_path}")


    except Exception as error:
        print(f'An error occurred: {error}')

if __name__ == '__main__':
    service = get_gmail_service()
    download_emails(service)