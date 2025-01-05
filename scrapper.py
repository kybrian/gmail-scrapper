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
from urllib.parse import urlparse, urljoin
import requests
from io import BytesIO
from PIL import Image

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
EMAIL_ADDRESS = "targetemail@gmail.com"


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

def is_valid_url(url):
    try:
        result = urlparse(url)
        return all([result.scheme, result.netloc])
    except ValueError:
        return False


def download_image_data(url):
    try:
       if not is_valid_url(url):
            return None
       response = requests.get(url, stream=True)
       response.raise_for_status()  # Raises HTTPError for bad responses (4xx or 5xx)
       img_io = BytesIO(response.content)
       image = Image.open(img_io)
       if image.format != 'PNG':
            img_io.seek(0)
            img_io = BytesIO()
            image.save(img_io, format="PNG")
       img_io.seek(0)
       return base64.b64encode(img_io.read()).decode()

    except Exception as error:
        print(f"Error downloading image: {url} {error}")
        return None


def html_to_pdf(html_content, output_path):
    try:
        with open(output_path, "wb") as pdf_file:
            pisa_status = pisa.CreatePDF(html_content, dest=pdf_file)
        if pisa_status.err:
            print(f"Error converting to PDF: {pisa_status.err}")
    except Exception as error:
        print(f"Error converting to PDF: {error}")

def download_attachments(service, message, folder_path):
   for part in message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        file_name = part.get_filename()
        if bool(file_name):
            file_path = os.path.join(folder_path, file_name)
            if not os.path.isfile(file_path):
                try:
                    att_id = part.get_payload()
                    att = service.users().messages().attachments().get(userId='me', messageId=message['id'],id=att_id).execute()
                    data = att['data']
                    file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                    with open(file_path, 'wb') as f:
                        f.write(file_data)
                    print(f'  Downloaded attachment: {file_name}')
                except Exception as error:
                    print(f'Error saving attachment: {file_name} {error}')


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

            folder_name = 'downloaded_emails_pdf'
            os.makedirs(folder_name, exist_ok=True)

            # Create sub-folder for this specific email
            email_folder_name = f"{date_header}-{sender_header}-{subject_header}".replace(":","").replace("/","").replace("\\","")
            email_folder_path = os.path.join(folder_name, email_folder_name)
            os.makedirs(email_folder_path, exist_ok=True)

            file_path = os.path.join(email_folder_path, f"email.pdf")

            # Process Attachments
            download_attachments(service, mime_msg, email_folder_path)

            # Extract HTML
            html_content = ""
            if mime_msg.is_multipart():
                for part in mime_msg.walk():
                     if part.get_content_type() == 'text/html':
                        html_content = part.get_payload(decode=True).decode("utf-8", errors='ignore')
                        break
                if not html_content: #If there's no html part just get any other text
                    for part in mime_msg.walk():
                        if part.get_content_type() == 'text/plain':
                            html_content =  part.get_payload(decode=True).decode("utf-8", errors='ignore').replace('\n', '<br>')
                            break
            else:
                if mime_msg.get_content_type() == 'text/html':
                    html_content = mime_msg.get_payload(decode=True).decode("utf-8", errors='ignore')
                elif mime_msg.get_content_type() == 'text/plain':
                    html_content = mime_msg.get_payload(decode=True).decode("utf-8", errors='ignore').replace('\n', '<br>')

            if not html_content:
                html_content = "<body><h1>No Email Content</h1><body>"


            # Download and Base64 Encode Images
            soup = BeautifulSoup(html_content, "html.parser")
            for img in soup.find_all('img'):
                img_url = img.get('src')
                if img_url:
                    if not urlparse(img_url).netloc:
                        img_url = urljoin("https://mail.google.com", img_url)
                    img_data = download_image_data(img_url)
                    if img_data:
                        img['src'] = f'data:image/png;base64,{img_data}'

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
            {str(soup)}
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