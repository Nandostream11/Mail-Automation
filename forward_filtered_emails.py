# This script forwards emails from a specific sender that match certain criteria.
# It checks the subject for a keyword, verifies recipient conditions, and sends the email as an attachment.
import os
import base64
import datetime
import email
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import platform

# Load environment variables
load_dotenv()

FORWARD_TO = os.getenv('FORWARD_TO')
CREDENTIALS_FILE = os.getenv('CREDENTIALS_FILE', 'credentials.json')
TOKEN_FILE = os.getenv('TOKEN_FILE', 'token.json')

if platform.system() == "Windows":
    import winsound

SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose'
]

def authenticate_gmail():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
            with open(TOKEN_FILE, 'w') as token:
                token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def get_date_24_months_ago():
    today = datetime.date.today()
    last_year = today.replace(year=today.year - 2)
    return last_year.strftime('%Y/%m/%d')

def search_emails(service, query):
    messages = []
    request = service.users().messages().list(userId='me', q=query, maxResults=100)
    while request is not None:
        response = request.execute()
        messages.extend(response.get('messages', []))
        request = service.users().messages().list_next(request, response)
    return messages

def subject_contains_keyword(service, msg_id, keyword):
    try:
        msg_data = service.users().messages().get(
            userId='me',
            id=msg_id,
            format='metadata',
            metadataHeaders=["Subject"]
        ).execute()
        headers = {h['name']: h['value'] for h in msg_data.get('payload', {}).get('headers', [])}
        subject = headers.get("Subject", "").lower()
        return keyword.lower() in subject
    except Exception as e:
        print(f"X Error checking subject for message {msg_id}: {e}")
        return False

def get_recipients(service, msg_id):
    try:
        msg = service.users().messages().get(userId='me', id=msg_id, format='metadata', metadataHeaders=[
            "To", "Cc", "Bcc", "Delivered-To", "X-Original-To"
        ]).execute()
        headers_list = msg.get('payload', {}).get('headers', [])
        headers = {h['name']: h['value'] for h in headers_list}
        all_fields = [
            headers.get("To", ""),
            headers.get("Delivered-To", ""),
            headers.get("X-Original-To", "")
        ]
        recipients = " ".join(all_fields).lower()
        return recipients
    except Exception as e:
        print(f"X Error retrieving recipients for message {msg_id}: {e}")
        return ""

def forward_email(service, msg_id, forward_to):
    try:
        if not subject_contains_keyword(service, msg_id, "Campus Placement"):
            print(f"Skipping message {msg_id} — subject does not contain keyword.")
            return

        recipients = get_recipients(service, msg_id)
        if "FILTER_MAIL" not in recipients:
            print(f"Skipping message {msg_id} — recipient condition not met.")
            return

        original_msg = service.users().messages().get(userId='me', id=msg_id, format='raw').execute()
        raw_bytes = base64.urlsafe_b64decode(original_msg['raw'].encode("UTF-8"))

        outer = MIMEMultipart()
        outer['To'] = forward_to
        outer['From'] = 'me'
        outer['Subject'] = "Fwd: Message as requested"

        outer.attach(MIMEText("Forwarding attached email as requested.", 'plain'))

        eml_attachment = MIMEBase('message', 'rfc822')
        eml_attachment.set_payload(raw_bytes)
        encoders.encode_base64(eml_attachment)
        eml_attachment.add_header('Content-Disposition', 'attachment', filename='forwarded_message.eml')
        outer.attach(eml_attachment)

        raw_forward = base64.urlsafe_b64encode(outer.as_bytes()).decode()
        send_body = {"raw": raw_forward}
        service.users().messages().send(userId='me', body=send_body).execute()

        print(f" OK Safely forwarded message ID: {msg_id} \a")
        if platform.system() == "Windows":
            winsound.Beep(1000, 200)

    except Exception as e:
        print(f" Failed to forward message ID {msg_id}: {e}")

def main():
    service = authenticate_gmail()
    if not FORWARD_TO:
        print("FORWARD_TO not defined in .env.")
        return

    date_24_months_ago = get_date_24_months_ago()
    query = f'from:FROM after:{date_24_months_ago}'
    messages = search_emails(service, query)

    if not messages:
        print("No matching messages found.")
        return

    print(f"Found {len(messages)} messages. Forwarding to {FORWARD_TO}...\n")
    for msg in messages:
        forward_email(service, msg['id'], FORWARD_TO)

if __name__ == '__main__':
    main()
