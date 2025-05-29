# This script sends personalized emails using a Gmail draft template to a list of recipients from an Excel file.
# It requires the Gmail API and pandas library to read the Excel file.
import os
import base64
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from email import message_from_string

# === Load environment variables ===
load_dotenv()

TOKEN_PATH = os.getenv("GMAIL_TOKEN_PATH", "token.json")
SPREADSHEET_PATH = os.getenv("SPREADSHEET_PATH", "gg.xlsx")
SPREADSHEET_SHEET = os.getenv("SPREADSHEET_SHEET", "Sheet1")
DRAFT_SUBJECT = os.getenv("DRAFT_SUBJECT")

# === Gmail authentication ===
def gmail_authenticate():
    if not os.path.exists(TOKEN_PATH):
        raise FileNotFoundError(f"Token file not found: {TOKEN_PATH}")
    SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
    creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    return build('gmail', 'v1', credentials=creds)

# === Fetch draft matching a subject ===
def get_template_body(service, subject_text):
    drafts = service.users().drafts().list(userId='me').execute()
    for draft in drafts.get('drafts', []):
        draft_id = draft['id']
        draft_data = service.users().drafts().get(userId='me', id=draft_id, format='raw').execute()
        msg_raw = draft_data.get('message', {}).get('raw')

        if msg_raw:
            msg_bytes = base64.urlsafe_b64decode(msg_raw.encode('UTF-8'))
            msg_str = msg_bytes.decode('utf-8')
            msg_obj = message_from_string(msg_str)
            if msg_obj.get('Subject') == subject_text:
                return draft_data
    raise Exception(f"Draft with subject '{subject_text}' not found.")

# === Send personalized draft ===
def send_draft_to_recipient(service, draft_data, recipient_email):
    msg_raw = draft_data.get('message', {}).get('raw')
    if not msg_raw:
        raise Exception("Draft message doesn't contain 'raw' content.")

    msg_bytes = base64.urlsafe_b64decode(msg_raw.encode('utf-8'))
    msg_str = msg_bytes.decode('utf-8')
    mime_msg = message_from_string(msg_str)

    # Replace recipient
    if mime_msg['To']:
        mime_msg.replace_header("To", recipient_email)
    else:
        mime_msg.add_header("To", recipient_email)

    # Re-encode and send
    updated_raw = base64.urlsafe_b64encode(mime_msg.as_bytes()).decode('utf-8')
    message = {'raw': updated_raw}
    send_result = service.users().messages().send(userId='me', body=message).execute()
    print(f" Email sent to {recipient_email}. Message ID: {send_result['id']}")

# === Main function ===
def main():
    if not os.path.exists(SPREADSHEET_PATH):
        raise FileNotFoundError(f"Spreadsheet not found at {SPREADSHEET_PATH}")

    df = pd.read_excel(SPREADSHEET_PATH, sheet_name=SPREADSHEET_SHEET)
    if 'Email' not in df.columns:
        raise KeyError("Excel file must contain an 'Email' column.")

    service = gmail_authenticate()
    draft_data = get_template_body(service, DRAFT_SUBJECT)

    for index, row in df.iterrows():
        recipient_email = row['Email']
        if not isinstance(recipient_email, str) or not recipient_email.strip():
            print(f"Skipping invalid email at row {index}: {recipient_email}")
            continue
        send_draft_to_recipient(service, draft_data, recipient_email.strip())

if __name__ == '__main__':
    main()
