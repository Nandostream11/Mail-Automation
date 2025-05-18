import os
import base64
import pandas as pd
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from email import message_from_string

# === Authentication ===
def gmail_authenticate():
    SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    return build('gmail', 'v1', credentials=creds)

# === Fetch draft by subject ===
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

# === Send updated message ===
def send_draft_to_recipient(service, draft_data, recipient_email):
    msg_raw = draft_data.get('message', {}).get('raw')
    if not msg_raw:
        raise Exception("Draft message doesn't contain 'raw' content.")

    msg_bytes = base64.urlsafe_b64decode(msg_raw.encode('utf-8'))
    msg_str = msg_bytes.decode('utf-8')
    mime_msg = message_from_string(msg_str)

    if mime_msg['To']:
        mime_msg.replace_header("To", recipient_email)
    else:
        mime_msg.add_header("To", recipient_email)

    updated_raw = base64.urlsafe_b64encode(mime_msg.as_bytes()).decode('utf-8')
    message = {'raw': updated_raw}
    send_result = service.users().messages().send(userId='me', body=message).execute()
    print(f"Email sent to {recipient_email}. Message ID: {send_result['id']}")

# === Main ===
def main():
    service = gmail_authenticate()
    df = pd.read_excel('gg.xlsx', sheet_name='Sheet1')

    if 'Email' not in df.columns:
        raise KeyError("Excel file must contain 'Email' column.")

    draft_data = get_template_body(service, "Robotics Intern Application- Anand Vardhan") #Subject of the mail

    for index, row in df.iterrows():
        recipient_email = row['Email']
        send_draft_to_recipient(service, draft_data, recipient_email)

if __name__ == '__main__':
    main()
