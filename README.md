# Gmail Draft Sender Automation 📬

This project allows you to send formatted Gmail draft emails to a list of recipients using Python and the Gmail API. It maintains the original formatting and content of your Gmail draft and sends it to emails listed in an Excel sheet.

## Features 🚀

- Uses Gmail API for secure, authenticated access.
- Preserves formatting, links, and subject from your Gmail draft.
- Sends personalized emails in bulk using data from an Excel sheet.
- Minimal configuration, ready to scale.

## Project Structure 📁
.
├── auto_gmail.py               # Main script to send mails
├── gg.xlsx                     # Excel file with recipient emails (column: Email)
├── forward_filtered_emails.py  # Script to filter & forward mails in .eml format
├── token.json                  # OAuth token after authentication
├── credentials.json            # OAuth credentials (downloaded from Google Cloud)
└── README.md 

## Requirements 🧰

- Python 3.7+
- Gmail account
- Google Cloud project with Gmail API enabled & `credentials.json` downloaded

## Install Dependencies 📦

```bash
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib pandas openpyxl
```

## 🔑 Google Cloud Setup
1. Go to Google Cloud Console.
2. Create a new project (or use an existing one).
3. Enable Gmail API.
4. Go to APIs & Services → Credentials.
5. Click Create Credentials → OAuth Client ID.
6. Application type: Desktop App.
7. Download the credentials.json file and place it in your project folder.

### First-Time Authentication🔐
Run this snippet once to authenticate and generate token.json:

```bash
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
creds = flow.run_local_server(port=0)

with open('token.json', 'w') as token_file:
    token_file.write(creds.to_json())
```

## 📝 Prepare Gmail Draft
Go to Gmail Drafts

- Create a formatted draft with your content and subject (e.g., "Robotics Intern Application- Anand Vardhan")
- Leave the 'To' field empty — the script will fill it in.

📄 Excel Format (gg.xlsx)
- Make sure your Excel file (gg.xlsx) contains a column named exactly: Email
- Each row should contain one recipient email.

### Run the Script▶️
```bash
python auto_gmail.py
```
This will:
- Fetch the Gmail draft by subject
- Send it to each recipient in the Excel sheet
- Maintain original formatting and subject

🛑 Notes
- Be careful with sending limits on your Gmail account.
- Use test accounts when debugging.
- The script only sends one email per row; duplicates are not automatically filtered.

👨‍💻 Author
Anand Vardhan – [LinkedIn](https://www.linkedin.com/in/anandvardhanrbtics/)
Feel free to fork, star, or contribute!

🧾 License
MIT License — feel free to use, modify, and share.