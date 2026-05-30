from gmail_auth import get_gmail_service
import base64
import os

def fetch_emails(query="in:inbox", output="./gmail_emails"):
    service = get_gmail_service()
    results = service.users().messages().list(userId="me", q=query).execute()

    os.makedirs(output, exist_ok=True)

    for msg in results.get("messages", []):
        msg_id = msg["id"]
        message = service.users().messages().get(userId="me", id=msg_id, format="raw").execute()
        raw = base64.urlsafe_b64decode(message["raw"])

        with open(f"{output}/{msg_id}.eml", "wb") as f:
            f.write(raw)
