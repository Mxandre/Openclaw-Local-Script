import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta

# ----------------------------
# email configuration
# ----------------------------
IMAP_SERVER = "***"  ##example : imaps.utc.fr
IMAP_PORT = 993
EMAIL = "***"
PASSWORD = "***"  # 

# ----------------------------
# connect to email server
# ----------------------------
mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
mail.login(EMAIL, PASSWORD)
mail.select("INBOX")  

# ----------------------------
# search for recent emails
# ----------------------------
since = (datetime.now() - timedelta(days=2)).strftime("%d-%b-%Y")
status, messages = mail.search(None, f'(SINCE "{since}")')
ids = messages[0].split()

# ----------------------------
# parse emails
# ----------------------------
emails = []

for i in ids[-5:]:  # only fetch the most recent 5 emails, can be adjusted
    status, msg_data = mail.fetch(i, "(RFC822)")
    raw_msg = msg_data[0][1]
    msg = email.message_from_bytes(raw_msg)

    # parse email subject, handle encoding automatically (in case of french or other languages)
    subject_header = decode_header(msg["Subject"])[0]
    subject = subject_header[0]
    encoding = subject_header[1]
    if isinstance(subject, bytes):
        if encoding:
            subject = subject.decode(encoding)
        else:
            subject = subject.decode("latin-1")  # fallback for common French encoding

    # parse sender
    from_header = decode_header(msg.get("From"))[0]
    sender = from_header[0]
    encoding = from_header[1]
    if isinstance(sender, bytes):
        if encoding:
            sender = sender.decode(encoding)
        else:
            sender = sender.decode("latin-1")

    # optional: parse email body (text part)
    body_text = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get("Content-Disposition"))
            if ctype == "text/plain" and "attachment" not in cdispo:
                charset = part.get_content_charset()
                payload = part.get_payload(decode=True)
                if charset:
                    body_text = payload.decode(charset, errors="replace")
                else:
                    body_text = payload.decode("latin-1", errors="replace")
                break
    else:
        charset = msg.get_content_charset()
        payload = msg.get_payload(decode=True)
        if charset:
            body_text = payload.decode(charset, errors="replace")
        else:
            body_text = payload.decode("latin-1", errors="replace")

    emails.append({
        "subject": subject,
        "from": sender,
        "body": body_text
    })

# ----------------------------
# output email summaries as JSON to be used by rob
# ----------------------------
import json
print(json.dumps(emails, ensure_ascii=False))
