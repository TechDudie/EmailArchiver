# EmailArchiver by TechnoDot
# Based on a IMAP implementation by pl608

import imaplib
import email
from email.header import decode_header
import os
data = open("credentials.txt").read().split()
username = data[0]
password = data[1]
try:
    os.mkdir("emails")
except:
    print("Emails folder already exists, move/delete it plz")
    exit()
imap = imaplib.IMAP4_SSL("imap.gmail.com", port=993)
imap.login(username, password)
print("Logged in to email")
status, messages = imap.select("INBOX")
messages = int(messages[0])
print(str(messages) + " messages to download")
for i in range(messages, 0, -1):
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                try:
                    subject = subject.decode(encoding)
                except:
                    try:
                        subject = subject.decode("utf-8")
                    except:
                        print("Unable to decode email")
                        continue
            sender, encoding = decode_header(msg.get("From"))[0]
            if isinstance(sender, bytes):
                try:
                    sender = sender.decode(encoding)
                except:
                    try:
                        sender = sender.decode("utf-8")
                    except:
                        print("Unable to decode email")
                        continue
            timestamp = msg["date"]
            file_title = f"{subject} | {sender} | {timestamp}".replace("/", "\\")
            text_body = f"Subject: {subject}\nFrom: {sender}\nTimestamp: {timestamp}\n\n"
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        text_body += body
                        save = open(os.path.join("emails", file_title + ".txt"), "w")
                        save.write(text_body)
                        save.close()
                    elif "attachment" in content_disposition:
                        filename = part.get_filename()
                        if filename:
                            open(os.path.join("emails", filename), "wb").write(part.get_payload(decode=True))
            else:
                content_type = msg.get_content_type()
                body = msg.get_payload(decode=True).decode()
                if content_type == "text/plain":
                    text_body += body
                    save = open(os.path.join("emails", file_title + ".txt"), "w")
                    save.write(text_body)
                    save.close()
            if content_type == "text/html":
                save = open(os.path.join("emails", file_title + ".html"), "w")
                save.write(body)
                save.close()
imap.close()
imap.logout()