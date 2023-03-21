import imaplib
import base64
import os
import email



email_user = "*@gmail.com"

email_pass = "xyz"

# Connect to the email server:
mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
mail.login(email_user, email_pass)

# Select the inbox folder:
mail.select("Templates")

# Read all email UIDs and retrieve the latest received email:
type, data = mail.search(None, "ALL")
mail_ids = data[0]
latest_email_uid = data[0].split()[-1]
id_list = mail_ids.split()

for num in data[0].split():
    typ, data = mail.fetch(num, "(RFC822)")
    raw_email = data[0][1]

    # converts byte literal to string removing b''
    raw_email_string = raw_email.decode("utf-8")
    email_message = email.message_from_string(raw_email_string)

    # downloading attachments
    for part in email_message.walk():
        # this part comes from the snipet I don't understand yet...
        if part.get_content_maintype() == "multipart":
            continue
        if part.get("Content-Disposition") is None:
            continue
        fileName = part.get_filename()
        if bool(fileName):
            filePath = os.path.join("/Users/matthew.lee/Invoice", fileName)
            if not os.path.isfile(filePath):
                fp = open(filePath, "wb")
                fp.write(part.get_payload(decode=True))
                fp.close()
            subject = str(email_message).split("Subject: ", 1)[1].split("\nTo:", 1)[0]
            print(
                'Downloaded "{file}" from email titled "{subject}" with UID {uid}.'.format(
                    file=fileName, subject=subject, uid=latest_email_uid.decode("utf-8")
                )
            )
