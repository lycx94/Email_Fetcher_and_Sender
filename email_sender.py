import os
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import mimetypes

def send_email(subject, body, to_email, attachment_path=None):
    gmail_email = ""
    gmail_password = ""

    message = MIMEMultipart()
    message["From"] = gmail_email
    message["To"] = to_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    if attachment_path:
        attach_file(message, attachment_path)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(gmail_email, gmail_password)
        server.sendmail(gmail_email, to_email, message.as_string())

def attach_file(message, attachment_path):
    mime_type, _ = mimetypes.guess_type(attachment_path)
    mime_type, mime_subtype = mime_type.split('/', 1)

    with open(attachment_path, 'rb') as ap:
        attachment = MIMEBase(mime_type, mime_subtype)
        attachment.set_payload(ap.read())
        encoders.encode_base64(attachment)
        attachment.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attachment_path)}")
        message.attach(attachment)

def email_draft(attribute):
    cafe = attribute[0]
    recipient_email = attribute[1]

    print("\nSending email to {}".format(recipient_email))

    subject = "Inquiry About Barista/All-Rounder position"
    body = """

            """

    attachment_path = ""
    
    try:
        send_email(subject, body, recipient_email, attachment_path)
    except Exception as e:
        print(f"An error occurred: {e}")


def main():
    try:
        with open("parsed_cafe_names.csv", "r", newline='', encoding='utf-8') as file:
            csv_reader = csv.reader(file)            
            next(csv_reader)
            for row in csv_reader:
                # print(row)
                email_draft(row)
    except Exception as e:
        print("Fail. {}".format(e))
    finally:
        print("DONE")


if __name__=="__main__":
    main()