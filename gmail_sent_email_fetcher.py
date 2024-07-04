import imaplib
import email
from email.header import decode_header
import datetime
import re
import csv


# IMAP server and credentials
IMAP_SERVER = ""
EMAIL = ""
PASSWORD = ""

def decode_email_header(header):
    # Decode email header, handling errors gracefully
    decoded_parts = []
    for part, encoding in decode_header(header):
        if isinstance(part, bytes):
            try:
                decoded_parts.append(part.decode(encoding or 'utf-8', errors='replace'))
            except Exception as e:
                decoded_parts.append(f"Decode error: {e}")
        else:
            decoded_parts.append(part)
    return ''.join(decoded_parts)

def fetch_emails(start_date, end_date):
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select('"[Gmail]/Sent Mail"') # Use 'inbox' or '[Gmail]/Sent Mail' as per your requirement

        start_date = datetime.datetime.strptime(start_date, "%Y/%m/%d").strftime("%d-%b-%Y")
        end_date = (datetime.datetime.strptime(end_date, "%Y/%m/%d") + datetime.timedelta(days=1)).strftime("%d-%b-%Y")

        # result, data = mail.search(None, f'(SINCE "{start_date}" BEFORE "{end_date}")')
        result, data = mail.search(None, f'(SENTSINCE "{start_date}" SENTBEFORE "{end_date}")')

        messages = []
        for num in data[0].split():
            typ, msg_data = mail.fetch(num, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    email_data = {}
                    email_data['From'] = decode_email_header(msg['From'])
                    email_data['Subject'] = decode_email_header(msg['Subject'])
                    email_data['To'] = decode_email_header(msg['To'])

                    # Process email content
                    content = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            if content_type == 'text/plain':
                                try:
                                    content += part.get_payload(decode=True).decode()
                                except Exception as e:
                                    content += f"Decode error: {e}\n"
                    else:
                        content_type = msg.get_content_type()
                        if content_type == 'text/plain':
                            try:
                                content = msg.get_payload(decode=True).decode()
                            except Exception as e:
                                content = f"Decode error: {e}"

                    email_data['Content'] = content.strip()

                    messages.append(email_data)

        mail.logout()
        return messages

    except Exception as e:
        print(f"An error occurred: {e}")
        return []
    
def parse_email_content(email_content):
    # Regex pattern to match cafe name after "Dear"
    pattern = r"Dear\s+(?P<cafe>[^,.]+)"
    match = re.search(pattern, email_content, re.IGNORECASE | re.DOTALL)
    if match:
        return match.group('cafe').strip()
    else:
        return None
    
def check_duplicate(cafe_content, to_email, existing_entries):
    for entry in existing_entries:
        if entry['Cafe'] == cafe_content and entry['Email'] == to_email:
            return True
    return False

def main():
    start_date = '2024/01/15'
    end_date = '2024/01/23'

    emails = fetch_emails(start_date, end_date)

    # CSV file setup
    csv_filename = 'parsed_cafe_names.csv'
    csv_header = ['Cafe', 'Email']

    existing_entries = []
    try:
        with open(csv_filename, mode='r', newline='', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                existing_entries.append(row)
    except FileNotFoundError:
        pass 


    with open(csv_filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(csv_header)

        for email in emails:
            cafe_content = parse_email_content(email['Content'])
            to_email = email['To']
            if cafe_content:
                if not check_duplicate(cafe_content, to_email, existing_entries):
                    print(f"Parsing result - Cafe: {cafe_content}, Email: {to_email}")
                    print("-"*30)
                    writer.writerow([cafe_content, to_email])
                    existing_entries.append({'Cafe': cafe_content, 'Email': to_email})
                else:
                    print(f"Cafe info already in the file - Cafe: {cafe_content}, Email: {to_email}")
                    print("-"*30)
            else:
                print("Parsing failed:", email['Content'][:80])
                print("-"*30)
    print("DONE")



if __name__ == '__main__':
    main()