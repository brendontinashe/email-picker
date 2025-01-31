import imaplib
import email
import re
from email.header import decode_header
import openpyxl
import sys

def clean(text):
    return "".join(c if c.isalnum() else "_" for c in text)

def get_email_content(msg):
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                return part.get_payload(decode=True).decode()
    else:
        return msg.get_payload(decode=True).decode()

def extract_info(body):
    date_match = re.search(r'Date:\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})', body)
    centre_match = re.search(r'Test Centre:\s*(.+)', body)
    
    date = date_match.group(1) if date_match else ''
    centre = centre_match.group(1) if centre_match else ''
    
    return date, centre

def main():
    username = input("Enter your email address: ")
    password = input("Enter your email password: ")
    
    imap_server = "imap.gmail.com"
    imap = None
    
    try:
        print("Connecting to the IMAP server...")
        imap = imaplib.IMAP4_SSL(imap_server)
        
        print("Attempting to log in...")
        imap.login(username, password)
        
        print("Login successful. Selecting inbox...")
        imap.select("INBOX")
        
        print("Searching for emails...")
        _, message_numbers = imap.search(None, "ALL")
        
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Subject", "Date", "Test Centre"])
        
        print("Processing emails...")
        for num in message_numbers[0].split():
            _, msg_data = imap.fetch(num, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding or "utf-8")
                    
                    body = get_email_content(msg)
                    date, centre = extract_info(body)
                    
                    sheet.append([subject, date, centre])
        
        workbook.save("email_data.xlsx")
        print("Data has been saved to email_data.xlsx")
        
    except imaplib.IMAP4.error as e:
        print(f"An IMAP error occurred: {e}")
        print("This could be due to:")
        print("1. Incorrect email or password")
        print("2. Less secure app access is not enabled (for Gmail)")
        print("3. You're not using an App Password (for Gmail with 2-factor authentication)")
        print("\nPlease check your credentials and Gmail settings.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    finally:
        if imap:
            try:
                imap.close()
                imap.logout()
            except:
                pass  # Ignore errors during logout

if __name__ == "__main__":
    main()