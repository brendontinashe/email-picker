import imaplib
import email
import re
import csv
import sys
from email.header import decode_header

def create_csv():
    """Create CSV file with headers and return success status"""
    try:
        with open('email_data.csv', 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["Subject", "Date", "Test Centre"])
        print("CSV file 'email_data.csv' created successfully.\n")
        return True
    except Exception as e:
        print(f"⚠️ Error creating CSV file: {e}")
        return False

def get_email_content(msg):
    """Extract plain text content from email"""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                return part.get_payload(decode=True).decode()
    return msg.get_payload(decode=True).decode()

def extract_info(body):
    """Extract structured data using regex patterns"""
    date_match = re.search(r'Date:\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})', body)
    centre_match = re.search(r'Test Centre:\s*(.+)', body)
    return (
        date_match.group(1) if date_match else 'N/A',
        centre_match.group(1).strip() if centre_match else 'N/A'
    )

def main():
    # Step 1: Create CSV file
    if not create_csv():
        sys.exit(1)

    # Step 2: Get credentials
    print("Please enter your email credentials:")
    username = input("Email address: ")
    password = input("Password: ")
    print("\nAttempting to connect...")

    try:
        # Step 3: Connect to email server
        with imaplib.IMAP4_SSL("imap.gmail.com") as imap:
            imap.login(username, password)
            imap.select("INBOX")

            # Step 4: Process emails
            _, message_numbers = imap.search(None, "ALL")
            print(f"Found {len(message_numbers[0].split())} emails to process")

            with open('email_data.csv', 'a', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                processed_count = 0

                for num in message_numbers[0].split():
                    _, msg_data = imap.fetch(num, "(RFC822)")
                    
                    for response_part in msg_data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])
                            subject = decode_header(msg["Subject"])[0][0]
                            subject = subject.decode() if isinstance(subject, bytes) else str(subject)
                            
                            body = get_email_content(msg)
                            date, centre = extract_info(body)
                            
                            writer.writerow([subject, date, centre])
                            processed_count += 1

                print(f"\nSuccessfully processed {processed_count} emails")
                print("Data appended to email_data.csv")

    except imaplib.IMAP4.error as e:
        print(f"\n🔐 Authentication failed: {e}")
        print("Possible reasons:")
        print("- Incorrect email/password")
        print("- Less Secure Apps access disabled (for Gmail)")
        print("- 2FA enabled without app password")
    except Exception as e:
        print(f"\n⚠️ Unexpected error: {e}")

if __name__ == "__main__":
    main()