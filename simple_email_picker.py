# Import required libraries
import imaplib       # For IMAP email protocol handling
import email         # For email message parsing
import re            # For regular expressions (pattern matching)
from email.header import decode_header  # For decoding email headers
import openpyxl      # For Excel spreadsheet operations
import sys           # For system-specific operations (not directly used here)

def clean(text):
    """Sanitize text by replacing non-alphanumeric characters with underscores.
    Useful for creating safe filenames or database fields."""
    return "".join(c if c.isalnum() else "_" for c in text)

def get_email_content(msg):
    """Extract plain text content from email message.
    Handles both multipart and single-part emails."""
    if msg.is_multipart():
        # Check all parts of multipart email
        for part in msg.walk():
            content_type = part.get_content_type()
            # Prioritize plain text content
            if content_type == "text/plain":
                return part.get_payload(decode=True).decode()
    else:
        # Handle simple email structure
        return msg.get_payload(decode=True).decode()

def extract_info(body):
    """Extract structured data from email body using regular expressions.
    Returns tuple containing (date, test_centre)"""
    # Pattern for date (YYYY-MM-DD HH:MM:SS format)
    date_match = re.search(r'Date:\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})', body)
    # Pattern for test center (any characters after "Test Centre:")
    centre_match = re.search(r'Test Centre:\s*(.+)', body)
    
    # Extract matches or return empty strings
    date = date_match.group(1) if date_match else ''
    centre = centre_match.group(1).strip() if centre_match else ''
    
    return date, centre

def main():
    # Collect user credentials securely (note: input() shows password in plaintext)
    username = input("Enter your email address: ")
    password = input("Enter your email password: ")
    
    # IMAP server configuration (Gmail specific)
    imap_server = "imap.gmail.com"
    imap = None  # IMAP connection handler
    
    try:
        # Establish secure IMAP connection
        print("Connecting to the IMAP server...")
        imap = imaplib.IMAP4_SSL(imap_server)
        
        # Authenticate user
        print("Attempting to log in...")
        imap.login(username, password)
        
        # Access inbox
        print("Login successful. Selecting inbox...")
        imap.select("INBOX")  # Could also specify 'readonly=True' for safety
        
        # Retrieve all email IDs
        print("Searching for emails...")
        _, message_numbers = imap.search(None, "ALL")  # 'ALL' could be changed to 'UNSEEN'
        
        # Create Excel workbook and header row
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Subject", "Date", "Test Centre"])  # Column headers
        
        # Process each email
        print("Processing emails...")
        for num in message_numbers[0].split():
            # Fetch entire email (RFC822 protocol)
            _, msg_data = imap.fetch(num, "(RFC822)")
            
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    # Parse email message
                    msg = email.message_from_bytes(response_part[1])
                    
                    # Decode subject header
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding or "utf-8")
                    
                    # Extract and process email content
                    body = get_email_content(msg)
                    date, centre = extract_info(body)
                    
                    # Add data to spreadsheet
                    sheet.append([subject, date, centre])
        
        # Save workbook to file
        workbook.save("email_data.xlsx")
        print("Data has been saved to email_data.xlsx")
        
    except imaplib.IMAP4.error as e:
        # Handle IMAP-specific errors
        print(f"An IMAP error occurred: {e}")
        print("This could be due to:")
        print("1. Incorrect email or password")
        print("2. Less secure app access is not enabled (for Gmail)")
        print("3. You're not using an App Password (for Gmail with 2-factor authentication)")
        print("\nPlease check your credentials and Gmail settings.")
    except Exception as e:
        # Catch-all for other exceptions
        print(f"An unexpected error occurred: {e}")
    finally:
        # Cleanup connection
        if imap:
            try:
                imap.close()
                imap.logout()
            except:
                pass  # Ignore errors during logout process

if __name__ == "__main__":
    # Start program execution
    main()