# email-picker
1. Core Workflow
Creates CSV File First

Initializes email_data.csv with headers: Subject, Date, Test Centre

Confirms creation with: "CSV file created successfully"

Handles Authentication

Securely collects email/password via terminal input

Connects to Gmail via IMAP over SSL (port 993)

Processes Emails

Searches all emails in the inbox

Decodes email subjects and bodies

Uses regex to extract:

Dates (YYYY-MM-DD HH:MM:SS)

Test Centers (text after "Test Centre:")

Saves Data

Appends extracted data to the CSV

Reports progress ("Processed X emails")

2. Key Components
Component	Purpose
create_csv()	Creates CSV file with headers
get_email_content()	Extracts plain text from email body (handles multipart emails)
extract_info()	Uses regex to find dates/test centers
Context Managers	Safely handles files/connections (with statements)
3. Error Handling
Specific Errors:

IMAP authentication failures

Network/connection issues

General Errors:

File permission problems

Unexpected crashes

User Feedback:

Clear error messages with troubleshooting hints

4. Key Features
Security: SSL encryption for email connections

Reliability: Automatic cleanup of resources

User Feedback: Progress updates and success/failure notifications

Compatibility: Works with any IMAP-enabled email provider

5. Important Notes
Requires IMAP access to be enabled in Gmail settings

Uses App Passwords if 2-factor authentication is enabled

Stores credentials only temporarily in memory

The script focuses on structured email processing while prioritizing security and user transparency.