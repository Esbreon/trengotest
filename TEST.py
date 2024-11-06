import imaplib
import os

def test_outlook_connection():
    imap_server = "outlook.office365.com"
    email_address = os.environ.get('OUTLOOK_EMAIL')
    password = os.environ.get('OUTLOOK_PASSWORD')

    try:
        print("Connecting to Outlook IMAP server...")
        imap = imaplib.IMAP4_SSL(imap_server)
        
        # Set debug level for more insights during connection
        imap.debug = 4  # This will print detailed debugging information to stdout
        
        # Attempt login
        imap.login(email_address, password)
        print("Connection successful.")
        
        # Close connection
        imap.logout()
    except imaplib.IMAP4.error as e:
        print(f"IMAP error: {e}")
    except Exception as e:
        print(f"General error: {e}")

# Run test function
test_outlook_connection()
