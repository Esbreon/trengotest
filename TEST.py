import imaplib
import os

def simple_imap_login():
    imap_server = "outlook.office365.com"
    email_address = os.getenv('OUTLOOK_EMAIL')
    password = os.getenv('OUTLOOK_PASSWORD')

    try:
        print("Trying to connect to IMAP server...")
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(email_address, password)
        print("Login successful")
        imap.logout()
    except imaplib.IMAP4.error as e:
        print(f"IMAP error: {e}")
    except Exception as e:
        print(f"General error: {e}")

# Zorg dat OUTLOOK_EMAIL en OUTLOOK_PASSWORD zijn ingesteld in je terminal voordat je dit script uitvoert
simple_imap_login()
