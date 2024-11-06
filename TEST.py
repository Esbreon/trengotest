import msal
import imaplib
import os

# Environment variabelen voor de app-configuratie
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
TENANT_ID = os.getenv('AZURE_TENANT_ID')

# Aanmelden via OAuth2 met MSAL
def get_oauth2_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    scope = ["https://outlook.office365.com/.default"]
    
    result = app.acquire_token_for_client(scopes=scope)
    if "access_token" in result:
        print("Access token verkregen!")
        return result['access_token']
    else:
        print("Fout bij verkrijgen token:", result.get("error_description"))
        return None

# Verbinden met IMAP met behulp van het OAuth2-token
def connect_to_outlook_with_oauth():
    access_token = get_oauth2_token()
    if not access_token:
        print("Geen toegangstoken verkregen.")
        return

    # Maak IMAP-verbinding
    imap_server = "outlook.office365.com"
    email_address = os.getenv('OUTLOOK_EMAIL')

    try:
        print("Connecting to IMAP server...")
        imap = imaplib.IMAP4_SSL(imap_server)
        
        # OAuth2 inlogstring opstellen
        auth_string = f"user={email_address}\x01auth=Bearer {access_token}\x01\x01"
        imap.authenticate("XOAUTH2", lambda x: auth_string)
        
        print("Connection successful!")
        imap.logout()
    except imaplib.IMAP4.error as e:
        print(f"IMAP error: {e}")
    except Exception as e:
        print(f"General error: {e}")

# Testfunctie aanroepen
connect_to_outlook_with_oauth()
