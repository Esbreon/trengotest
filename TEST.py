import msal
import imaplib
import os

# Instellingen voor OAuth2-authenticatie
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')        # Client ID van de app-registratie in Azure
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET') # Client Secret van de app-registratie in Azure
TENANT_ID = os.getenv('AZURE_TENANT_ID')         # Tenant ID van je Azure AD
OUTLOOK_EMAIL = os.getenv('OUTLOOK_EMAIL')       # Het e-mailadres waarvoor je toegang wilt

# Functie om OAuth2-toegangstoken te verkrijgen
def get_oauth2_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    
    # De scope specificeert de toegangsniveaus, hier voor de IMAP-service van Outlook
    scope = ["https://outlook.office365.com/.default"]
    
    # Token ophalen met de `acquire_token_for_client` functie
    result = app.acquire_token_for_client(scopes=scope)
    if "access_token" in result:
        print("Access token succesvol verkregen.")
        return result['access_token']
    else:
        print("Error bij verkrijgen van token:", result.get("error_description"))
        return None

# Functie om verbinding te maken met de IMAP-server via OAuth2
def connect_to_outlook_with_oauth():
    access_token = get_oauth2_token()
    if not access_token:
        print("Geen toegangstoken verkregen, afsluiten...")
        return

    # IMAP-serverinstellingen
    imap_server = "outlook.office365.com"
    
    try:
        print("Verbinding maken met de IMAP-server...")
        imap = imaplib.IMAP4_SSL(imap_server)
        
        # Opbouw van de OAuth2-authenticatiestring
        auth_string = f"user={OUTLOOK_EMAIL}\x01auth=Bearer {access_token}\x01\x01"
        imap.authenticate("XOAUTH2", lambda x: auth_string)
        
        print("Succesvol verbonden met de IMAP-server!")
        
        # Bijvoorbeeld: mappen ophalen om de verbinding te testen
        status, mailboxes = imap.list()
        if status == "OK":
            print("Beschikbare mappen:")
            for mailbox in mailboxes:
                print(mailbox.decode())
        
        imap.logout()
        
    except imaplib.IMAP4.error as e:
        print(f"IMAP fout: {e}")
    except Exception as e:
        print(f"Algemene fout: {e}")

# Hoofdfunctie aanroepen
if __name__ == "__main__":
    # Zorg dat de benodigde omgevingsvariabelen aanwezig zijn
    required_vars = ['AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'AZURE_TENANT_ID', 'OUTLOOK_EMAIL']
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print(f"Fout: Missende omgevingsvariabelen: {', '.join(missing_vars)}")
    else:
        connect_to_outlook_with_oauth()
