import os
import sys
import requests
import pandas as pd
from datetime import datetime
import msal

class OutlookClient:
    def __init__(self):
        self.client_id = os.getenv('AZURE_CLIENT_ID')
        self.client_secret = os.getenv('AZURE_CLIENT_SECRET')
        self.tenant_id = os.getenv('AZURE_TENANT_ID')
        self.username = os.getenv('OUTLOOK_EMAIL')
        self.password = os.getenv('OUTLOOK_PASSWORD')
        
        # Initialize MSAL client
        self.app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )
        
    def get_token(self):
        """Get access token for Microsoft Graph API"""
        scopes = ['https://graph.microsoft.com/Mail.Read',
                 'https://graph.microsoft.com/Mail.ReadWrite',
                 'https://graph.microsoft.com/User.Read']
                 
        result = self.app.acquire_token_by_username_password(
            username=self.username,
            password=self.password,
            scopes=scopes
        )
        
        if "access_token" not in result:
            error_msg = result.get('error_description', 'Unknown error')
            print(f"Token acquisition failed. Error: {error_msg}")
            raise Exception(f"Failed to obtain token: {error_msg}")
            
        return result["access_token"]

    def verify_permissions(self, token):
        """Verify that we have the necessary permissions"""
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        # Test reading messages
        test_url = 'https://graph.microsoft.com/v1.0/me/messages?$top=1'
        response = requests.get(test_url, headers=headers)
        
        return response.status_code == 200

    def download_excel_attachment(self, sender_email, subject_line):
        """Downloads Excel attachment from specific email using Microsoft Graph API."""
        print(f"\nZoeken naar emails van {sender_email} met onderwerp '{subject_line}'...")
        
        token = self.get_token()
        
        if not self.verify_permissions(token):
            raise Exception("Insufficient permissions to access mailbox")
            
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        try:
            filter_query = f"from/emailAddress/address eq '{sender_email}' and subject eq '{subject_line}' and isRead eq false"
            url = f'https://graph.microsoft.com/v1.0/me/messages'
            params = {'$filter': filter_query, '$select': 'id,subject,hasAttachments'}
            
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()
            
            messages = response.json().get('value', [])
            
            if not messages:
                print("Geen nieuwe emails gevonden")
                return None
                
            print("Nieuwe email(s) gevonden, bijlage controleren...")
            
            for message in messages:
                if not message.get('hasAttachments'):
                    continue
                    
                try:
                    message_id = message['id']
                    attachments_url = f'https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments'
                    attachments_response = requests.get(attachments_url, headers=headers)
                    attachments_response.raise_for_status()
                    
                    attachments = attachments_response.json().get('value', [])
                    
                    for attachment in attachments:
                        filename = attachment.get('name', '')
                        if filename.endswith('.xlsx'):
                            print(f"Excel bijlage gevonden: {filename}")
                            
                            content = attachment.get('contentBytes')
                            if content:
                                filepath = f"downloads/{datetime.now().strftime('%Y%m%d')}_{filename}"
                                os.makedirs('downloads', exist_ok=True)
                                
                                print(f"Opslaan als: {filepath}")
                                
                                import base64
                                with open(filepath, 'wb') as f:
                                    f.write(base64.b64decode(content))
                                
                                return filepath
                
                except requests.exceptions.HTTPError as e:
                    print(f"Fout bij verwerken van specifieke email: {str(e)}")
                    continue
                    
            print("Geen Excel bijlage gevonden in nieuwe emails")
            return None
            
        except requests.exceptions.HTTPError as e:
            print(f"HTTP Error bij API aanroep: {str(e)}")
            if e.response is not None:
                print(f"Response body: {e.response.text}")
            raise


def send_whatsapp_message(naam, mobielnummer):
    """Sends WhatsApp message via Trengo with the template."""
    if not mobielnummer:
        print(f"Geen geldig telefoonnummer voor {naam}")
        return
        
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    payload = {
        "recipient_phone_number": mobielnummer,
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_FB_PW'),
        "params": [{"type": "body", "key": "{{1}}", "value": str(naam)}]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": "Bearer " + os.environ.get('TRENGO_API_KEY')
    }
    
    try:
        print(f"Versturen WhatsApp bericht naar {mobielnummer} voor {naam}...")
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        print(f"Trengo response: {response.text}")
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Fout bij versturen bericht: {str(e)}")
        raise


def process_excel_file(filepath):
    try:
        df = pd.read_excel(filepath)
        
        if df.empty:
            print("Geen data gevonden in Excel bestand")
            return
        
        df = df.rename(columns={'Naam bewoner': 'naam', 'Mobielnummer': 'mobielnummer'})
        
        for _, row in df.iterrows():
            send_whatsapp_message(row['naam'], row['mobielnummer'])
            
    except Exception as e:
        print(f"Fout bij verwerken Excel bestand: {str(e)}")
        raise


def process_data():
    outlook = OutlookClient()
    excel_file = outlook.download_excel_attachment(os.getenv('SENDER_EMAIL'), os.getenv('SUBJECT_LINE_PW_FB'))
    if excel_file:
        process_excel_file(excel_file)
        os.remove(excel_file)


if __name__ == "__main__":
    process_data()
