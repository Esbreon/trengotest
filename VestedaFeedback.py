import os
import sys
import requests
import pandas as pd
from datetime import datetime
import msal
import pytz

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
        
        if response.status_code != 200:
            print(f"Permission verification failed. Status: {response.status_code}")
            print(f"Response: {response.text}")
            return False
        return True

    def download_excel_attachment(self, sender_email, subject_line):
        """Downloads Excel attachment from specific email using Microsoft Graph API."""
        print(f"\nZoeken naar emails van {sender_email} met onderwerp '{subject_line}'...")
        
        token = self.get_token()
        
        # Verify permissions before proceeding
        if not self.verify_permissions(token):
            raise Exception("Insufficient permissions to access mailbox")
            
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        try:
            # Search for unread emails matching criteria
            filter_query = f"from/emailAddress/address eq '{sender_email}' and subject eq '{subject_line}' and isRead eq false"
            url = f'https://graph.microsoft.com/v1.0/me/messages'
            params = {
                '$filter': filter_query,
                '$select': 'id,subject,hasAttachments'
            }
            
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
                    # Get attachments for this message
                    message_id = message['id']
                    attachments_url = f'https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments'
                    attachments_response = requests.get(attachments_url, headers=headers)
                    attachments_response.raise_for_status()
                    
                    attachments = attachments_response.json().get('value', [])
                    
                    for attachment in attachments:
                        filename = attachment.get('name', '')
                        if filename.endswith('.xlsx'):
                            print(f"Excel bijlage gevonden: {filename}")
                            
                            # Download attachment content
                            content = attachment.get('contentBytes')
                            if content:
                                filepath = f"downloads/{datetime.now().strftime('%Y%m%d')}_{filename}"
                                os.makedirs('downloads', exist_ok=True)
                                
                                print(f"Opslaan als: {filepath}")
                                
                                # Decode and save attachment
                                import base64
                                with open(filepath, 'wb') as f:
                                    f.write(base64.b64decode(content))
                                    
                                try:
                                    # Mark message as read
                                    update_url = f'https://graph.microsoft.com/v1.0/me/messages/{message_id}'
                                    update_response = requests.patch(
                                        update_url,
                                        headers=headers,
                                        json={'isRead': True}
                                    )
                                    update_response.raise_for_status()
                                    print("Email gemarkeerd als gelezen")
                                except requests.exceptions.HTTPError as e:
                                    print(f"Waarschuwing: Kon email niet als gelezen markeren: {str(e)}")
                                    # Continue processing even if marking as read fails
                                    
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
        except Exception as e:
            print(f"Onverwachte fout: {str(e)}")
            raise

def format_phone_number(phone):
    """Zorg dat telefoonnummer correct wordt geformatteerd."""
    if pd.isna(phone):
        return None  # Geen telefoonnummer
    phone = str(phone).strip()  # Strip spaties en zorg voor string
    if phone.endswith('.0'):  # Verwijder onnodige .0
        phone = phone.split('.')[0]
    return phone

def send_whatsapp_message(naam, dp_nummer, mobielnummer):
    """Sends WhatsApp message via Trengo with the template."""
    if not mobielnummer:
        print(f"Geen geldig telefoonnummer voor {naam}")
        return
        
    url = "https://app.trengo.com/api/v2/wa_sessions"
    formatted_phone = format_phone_number(mobielnummer)
    
    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_FB_VES'),
        "params": [
            {"type": "body", "key": "{{1}}", "value": str(naam)},
            {"type": "body", "key": "{{2}}", "value": str(dp_nummer)},
        ]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": "Bearer " + os.environ.get('TRENGO_API_KEY')
    }
    
    try:
        print(f"Versturen WhatsApp bericht naar {formatted_phone} voor {naam} (DP: {dp_nummer})...")
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        print(f"Trengo response: {response.text}")
        return response.json()
    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error bij versturen bericht: {str(e)}")
        if e.response is not None:
            print(f"Response body: {e.response.text}")
        raise
    except Exception as e:
        print(f"Fout bij versturen bericht: {str(e)}")
        raise

def process_excel_file(filepath):
    try:
        print(f"\nVerwerken Excel bestand: {filepath}")
        df = pd.read_excel(filepath)
        
        if df.empty:
            print("Geen data gevonden in Excel bestand")
            return
        
        print(f"Aantal rijen gevonden: {len(df)}")
        print(f"Kolommen in bestand: {', '.join(df.columns)}")
        
        column_mapping = {
            'Naam bewoner': 'fields.Naam bewoner',
            'DP Nummer': 'fields.DP Nummer', 
            'Mobielnummer': 'fields.Mobielnummer'
        }
        
        missing_columns = [col for col in column_mapping.keys() if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missende kolommen in Excel: {', '.join(missing_columns)}")
        
        df = df.rename(columns=column_mapping)
        
        df_unique = df.drop_duplicates(subset=['fields.DP Nummer'])
        
        if len(df_unique) < len(df):
            print(f"Let op: {len(df) - len(df_unique)} dubbele afspraken verwijderd")
        
        for index, row in df_unique.iterrows():
            try:
                print(f"\nVerwerken rij {index + 1}/{len(df_unique)}")
                mobielnummer = format_phone_number(row['fields.Mobielnummer'])
                if not mobielnummer:
                    print(f"Geen geldig telefoonnummer voor {row['fields.Naam bewoner']}, deze overslaan")
                    continue
                
                send_whatsapp_message(
                    naam=row['fields.Naam bewoner'],
                    dp_nummer=row['fields.DP Nummer'],
                    mobielnummer=mobielnummer
                )
                
                print(f"Bericht verstuurd voor {row['fields.Naam bewoner']} (DP: {row['fields.DP Nummer']})")
                
            except Exception as e:
                print(f"Fout bij verwerken rij {index}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Fout bij verwerken Excel bestand: {str(e)}")
        raise

def process_data():
    """Main function to check email and process Excel."""
    print(f"\n=== Start nieuwe verwerking: {datetime.now()} ===")
    
    try:
        outlook = OutlookClient()
        
        try:
            excel_file = outlook.download_excel_attachment(
                sender_email=os.environ.get('SENDER_EMAIL'),
                subject_line=os.environ.get('SUBJECT_LINE_VES_FB')
            )
            
            if excel_file:
                try:
                    process_excel_file(excel_file)
                finally:
                    # Always try to clean up the file
                    if os.path.exists(excel_file):
                        print(f"\nVerwijderen tijdelijk bestand: {excel_file}")
                        os.remove(excel_file)
            else:
                print("Geen nieuwe Excel bestanden gevonden om te verwerken")
                
        except Exception as e:
            print(f"Fout bij verwerken emails: {str(e)}")
            
    except Exception as e:
        print(f"Algemene fout: {str(e)}")

if __name__ == "__main__":
    process_data()
