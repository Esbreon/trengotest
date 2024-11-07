import os
import sys
import requests
import pandas as pd
from datetime import datetime
import msal
from apscheduler.schedulers.blocking import BlockingScheduler

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
        result = self.app.acquire_token_by_username_password(
            username=self.username,
            password=self.password,
            scopes=['Mail.ReadWrite', 'User.Read']  # Mail.Read vervangen door Mail.ReadWrite
        )
        
        if "access_token" not in result:
            raise Exception(f"Failed to obtain token: {result.get('error_description')}")
            
        return result["access_token"]

    def download_excel_attachment(self, sender_email, subject_line):
        """Downloads Excel attachment from specific email using Microsoft Graph API."""
        print(f"\nZoeken naar emails van {sender_email} met onderwerp '{subject_line}'...")
        
        token = self.get_token()
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
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
                            
                        # Mark message as read
                        update_url = f'https://graph.microsoft.com/v1.0/me/messages/{message_id}'
                        update_response = requests.patch(
                            update_url,
                            headers=headers,
                            json={'isRead': True}
                        )
                        update_response.raise_for_status()
                        print("Email gemarkeerd als gelezen")
                        
                        return filepath
                        
        print("Geen Excel bijlage gevonden in nieuwe emails")
        return None

def format_date(date_str):
    """Formatteert datum naar dd MMM yy formaat met Nederlandse maandafkortingen."""
    try:
        nl_month_abbr = {
            1: 'jan', 2: 'feb', 3: 'mrt', 4: 'apr', 5: 'mei', 6: 'jun',
            7: 'jul', 8: 'aug', 9: 'sep', 10: 'okt', 11: 'nov', 12: 'dec'
        }
        
        if isinstance(date_str, datetime):
            date_obj = date_str
        else:
            try:
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                try:
                    date_obj = datetime.strptime(date_str, '%d/%m/%Y')
                except ValueError:
                    date_obj = datetime.strptime(date_str, '%d-%m-%Y')
        
        day = date_obj.day
        month = date_obj.month
        year = str(date_obj.year)[2:]
        
        return f"{day} {nl_month_abbr[month]} {year}"
    
    except Exception as e:
        print(f"Fout bij formatteren datum {date_str}: {str(e)}")
        return date_str

def format_phone_number(phone):
    """Formats phone number for Trengo."""
    phone = ''.join(filter(str.isdigit, str(phone)))
    if phone.startswith('0'):
        phone = '31' + phone[1:]
    elif not phone.startswith('31'):
        phone = '31' + phone
    return phone

def send_whatsapp_message(naam_bewoner, datum, tijdvak, reparatieduur, mobielnummer):
    """Sends WhatsApp message via Trengo with the template."""
    url = "https://app.trengo.com/api/v2/wa_sessions"
    test_nummer = "+31 6 53610195"
    formatted_phone = format_phone_number(test_nummer)
    formatted_date = format_date(datum)
    
    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID'),
        "params": [
            {"type": "body", "key": "{{1}}", "value": str(naam_bewoner)},
            {"type": "body", "key": "{{2}}", "value": formatted_date},
            {"type": "body", "key": "{{3}}", "value": str(tijdvak)},
            {"type": "body", "key": "{{4}}", "value": str(reparatieduur)}
        ]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": "Bearer " + os.environ.get('TRENGO_API_KEY')
    }
    
    try:
        print(f"Versturen WhatsApp bericht naar TEST NUMMER {formatted_phone} voor {naam_bewoner}...")
        print(f"Bericht details: Datum={formatted_date}, Tijdvak={tijdvak}, Reparatieduur={reparatieduur}")
        response = requests.post(url, json=payload, headers=headers)
        print(f"Trengo response: {response.text}")
        return response.json()
    except Exception as e:
        print(f"Fout bij versturen bericht: {str(e)}")
        raise

def process_excel_file(filepath):
    """Verwerkt Excel bestand en stuurt berichten."""
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
            'Datum bezoek': 'fields.Datum bezoek',
            'Tijdvak': 'fields.Tijdvak',
            'Reparatieduur': 'fields.Reparatieduur',
            'Mobielnummer': 'fields.Mobielnummer'
        }
        
        df = df.rename(columns=column_mapping)
        
        for index, row in df.iterrows():
            try:
                print(f"\nVerwerken rij {index + 1}/{len(df)}")
                if pd.isna(row['fields.Mobielnummer']):
                    print(f"Geen telefoonnummer gevonden voor {row['fields.Naam bewoner']}, deze overslaan")
                    continue
                
                send_whatsapp_message(
                    naam_bewoner=row['fields.Naam bewoner'],
                    datum=row['fields.Datum bezoek'],
                    tijdvak=row['fields.Tijdvak'],
                    reparatieduur=row['fields.Reparatieduur'],
                    mobielnummer=row['fields.Mobielnummer']
                )
                
                print(f"Bericht verstuurd voor {row['fields.Naam bewoner']}")
                
            except Exception as e:
                print(f"Fout bij verwerken rij {index}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Fout bij verwerken Excel bestand: {str(e)}")

def process_data():
    """Main function to check email and process Excel."""
    print(f"\n=== Start nieuwe verwerking: {datetime.now()} ===")
    
    try:
        outlook = OutlookClient()
        
        try:
            excel_file = outlook.download_excel_attachment(
                sender_email=os.environ.get('SENDER_EMAIL'),
                subject_line=os.environ.get('SUBJECT_LINE')
            )
            
            if excel_file:
                process_excel_file(excel_file)
                print(f"\nVerwijderen tijdelijk bestand: {excel_file}")
                os.remove(excel_file)
            else:
                print("Geen nieuwe Excel bestanden gevonden om te verwerken")
                
        except Exception as e:
            print(f"Fout bij verwerken emails: {str(e)}")
            
    except Exception as e:
        print(f"Algemene fout: {str(e)}")

# Start script
if __name__ == "__main__":
    print("\n=== ENVIRONMENT CHECK ===")
    required_vars = [
        'AZURE_CLIENT_ID',
        'AZURE_CLIENT_SECRET', 
        'AZURE_TENANT_ID',
        'OUTLOOK_EMAIL', 
        'OUTLOOK_PASSWORD',
        'SENDER_EMAIL', 
        'SUBJECT_LINE',
        'WHATSAPP_TEMPLATE_ID',
        'TRENGO_API_KEY'
    ]
    
    missing_vars = [var for var in required_vars if not os.environ.get(var)]
    if missing_vars:
        print(f"ERROR: Missende environment variables: {', '.join(missing_vars)}")
        sys.exit(1)
    
    print("Alle environment variables zijn ingesteld")
    
    print("\n=== EERSTE TEST ===")
    print("Handmatige test uitvoeren...")
    try:
        process_data()
        print("Handmatige test compleet")
    except Exception as e:
        print(f"Fout tijdens handmatige test: {str(e)}")
    
    print("\n=== SCHEDULER STARTEN ===")
    print("Automatische verwerking elke 30 minuten wordt gestart...")
    scheduler = BlockingScheduler()
    scheduler.add_job(process_data, 'interval', minutes=30)
    
    try:
        print("Scheduler draait nu... (Ctrl+C om te stoppen)")
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        print('\nScript gestopt')
