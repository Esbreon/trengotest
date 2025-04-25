import os
import sys
import requests
import pandas as pd
from datetime import datetime
import msal
from apscheduler.schedulers.blocking import BlockingScheduler
import math

CUSTOM_FIELDS = {
    "locatie": 613776,
    "element": 618192,
    "defect": 618193,
    "werkbonnummer": 618194,
    "binnen_of_buiten": 618205
}

class OutlookClient:
    def __init__(self):
        self.client_id = os.getenv('AZURE_CLIENT_ID')
        self.client_secret = os.getenv('AZURE_CLIENT_SECRET')
        self.tenant_id = os.getenv('AZURE_TENANT_ID')
        self.username = os.getenv('OUTLOOK_EMAIL')
        self.password = os.getenv('OUTLOOK_PASSWORD')

        self.app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )

    def get_token(self):
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
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }

        test_url = 'https://graph.microsoft.com/v1.0/me/messages?$top=1'
        response = requests.get(test_url, headers=headers)

        if response.status_code != 200:
            print(f"Permission verification failed. Status: {response.status_code}")
            print(f"Response: {response.text}")
            return False
        return True

    def download_excel_attachment(self, sender_email, subject_line):
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

                            import base64
                            with open(filepath, 'wb') as f:
                                f.write(base64.b64decode(content))

                            try:
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

                            return filepath

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

def format_date(date_str):
    try:
        nl_month_abbr = {
            1: 'januari', 2: 'februari', 3: 'maart', 4: 'april', 5: 'mei', 6: 'juni',
            7: 'juli', 8: 'augustus', 9: 'september', 10: 'oktober', 11: 'november', 12: 'december'
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
        return f"{date_obj.day} {nl_month_abbr[date_obj.month]} {date_obj.year}"
    except Exception as e:
        print(f"Fout bij formatteren datum {date_str}: {str(e)}")
        return date_str

def format_phone_number(phone):
    if pd.isna(phone):
        return None
    phone = str(phone).strip()
    if phone.endswith('.0'):
        phone = phone.split('.')[0]
    return phone

def safe_str(val):
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return ""
        if val.is_integer():
            return str(int(val))
    if str(val).strip().lower() in {"nan", "inf", "none"}:
        return ""
    return str(val).strip()

def send_whatsapp_message(naam_bewoner, dag, datum, tijdvak, reparatieduur, dp_nummer, mobielnummer, locatie, element, defect, werkbonnummer, binnen_of_buiten):
    if not mobielnummer:
        print(f"Geen geldig telefoonnummer voor {naam_bewoner}")
        return

    url = "https://app.trengo.com/api/v2/wa_sessions"
    formatted_phone = format_phone_number(mobielnummer)
    formatted_date = format_date(datum)

    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_PW_BEVESTIGING'),
        "params": [
            {"type": "body", "key": "{{1}}", "value": str(naam_bewoner)},
            {"type": "body", "key": "{{2}}", "value": str(dag)},
            {"type": "body", "key": "{{3}}", "value": formatted_date},
            {"type": "body", "key": "{{4}}", "value": str(tijdvak)},
            {"type": "body", "key": "{{5}}", "value": str(reparatieduur)},
            {"type": "body", "key": "{{6}}", "value": str(dp_nummer)}
        ]
    }

    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": "Bearer " + os.environ.get('TRENGO_API_KEY')
    }

    try:
        print(f"Versturen WhatsApp bericht naar {formatted_phone} voor {naam_bewoner}...")
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        response_json = response.json()
        print(f"Trengo response: {response.text}")

        ticket_id = response_json.get('message', {}).get('ticket_id')
        if not ticket_id:
            print("Geen ticket_id ontvangen van Trengo, kan custom fields niet instellen.")
            return

        field_payloads = [
            (CUSTOM_FIELDS['locatie'], safe_str(locatie)),
            (CUSTOM_FIELDS['element'], safe_str(element)),
            (CUSTOM_FIELDS['defect'], safe_str(defect)),
            (CUSTOM_FIELDS['werkbonnummer'], safe_str(werkbonnummer)),
            (CUSTOM_FIELDS['binnen_of_buiten'], safe_str(binnen_of_buiten))
        ]

        for field_id, value in field_payloads:
            custom_field_url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}/custom_fields"
            custom_field_payload = {
                "custom_field_id": field_id,
                "value": value
            }
            print(f"Updating custom field {field_id}...")
            field_response = requests.post(custom_field_url, json=custom_field_payload, headers=headers)
            field_response.raise_for_status()

        print(f"Bericht + custom fields ingesteld voor {naam_bewoner}")
        return response_json

    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error bij versturen bericht: {str(e)}")
        if e.response is not None:
            print(f"Response body: {e.response.text}")
        raise
    except Exception as e:
        print(f"Fout bij versturen bericht: {str(e)}")
        raise

def process_excel_file(filepath):
    print(f"\nVerwerken Excel bestand: {filepath}")
    df = pd.read_excel(filepath)
    if df.empty:
        print("Geen data gevonden in Excel bestand")
        return
    print(f"Aantal rijen gevonden: {len(df)}")
    for index, row in df.iterrows():
        try:
            print(f"\nVerwerken rij {index + 1}/{len(df)}")
            send_whatsapp_message(
                naam_bewoner=row['Naam bewoner'],
                dag=row['Dag'],
                datum=row['Datum bezoek'],
                tijdvak=row['Tijdvak'],
                reparatieduur=row['Reparatieduur'],
                dp_nummer=row['DP Nummer'],
                mobielnummer=row['Mobielnummer'],
                locatie=row['Locatie'],
                element=row['Element'],
                defect=row['Defect'],
                werkbonnummer=row['Werkbonnummer'],
                binnen_of_buiten=row['Binnen of buiten']
            )
        except Exception as e:
            print(f"Fout bij verwerken rij {index}: {str(e)}")
            continue

def process_data():
    print(f"\n=== Start nieuwe verwerking: {datetime.now()} ===")
    try:
        outlook = OutlookClient()
        sender_email = os.environ.get('SENDER_EMAIL')
        subject_line = os.environ.get('SUBJECT_LINE_PW_BEVESTIGING')
        if not sender_email or not subject_line:
            raise EnvironmentError("SENDER_EMAIL en/of SUBJECT_LINE_PW_BEVESTIGING niet ingesteld in environment")
        try:
            excel_file = outlook.download_excel_attachment(
                sender_email=sender_email,
                subject_line=subject_line
            )
            if excel_file:
                try:
                    process_excel_file(excel_file)
                finally:
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
    print("\n=== ENVIRONMENT CHECK ===")
    required_vars = [
        'AZURE_CLIENT_ID', 
        'AZURE_CLIENT_SECRET', 
        'AZURE_TENANT_ID',
        'OUTLOOK_EMAIL', 
        'OUTLOOK_PASSWORD',
        'SENDER_EMAIL', 
        'SUBJECT_LINE_PW_BEVESTIGING',
        'WHATSAPP_TEMPLATE_ID_PW_BEVESTIGING', 
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
        sys.exit(1)
