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
            raise Exception(f"Failed to obtain token: {error_msg}")
        return result["access_token"]

    def verify_permissions(self, token):
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        test_url = 'https://graph.microsoft.com/v1.0/me/messages?$top=1'
        response = requests.get(test_url, headers=headers)
        return response.status_code == 200

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
                            import base64
                            filepath = f"downloads/{datetime.now().strftime('%Y%m%d')}_{filename}"
                            os.makedirs('downloads', exist_ok=True)
                            with open(filepath, 'wb') as f:
                                f.write(base64.b64decode(content))

                            try:
                                update_url = f'https://graph.microsoft.com/v1.0/me/messages/{message_id}'
                                requests.patch(update_url, headers=headers, json={'isRead': True}).raise_for_status()
                                print("Email gemarkeerd als gelezen")
                            except:
                                print("Kon email niet als gelezen markeren")
                            return filepath

            print("Geen Excel bijlage gevonden in nieuwe emails")
            return None

        except requests.exceptions.RequestException as e:
            print(f"Fout bij ophalen van email of bijlage: {e}")
            raise

def format_date(date_str):
    try:
        if pd.isna(date_str) or str(date_str).lower() == "nat":
            return ""
        nl_month_abbr = {
            1: 'januari', 2: 'februari', 3: 'maart', 4: 'april', 5: 'mei', 6: 'juni',
            7: 'juli', 8: 'augustus', 9: 'september', 10: 'oktober', 11: 'november', 12: 'december'
        }
        if isinstance(date_str, datetime):
            date_obj = date_str
        else:
            for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
                try:
                    date_obj = datetime.strptime(date_str, fmt)
                    break
                except ValueError:
                    continue
            else:
                return str(date_str)
        return f"{date_obj.day} {nl_month_abbr[date_obj.month]} {date_obj.year}"
    except Exception as e:
        print(f"Fout bij formatteren datum {date_str}: {str(e)}")
        return str(date_str)

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

def send_whatsapp_message(naam, monteur, dagnaam, datum, tijdvak, reparatieduur, dp_nummer, mobielnummer,
                          locatie, element, defect, werkbonnummer, binnen_of_buiten):
    if not mobielnummer:
        print(f"Geen geldig telefoonnummer voor {naam}")
        return

    url = "https://app.trengo.com/api/v2/wa_sessions"
    formatted_phone = format_phone_number(mobielnummer)
    formatted_date = format_date(datum)

    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_PW_HERINNERING'),
        "params": [
            {"type": "body", "key": "{{1}}", "value": str(naam)},
            {"type": "body", "key": "{{2}}", "value": str(monteur)},
            {"type": "body", "key": "{{3}}", "value": str(dagnaam)},
            {"type": "body", "key": "{{4}}", "value": formatted_date},
            {"type": "body", "key": "{{5}}", "value": str(monteur)},
            {"type": "body", "key": "{{6}}", "value": str(tijdvak)},
            {"type": "body", "key": "{{7}}", "value": str(reparatieduur)},
            {"type": "body", "key": "{{8}}", "value": str(dp_nummer)}
        ]
    }

    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": "Bearer " + os.environ.get('TRENGO_API_KEY')
    }

    try:
        print(f"Versturen WhatsApp bericht naar {formatted_phone} voor {naam}...")
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        response_json = response.json()
        print(f"Trengo response: {response.text}")

        ticket_id = response_json.get('message', {}).get('ticket_id')
        if not ticket_id:
            print("Geen ticket_id ontvangen. Custom fields overslaan.")
            return response_json

        field_payloads = [
            (CUSTOM_FIELDS['locatie'], safe_str(locatie)),
            (CUSTOM_FIELDS['element'], safe_str(element)),
            (CUSTOM_FIELDS['defect'], safe_str(defect)),
            (CUSTOM_FIELDS['werkbonnummer'], safe_str(werkbonnummer)),
            (CUSTOM_FIELDS['binnen_of_buiten'], safe_str(binnen_of_buiten))
        ]

        for field_id, value in field_payloads:
            print(f"Custom field {field_id} = {repr(value)}")
            custom_field_url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}/custom_fields"
            custom_field_payload = {
                "custom_field_id": field_id,
                "value": value
            }
            field_response = requests.post(custom_field_url, json=custom_field_payload, headers=headers)
            field_response.raise_for_status()

        print(f"Bericht en custom fields succesvol verstuurd voor {naam}")
        return response_json

    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error bij versturen bericht: {str(e)}")
        if e.response is not None:
            print(f"Response body: {e.response.text}")
        raise
    except Exception as e:
        print(f"Fout bij verzenden WhatsApp bericht: {str(e)}")
        raise

def process_excel_file(filepath):
    print(f"\nVerwerken Excel bestand: {filepath}")
    df = pd.read_excel(filepath)

    if df.empty:
        print("Geen data gevonden in Excel bestand")
        return

    print(f"Aantal rijen gevonden: {len(df)}")
    required_columns = [
        'Naam bewoner', 'Datum bezoek', 'Reparatieduur', 'Mobielnummer',
        'Monteur', 'Dagnaam', 'DP Nummer', 'Tijdvak',
        'Locatie', 'Element', 'Defect', 'Werkbonnummer', 'Binnen of buiten'
    ]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Missende kolommen in Excel: {', '.join(missing)}")

    df['Datum bezoek'] = pd.to_datetime(df['Datum bezoek'])

    df_unique = df.drop_duplicates(subset=['Naam bewoner', 'Datum bezoek', 'DP Nummer'])
    if len(df_unique) < len(df):
        print(f"{len(df) - len(df_unique)} dubbele afspraken verwijderd")

    for index, row in df_unique.iterrows():
        try:
            print(f"\nVerwerken rij {index + 1}/{len(df_unique)}")
            mobielnummer = format_phone_number(row['Mobielnummer'])
            if not mobielnummer:
                print(f"Geen geldig telefoonnummer voor {row['Naam bewoner']}, overslaan")
                continue

            send_whatsapp_message(
                naam=row['Naam bewoner'],
                monteur=row['Monteur'],
                dagnaam=row['Dagnaam'],
                datum=row['Datum bezoek'],
                tijdvak=row['Tijdvak'],
                reparatieduur=row['Reparatieduur'],
                dp_nummer=row['DP Nummer'],
                mobielnummer=mobielnummer,
                locatie=row['Locatie'],
                element=row['Element'],
                defect=row['Defect'],
                werkbonnummer=row['Werkbonnummer'],
                binnen_of_buiten=row['Binnen of buiten']
            )
        except Exception as e:
            print(f"Fout bij verwerken rij {index + 1}: {str(e)}")
            continue

def process_data():
    print(f"\n=== Start nieuwe verwerking: {datetime.now()} ===")
    try:
        outlook = OutlookClient()
        sender_email = os.environ.get('SENDER_EMAIL')
        subject_line = os.environ.get('SUBJECT_LINE_PW_HERINNERING')

        if not sender_email or not subject_line:
            raise EnvironmentError("SENDER_EMAIL of SUBJECT_LINE_PW_HERINNERING ontbreekt in environment")

        excel_file = outlook.download_excel_attachment(sender_email, subject_line)

        if excel_file:
            try:
                process_excel_file(excel_file)
            finally:
                if os.path.exists(excel_file):
                    print(f"\nVerwijderen tijdelijk bestand: {excel_file}")
                    os.remove(excel_file)
        else:
            print("Geen nieuwe Excel bestanden gevonden")
    except Exception as e:
        print(f"Fout tijdens verwerking: {str(e)}")

if __name__ == "__main__":
    print("\n=== ENVIRONMENT CHECK ===")
    required_vars = [
        'AZURE_CLIENT_ID',
        'AZURE_CLIENT_SECRET',
        'AZURE_TENANT_ID',
        'OUTLOOK_EMAIL',
        'OUTLOOK_PASSWORD',
        'SENDER_EMAIL',
        'SUBJECT_LINE_PW_HERINNERING',
        'WHATSAPP_TEMPLATE_ID_PW_HERINNERING',
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
