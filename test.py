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
            raise Exception(f"Failed to obtain token: {result.get('error_description', 'Unknown error')}")
        return result["access_token"]

    def verify_permissions(self, token):
        headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
        test_url = 'https://graph.microsoft.com/v1.0/me/messages?$top=1'
        response = requests.get(test_url, headers=headers)
        return response.status_code == 200

    def download_excel_attachment(self, sender_email, subject_line):
        print(f"Zoeken naar emails van {sender_email} met onderwerp '{subject_line}'...")
        token = self.get_token()
        if not self.verify_permissions(token):
            raise Exception("Insufficient permissions to access mailbox")

        headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
        filter_query = f"from/emailAddress/address eq '{sender_email}' and subject eq '{subject_line}' and isRead eq false"
        url = f'https://graph.microsoft.com/v1.0/me/messages'
        params = {'$filter': filter_query, '$select': 'id,subject,hasAttachments'}

        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        messages = response.json().get('value', [])
        if not messages:
            print("Geen nieuwe emails gevonden")
            return None

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
                        update_url = f'https://graph.microsoft.com/v1.0/me/messages/{message_id}'
                        requests.patch(update_url, headers=headers, json={'isRead': True})
                        return filepath
        return None

def format_date(date_str):
    try:
        nl_month_abbr = {1: 'januari', 2: 'februari', 3: 'maart', 4: 'april', 5: 'mei', 6: 'juni',
                         7: 'juli', 8: 'augustus', 9: 'september', 10: 'oktober', 11: 'november', 12: 'december'}
        date_obj = date_str if isinstance(date_str, datetime) else datetime.strptime(str(date_str), '%Y-%m-%d')
        return f"{date_obj.day} {nl_month_abbr[date_obj.month]} {date_obj.year}"
    except Exception as e:
        print(f"Fout bij formatteren datum {date_str}: {str(e)}")
        return date_str

def format_phone_number(phone):
    if pd.isna(phone): return None
    phone = str(phone).strip()
    if phone.endswith('.0'): phone = phone.split('.')[0]
    return phone

def get_all_contacts():
    url = "https://app.trengo.com/api/v2/contacts"
    headers = {
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}",
        "Accept": "application/json"
    }

    contacts = []
    page = 1

    try:
        while True:
            response = requests.get(url, headers=headers, params={"page": page})
            response.raise_for_status()
            data = response.json()

            contacts.extend(data.get("data", []))

            if not data.get("meta", {}).get("has_next"):
                break
            page += 1

    except requests.exceptions.RequestException as e:
        print(f"Fout bij ophalen contactenlijst: {str(e)}")

    return contacts


def create_contact(name, phone_number):
    url = "https://app.trengo.com/api/v2/contacts"
    headers = {
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}",
        "Content-Type": "application/json"
    }
    payload = {"name": name, "phone": phone_number}
    try:
        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()
        print(f"Nieuw contact aangemaakt voor {phone_number}")
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Fout bij aanmaken contact: {str(e)}")
        return None

def ensure_contact_exists(name, phone_number):
    all_contacts = get_all_contacts()

    normalized_input = phone_number.replace(" ", "").replace("+", "").lstrip("0")

    for contact in all_contacts:
        existing = contact.get("phone", "")
        normalized_existing = existing.replace(" ", "").replace("+", "").lstrip("0")
        if normalized_existing.endswith(normalized_input):
            print(f"Contact bestaat al voor {phone_number}")
            return contact

    # No match found → create new
    print(f"Contact nog niet gevonden, aanmaken: {phone_number}")
    return create_contact(name, phone_number)


def send_whatsapp_message(naam_bewoner, dag, datum, tijdvak, reparatieduur, dp_nummer, mobielnummer):
    if not mobielnummer:
        print(f"Geen geldig telefoonnummer voor {naam_bewoner}")
        return

    formatted_phone = format_phone_number(mobielnummer)
    ensure_contact_exists(naam_bewoner, formatted_phone)

    url = "https://app.trengo.com/api/v2/wa_sessions"
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
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    try:
        print(f"Versturen WhatsApp bericht naar {formatted_phone} voor {naam_bewoner}...")
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        print(f"Trengo response: {response.text}")
        return response.json()
    except Exception as e:
        print(f"Fout bij versturen bericht: {str(e)}")
        raise

def process_excel_file(filepath):
    print(f"Verwerken Excel bestand: {filepath}")
    df = pd.read_excel(filepath)
    if df.empty:
        print("Geen data gevonden in Excel bestand")
        return

    column_mapping = {
        'Naam bewoner': 'fields.Naam bewoner',
        'Dag': 'fields.Dag',
        'Datum bezoek': 'fields.Datum bezoek',
        'Tijdvak': 'fields.Tijdvak',
        'Reparatieduur': 'fields.Reparatieduur',
        'Mobielnummer': 'fields.Mobielnummer',
        'DP Nummer': 'fields.DP Nummer'
    }

    missing_columns = [col for col in column_mapping if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missende kolommen in Excel: {', '.join(missing_columns)}")

    df = df.rename(columns=column_mapping)
    df['fields.Datum bezoek'] = pd.to_datetime(df['fields.Datum bezoek'])
    df_unique = df.drop_duplicates(subset=['fields.Naam bewoner', 'fields.Datum bezoek', 'fields.DP Nummer'])

    for index, row in df_unique.iterrows():
        try:
            mobielnummer = format_phone_number(row['fields.Mobielnummer'])
            if not mobielnummer:
                print(f"Geen geldig nummer voor {row['fields.Naam bewoner']}, overslaan")
                continue

            send_whatsapp_message(
                naam_bewoner=row['fields.Naam bewoner'],
                datum=row['fields.Datum bezoek'],
                dag=row['fields.Dag'],
                tijdvak=row['fields.Tijdvak'],
                reparatieduur=row['fields.Reparatieduur'],
                dp_nummer=row['fields.DP Nummer'],
                mobielnummer=mobielnummer
            )
        except Exception as e:
            print(f"Fout bij verwerken rij {index}: {str(e)}")

def process_data():
    print(f"=== Start nieuwe verwerking: {datetime.now()} ===")
    outlook = OutlookClient()
    try:
        excel_file = outlook.download_excel_attachment(
            sender_email=os.environ.get('TEST_EMAIL'),
            subject_line=os.environ.get('SUBJECT_LINE_PW_BEVESTIGING')
        )
        if excel_file:
            process_excel_file(excel_file)
            if os.path.exists(excel_file):
                os.remove(excel_file)
        else:
            print("Geen nieuwe Excel bestanden gevonden om te verwerken")
    except Exception as e:
        print(f"Fout bij verwerken emails: {str(e)}")

if __name__ == "__main__":
    required_vars = [
        'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'AZURE_TENANT_ID',
        'OUTLOOK_EMAIL', 'OUTLOOK_PASSWORD',
        'TEST_EMAIL', 'SUBJECT_LINE_PW_BEVESTIGING',
        'WHATSAPP_TEMPLATE_ID_PW_BEVESTIGING', 'TRENGO_API_KEY'
    ]
    missing_vars = [var for var in required_vars if not os.environ.get(var)]
    if missing_vars:
        print(f"ERROR: Missende environment variables: {', '.join(missing_vars)}")
        sys.exit(1)
    process_data()
