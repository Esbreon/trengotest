import os
import sys
import base64
import requests
import pandas as pd
from datetime import datetime
import msal
import math

# Config: Trengo Custom Field IDs
CUSTOM_FIELDS = {
    "plan_url": 618842,
    "locatie": 613776,
    "element": 618192,
    "defect": 618193,
    "werkbonnummer": 618194,
    "binnen_of_buiten": 618205
}

BASE_PLAN_URL = "https://fixzed.plannen.app/token/"

# === Outlook Client ===
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
        scopes = [
            'https://graph.microsoft.com/Mail.Read',
            'https://graph.microsoft.com/Mail.ReadWrite',
            'https://graph.microsoft.com/User.Read'
        ]
        result = self.app.acquire_token_by_username_password(
            username=self.username,
            password=self.password,
            scopes=scopes
        )
        if "access_token" not in result:
            raise Exception(f"Token acquisition failed: {result.get('error_description', 'Unknown error')}")
        return result["access_token"]

    def verify_permissions(self, token):
        test_url = 'https://graph.microsoft.com/v1.0/me/messages?$top=1'
        headers = {'Authorization': f'Bearer {token}'}
        response = requests.get(test_url, headers=headers)
        return response.status_code == 200

    def download_excel_attachment(self, sender_email, subject_line):
        print(f"Searching for unread emails from '{sender_email}' with subject '{subject_line}'...")
        token = self.get_token()

        if not self.verify_permissions(token):
            raise Exception("Insufficient Microsoft Graph permissions")

        headers = {'Authorization': f'Bearer {token}'}
        params = {
            '$filter': f"from/emailAddress/address eq '{sender_email}' and subject eq '{subject_line}' and isRead eq false",
            '$select': 'id,subject,hasAttachments'
        }

        response = requests.get("https://graph.microsoft.com/v1.0/me/messages", headers=headers, params=params)
        response.raise_for_status()
        messages = response.json().get("value", [])

        for msg in messages:
            if msg.get("hasAttachments"):
                msg_id = msg["id"]
                att_url = f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}/attachments"
                att_resp = requests.get(att_url, headers=headers)
                att_resp.raise_for_status()
                for att in att_resp.json().get("value", []):
                    if att["name"].endswith(".xlsx"):
                        filepath = f"downloads/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{att['name']}"
                        os.makedirs("downloads", exist_ok=True)
                        with open(filepath, "wb") as f:
                            f.write(base64.b64decode(att["contentBytes"]))
                        requests.patch(
                            f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}",
                            headers=headers,
                            json={'isRead': True}
                        )
                        print(f"Saved: {filepath}")
                        return filepath

        print("No Excel attachments found.")
        return None

# === Helpers ===
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

def format_phone_number(phone):
    if pd.isna(phone):
        return None
    phone = str(phone).strip()
    if phone.endswith('.0'):
        phone = phone.split('.')[0]
    return phone

# === Trengo WhatsApp Sender ===
def send_whatsapp_message(
    naam_bewoner,
    planregel,
    mobielnummer,
    locatie,
    element,
    defect,
    werkbonnummer,
    binnen_of_buiten
):
    if not mobielnummer:
        print(f"No valid phone number for {naam_bewoner}")
        return None

    formatted_phone = format_phone_number(mobielnummer)
    email = os.environ.get('TRUSTED_EMAIL')
    template_id = os.environ.get('WHATSAPP_TEMPLATE_ID_PLAN')

    url = "https://app.trengo.com/api/v2/wa_sessions"
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }

    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": template_id,
        "params": [
            {"type": "body", "key": "{{1}}", "value": str(naam_bewoner)}
        ]
    }

    try:
        print(f"Sending WhatsApp message to {naam_bewoner}...")
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        response_json = response.json()
        print(f"Trengo response: {response.text}")

        ticket_id = response_json.get('message', {}).get('ticket_id')
        if not ticket_id:
            print("No ticket_id received from Trengo, cannot set custom fields.")
            return

        combined_string = f"fixzed,{email},{planregel}"
        encoded = base64.b64encode(combined_string.encode('utf-8')).decode('utf-8')
        plan_url = BASE_PLAN_URL + encoded

        field_payloads = [
            (CUSTOM_FIELDS['plan_url'], safe_str(plan_url)),
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

        print(f"Message + custom fields set for {naam_bewoner}")
        return response_json

    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error while sending message: {str(e)}")
        if e.response is not None:
            print(f"Response body: {e.response.text}")
        raise
    except Exception as e:
        print(f"Error while sending message: {str(e)}")
        raise

# === Excel Processor ===
def process_excel_file(filepath):
    try:
        print(f"\nProcessing Excel file: {filepath}")
        df = pd.read_excel(filepath)

        if df.empty:
            print("No data found in Excel file")
            return

        print(f"Number of rows found: {len(df)}")
        print(f"Columns: {', '.join(df.columns)}")

        required_columns = ['Naam bewoner', 'Planregel', 'Mobielnummer', 'Locatie', 'Element', 'Defect', 'Werkbonnummer', 'Binnen of buiten']
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing columns in Excel: {', '.join(missing)}")

        df_unique = df.drop_duplicates(subset=['Naam bewoner', 'Planregel'])

        if len(df_unique) < len(df):
            print(f"{len(df) - len(df_unique)} duplicate rows removed")

        for index, row in df_unique.iterrows():
            try:
                print(f"\nProcessing row {index + 1}/{len(df_unique)}")

                send_whatsapp_message(
                    naam_bewoner=row['Naam bewoner'],
                    planregel=row['Planregel'],
                    mobielnummer=row['Mobielnummer'],
                    locatie=row['Locatie'],
                    element=row['Element'],
                    defect=row['Defect'],
                    werkbonnummer=row['Werkbonnummer'],
                    binnen_of_buiten=row['Binnen of buiten']
                )

            except Exception as e:
                print(f"Error processing row {index + 1}: {str(e)}")
                continue

    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        raise

# === Orchestration ===
def process_data():
    print(f"\n=== Starting new processing: {datetime.now()} ===")

    try:
        outlook = OutlookClient()
        sender_email = os.environ.get('TEST_EMAIL')
        subject_line = os.environ.get('SUBJECT_LINE_AUTO_PLAN')

        if not sender_email or not subject_line:
            raise EnvironmentError("TEST_EMAIL and/or SUBJECT_LINE_AUTO_PLAN not set")

        excel_file = outlook.download_excel_attachment(sender_email, subject_line)

        if excel_file:
            try:
                process_excel_file(excel_file)
            finally:
                if os.path.exists(excel_file):
                    print(f"\nRemoving temporary file: {excel_file}")
                    os.remove(excel_file)
        else:
            print("No new Excel files found to process")

    except Exception as e:
        print(f"General error: {str(e)}")

# === Entrypoint ===
if __name__ == "__main__":
    print("\n=== ENVIRONMENT CHECK ===")
    required_vars = [
        'AZURE_CLIENT_ID',
        'AZURE_CLIENT_SECRET',
        'AZURE_TENANT_ID',
        'OUTLOOK_EMAIL',
        'OUTLOOK_PASSWORD',
        'TEST_EMAIL',
        'SUBJECT_LINE_AUTO_PLAN',
        'WHATSAPP_TEMPLATE_ID_PLAN',
        'TRENGO_API_KEY',
        'TRUSTED_EMAIL'
    ]

    missing = [var for var in required_vars if not os.environ.get(var)]
    if missing:
        print(f"ERROR: Missing environment variables: {', '.join(missing)}")
        sys.exit(1)

    print("All environment variables are set")

    print("\n=== MANUAL TEST ===")
    try:
        process_data()
        print("Manual test complete")
    except Exception as e:
        print(f"Error during manual test: {str(e)}")
        sys.exit(1)
