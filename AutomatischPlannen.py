import os
import sys
import requests
import pandas as pd
from datetime import datetime
import msal
import base64

# Configuration
CUSTOM_FIELDS = {
    "plan_url": 618842,        # Original field
    "locatie": 613776,
    "element": 618192,
    "defect": 618193,
    "werkbonnummer": 618194
}

BASE_URL = "https://fixzed-a.plannen.app/token/"

# Outlook Client for Microsoft Graph API
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
        print(f"\nSearching for emails from {sender_email} with subject '{subject_line}'...")

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
                print("No new emails found")
                return None

            print("New email(s) found, checking attachments...")

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
                        print(f"Excel attachment found: {filename}")

                        content = attachment.get('contentBytes')
                        if content:
                            filepath = f"downloads/{datetime.now().strftime('%Y%m%d')}_{filename}"
                            os.makedirs('downloads', exist_ok=True)

                            print(f"Saving as: {filepath}")

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
                                print("Email marked as read")
                            except requests.exceptions.HTTPError as e:
                                print(f"Warning: Could not mark email as read: {str(e)}")

                            return filepath

            print("No Excel attachments found in new emails")
            return None

        except requests.exceptions.HTTPError as e:
            print(f"HTTP Error during API call: {str(e)}")
            if e.response is not None:
                print(f"Response body: {e.response.text}")
            raise
        except Exception as e:
            print(f"Unexpected error: {str(e)}")
            raise

# WhatsApp Messaging Logic
def send_whatsapp_message(naam_bewoner, planregel, mobielnummer, locatie, element, defect, werkbonnummer):
    email = os.environ.get('TRUSTED_EMAIL')

    if not mobielnummer:
        print(f"No valid phone number for {naam_bewoner}")
        return None

    url = "https://app.trengo.com/api/v2/wa_sessions"
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }

    template_payload = {
        "recipient_phone_number": str(mobielnummer),
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_PLAN'),
        "params": [
            {"type": "body", "key": "{{1}}", "value": str(naam_bewoner)}
        ]
    }

    try:
        print(f"Sending template message to {naam_bewoner}...")
        template_response = requests.post(url, json=template_payload, headers=headers)
        template_response.raise_for_status()
        print("Template message sent successfully")

        ticket_id = template_response.json().get('message', {}).get('ticket_id')

        if not ticket_id:
            print("No ticket ID received in template response")
            return None

        print(f"Ticket ID received: {ticket_id}")

        combined_string = f"fixzed,{email},{planregel}"
        encoded = base64.b64encode(combined_string.encode('utf-8')).decode('utf-8')
        complete_url = BASE_URL + encoded

        field_payloads = [
            (CUSTOM_FIELDS['plan_url'], complete_url),
            (CUSTOM_FIELDS['locatie'], locatie),
            (CUSTOM_FIELDS['element'], element),
            (CUSTOM_FIELDS['defect'], defect),
            (CUSTOM_FIELDS['werkbonnummer'], str(werkbonnummer))
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

        print(f"Process completed for {naam_bewoner}")
        return ticket_id
    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error sending message: {str(e)}")
        if e.response is not None:
            print(f"Response body: {e.response.text}")
        raise
    except Exception as e:
        print(f"Error sending message: {str(e)}")
        raise

# Excel File Processing
def process_excel_file(filepath):
    try:
        print(f"\nProcessing Excel file: {filepath}")
        df = pd.read_excel(filepath)

        if df.empty:
            print("No data found in Excel file")
            return

        print(f"Number of rows found: {len(df)}")
        print(f"Columns in file: {', '.join(df.columns)}")

        required_columns = ['Naam bewoner', 'Planregel', 'Mobielnummer', 'Locatie', 'Element', 'Defect', 'Werkbonnummer']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing columns in Excel: {', '.join(missing_columns)}")

        df_unique = df.drop_duplicates(subset=['Naam bewoner', 'Planregel'])

        if len(df_unique) < len(df):
            print(f"Note: {len(df) - len(df_unique)} duplicate entries removed")

        for index, row in df_unique.iterrows():
            try:
                print(f"\nProcessing row {index + 1}/{len(df_unique)}")
                mobielnummer = str(row['Mobielnummer']).strip()
                if pd.isna(mobielnummer) or not mobielnummer:
                    print(f"No valid phone number for {row['Naam bewoner']}, skipping")
                    continue

                send_whatsapp_message(
                    naam_bewoner=row['Naam bewoner'],
                    planregel=row['Planregel'],
                    mobielnummer=mobielnummer,
                    locatie=row['Locatie'],
                    element=row['Element'],
                    defect=row['Defect'],
                    werkbonnummer=row['Werkbonnummer']
                )

                print(f"Message sent for {row['Naam bewoner']}")

            except Exception as e:
                print(f"Error processing row {index}: {str(e)}")
                continue

    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        raise

# Main Execution
def process_data():
    print(f"\n=== Starting new processing: {datetime.now()} ===")

    try:
        outlook = OutlookClient()

        try:
            excel_file = outlook.download_excel_attachment(
                sender_email=os.environ.get('TEST_EMAIL'),
                subject_line=os.environ.get('SUBJECT_LINE_AUTO_PLAN')
            )

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
            print(f"Error processing emails: {str(e)}")

    except Exception as e:
        print(f"General error: {str(e)}")

if __name__ == "__main__":
    print("\n=== ENVIRONMENT CHECK ===")
    required_vars = [
        'AZURE_CLIENT_ID',
        'AZURE_CLIENT_SECRET', 
        'AZURE_TENANT_ID',
        'OUTLOOK_EMAIL', 
        'OUTLOOK_PASSWORD',
        'SENDER_EMAIL', 
        'SUBJECT_LINE_AUTO_PLAN',
        'TRENGO_API_KEY',
        'WHATSAPP_TEMPLATE_ID_PLAN',
        'TRUSTED_EMAIL'
    ]

    missing_vars = [var for var in required_vars if not os.environ.get(var)]
    if missing_vars:
        print(f"ERROR: Missing environment variables: {', '.join(missing_vars)}")
        sys.exit(1)

    print("All environment variables are set")

    print("\n=== INITIAL TEST ===")
    print("Running manual test...")
    try:
        process_data()
        print("Manual test complete")
    except Exception as e:
        print(f"Error during manual test: {str(e)}")
        sys.exit(1)
