import os
import sys
import requests
import pandas as pd
from datetime import datetime
import msal
import base64

# Configuration
CUSTOM_FIELD_ID = 618842  # Static Custom Field ID for Trengo
BASE_LINK_URL = "https://fixzed-a.plannen.app/token/"  # Tokenized link base

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
def send_whatsapp_message(naam_bewoner, planregel, mobielnummer):
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

        combined = f"fixzed,{email},{planregel}"
        encoded = base64.b64encode(combined.encode()).decode()
        full_url = f"{BASE_LINK_URL}{encoded}"

        custom_field_url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}/custom_fields"
        custom_field_payload = {
            "custom_field_id": CUSTOM_FIELD_ID,
            "value": full_url
        }

        print(f"Updating custom field for ticket {ticket_id}...")
        field_response = requests.post(custom_field_url, json=custom_field_payload, headers=headers)
        field_response.raise_for_status()
        print("Custom field updated successfully")

        print(f"Process completed for {naam_bewoner}")
        print(f"Details: Ticket ID={ticket_id}, Email={email}, Planregel={planregel}")

        return ticket_id
    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error sending message: {str(e)}")
        if e.response is not None:
            print(f"Response body: {e.response.text}")
        raise
    except Exception as e:
        print(f"Error sending message: {str(e)}")
        raise
