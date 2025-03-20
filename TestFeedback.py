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


def update_custom_field(ticket_id, task_id):
    """
    Updates the custom field in a Trengo ticket with the Taskid.

    Args:
        ticket_id (str): The ID of the Trengo ticket.
        task_id (str): The Task ID (fixed as 1111 for testing).

    Returns:
        bool: True if successful, False otherwise.
    """
    custom_field_url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}/custom_fields"

    payload = {
        "custom_field_id": 618194,  
        "value": task_id  # Fixed as "1111"
    }

    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": "Bearer " + os.environ.get('TRENGO_API_KEY')
    }

    try:
        response = requests.post(custom_field_url, json=payload, headers=headers)
        response.raise_for_status()
        print(f"Custom field updated successfully with Taskid: {task_id}")
        return True
    except requests.exceptions.RequestException as e:
        print(f"Error updating custom field: {str(e)}")
        if hasattr(e, 'response') and e.response is not None:
            print(f"Response content: {e.response.text}")
        return False
        

def send_whatsapp_message():
    """Sends WhatsApp message via Trengo with fixed test values and updates the custom field."""
    
    naam = "Tris"  # Fixed name
    mobielnummer = "+31653610195"  # Fixed test number
    task_id = "1111"  # Fixed Task ID
    
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    payload = {
        "recipient_phone_number": mobielnummer,
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_FB_PW'),
        "params": [{"type": "body", "key": "{{1}}", "value": naam}]
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
        
        response_data = response.json()
        print(f"Trengo response: {response.text}")

        # Extract ticket ID and update the custom field with Taskid
        ticket_id = response_data.get("message", {}).get("ticket_id")
        if ticket_id:
            update_custom_field(ticket_id, task_id)

        return response_data
    except requests.exceptions.RequestException as e:
        print(f"Fout bij versturen bericht: {str(e)}")
        raise


def process_data():
    """Function to trigger the WhatsApp message with fixed test values."""
    send_whatsapp_message()


if __name__ == "__main__":
    process_data()
