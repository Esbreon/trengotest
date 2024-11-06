import msal
import os
import requests

# Settings for OAuth2 authentication
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
TENANT_ID = os.getenv('AZURE_TENANT_ID')
OUTLOOK_EMAIL = os.getenv('OUTLOOK_EMAIL')

def get_oauth2_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    
    # Scopes for Microsoft Graph API
    # You can adjust these based on your needs
    scopes = [
        "https://graph.microsoft.com/.default"
    ]
    
    result = app.acquire_token_for_client(scopes=scopes)
    if "access_token" in result:
        print("Access token successfully obtained.")
        return result['access_token']
    else:
        print("Error getting token:", result.get("error_description"))
        return None

def get_email_messages():
    access_token = get_oauth2_token()
    if not access_token:
        print("No access token obtained, exiting...")
        return

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    # Base URL for Microsoft Graph API
    graph_api_endpoint = 'https://graph.microsoft.com/v1.0'

    try:
        # Get messages from inbox
        # You can modify this endpoint based on your needs
        messages_url = f"{graph_api_endpoint}/users/{OUTLOOK_EMAIL}/messages"
        
        # You can add query parameters for filtering, ordering, etc.
        params = {
            '$top': 10,  # Number of messages to retrieve
            '$select': 'subject,receivedDateTime,from',  # Fields to retrieve
            '$orderby': 'receivedDateTime DESC'  # Sort by date
        }
        
        response = requests.get(messages_url, headers=headers, params=params)
        response.raise_for_status()  # Raise an exception for bad status codes
        
        messages = response.json().get('value', [])
        print(f"Successfully retrieved {len(messages)} messages!")
        
        # Print message details
        for message in messages:
            print(f"Subject: {message.get('subject')}")
            print(f"From: {message.get('from', {}).get('emailAddress', {}).get('address')}")
            print(f"Received: {message.get('receivedDateTime')}")
            print("-" * 50)

    except requests.exceptions.RequestException as e:
        print(f"Error accessing Graph API: {e}")
    except Exception as e:
        print(f"General error: {e}")

if __name__ == "__main__":
    # Check for required environment variables
    required_vars = ['AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'AZURE_TENANT_ID', 'OUTLOOK_EMAIL']
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print(f"Error: Missing environment variables: {', '.join(missing_vars)}")
    else:
        get_email_messages()
