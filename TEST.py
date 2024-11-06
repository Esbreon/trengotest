import msal
import requests
import os

# Settings for OAuth2 authentication
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
TENANT_ID = os.getenv('AZURE_TENANT_ID')

# Define the user ID or email to access their inbox
USER_ID = "developer@boostix.nl"  # Replace with the actual user ID or email address

def test_minimal_connection():
    # Get OAuth2 token
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    
    # Request token with minimal scope
    scopes = ["https://graph.microsoft.com/.default"]
    
    result = app.acquire_token_for_client(scopes=scopes)
    
    if "access_token" not in result:
        print("Failed to get token!")
        print("Error:", result.get("error_description"))
        return False

    print("Successfully obtained access token!")
    print("Token type:", result.get("token_type"))
    print("Expires in:", result.get("expires_in"), "seconds")
    
    # Test connection by fetching the first 10 emails in the user's inbox
    headers = {
        'Authorization': f'Bearer {result["access_token"]}',
        'Content-Type': 'application/json'
    }

    try:
        # Request to get the first 10 messages in the specified user's inbox
        response = requests.get(
            f'https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders/inbox/messages?$top=10',
            headers=headers
        )
        
        print(f"\nAPI Response Status Code: {response.status_code}")
        
        if response.status_code == 200:
            emails = response.json().get('value', [])
            print("\nDisplaying the first 10 emails in the inbox:")
            for email in emails:
                print("Subject:", email.get('subject', 'No Subject'))
                print("From:", email.get('from', {}).get('emailAddress', {}).get('address', 'Unknown Sender'))
                print("Received:", email.get('receivedDateTime', 'No Date'))
                print("Preview:", email.get('bodyPreview', 'No Preview'))
                print("\n" + "="*50 + "\n")
        else:
            print("Failed to retrieve emails. Response Content:", response.text)
        
        return response.status_code == 200

    except Exception as e:
        print(f"Error testing connection: {e}")
        return False

if __name__ == "__main__":
    # Check for required environment variables
    required_vars = ['AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'AZURE_TENANT_ID']
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print(f"Error: Missing environment variables: {', '.join(missing_vars)}")
    else:
        test_minimal_connection()
