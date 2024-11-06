import msal
import requests
import os

# Settings for OAuth2 authentication
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
TENANT_ID = os.getenv('AZURE_TENANT_ID')

def test_connection():
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
        print("Error getting token:", result.get("error_description"))
        return False

    # Test connection with /me endpoint
    headers = {
        'Authorization': f'Bearer {result["access_token"]}',
        'Content-Type': 'application/json'
    }

    try:
        # Simple test using organization endpoint
        response = requests.get(
            'https://graph.microsoft.com/v1.0/organization',
            headers=headers
        )
        
        if response.status_code == 200:
            print("Connection successful!")
            print("Organization details:", response.json())
            return True
        else:
            print(f"Connection failed with status code: {response.status_code}")
            print("Error message:", response.text)
            return False
            
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
        test_connection()
