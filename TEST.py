import msal
import requests
import os

# Settings for OAuth2 authentication
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
TENANT_ID = os.getenv('AZURE_TENANT_ID')
USER_ID = "developer@boostix.nl"

def test_minimal_connection():
    # Get OAuth2 token
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    
    # For delegated permissions, we need to acquire token on behalf of user
    # First, get the authorization URL
    scopes = ['Mail.Read', 'User.Read']
    
    # Try to get token from cache first
    accounts = app.get_accounts()
    if accounts:
        # Assuming the first account is the one we want
        result = app.acquire_token_silent(scopes, account=accounts[0])
    else:
        # If no cached token, acquire token using client credentials
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" not in result:
        print("Failed to get token!")
        print("Error:", result.get("error_description"))
        return False

    print("Successfully obtained access token!")
    print("Token type:", result.get("token_type"))
    print("Expires in:", result.get("expires_in"), "seconds")
    
    # Test connection with error handling
    headers = {
        'Authorization': f'Bearer {result["access_token"]}',
        'Content-Type': 'application/json',
        'Prefer': 'outlook.body-content-type="text"'  # Request plain text body
    }
    
    try:
        # Try to access emails
        response = requests.get(
            f'https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders/inbox/messages?$top=10&$select=subject,from,receivedDateTime,bodyPreview',
            headers=headers
        )
        
        print(f"\nAPI Response Status Code: {response.status_code}")
        
        if response.status_code == 200:
            emails = response.json().get('value', [])
            if not emails:
                print("No emails found in the inbox.")
                return True
                
            print("\nDisplaying the first 10 emails in the inbox:")
            for email in emails:
                print("Subject:", email.get('subject', 'No Subject'))
                print("From:", email.get('from', {}).get('emailAddress', {}).get('address', 'Unknown Sender'))
                print("Received:", email.get('receivedDateTime', 'No Date'))
                print("Preview:", email.get('bodyPreview', 'No Preview')[:100] + '...' if email.get('bodyPreview') else 'No Preview')
                print("\n" + "="*50 + "\n")
        else:
            print("Failed to retrieve emails.")
            print("Response Status Code:", response.status_code)
            print("Response Content:", response.text)
            
            # Additional error handling for common status codes
            if response.status_code == 403:
                print("\nPermission Error: Make sure you have the following Delegated permissions in Azure AD:")
                print("- Mail.Read")
                print("- User.Read")
                print("\nPlease follow these steps:")
                print("1. Go to Azure Portal > App Registrations > Your App > API Permissions")
                print("2. Add the permissions as 'Delegated permissions'")
                print("3. Click on 'Grant admin consent' for your organization")
                print("\nAlso verify that:")
                print("- The user account has necessary mailbox permissions")
                print("- The application has been granted consent by an administrator")
            elif response.status_code == 401:
                print("\nAuthentication Error: Your token might be invalid or expired")
            
        return response.status_code == 200
        
    except requests.exceptions.RequestException as e:
        print(f"Network error occurred: {e}")
        return False
    except Exception as e:
        print(f"Unexpected error: {e}")
        return False

if __name__ == "__main__":
    # Check for required environment variables
    required_vars = ['AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'AZURE_TENANT_ID']
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print(f"Error: Missing environment variables: {', '.join(missing_vars)}")
    else:
        test_minimal_connection()
