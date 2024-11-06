import msal
import requests
import os

def verify_permissions():
    """Verify Microsoft Graph API permissions and access"""
    
    # Print environment variables check
    env_vars = ['AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'AZURE_TENANT_ID', 'OUTLOOK_EMAIL', 'OUTLOOK_PASSWORD']
    print("\nEnvironment Variables Check:")
    for var in env_vars:
        value = os.getenv(var)
        if value:
            masked_value = value[:4] + '*' * (len(value) - 4)
            print(f"✓ {var} is set (value: {masked_value})")
        else:
            print(f"✗ {var} is not set")
    
    # Initialize MSAL client
    authority = f"https://login.microsoftonline.com/{os.getenv('AZURE_TENANT_ID')}"
    app = msal.ConfidentialClientApplication(
        client_id=os.getenv('AZURE_CLIENT_ID'),
        client_credential=os.getenv('AZURE_CLIENT_SECRET'),
        authority=authority
    )
    
    print("\nAttempting to acquire token...")
    
    # First try to get token silently
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(['Mail.Read', 'User.Read'], account=accounts[0])
    
    if not result:
        # Try to get token with username/password
        result = app.acquire_token_by_username_password(
            username=os.getenv('OUTLOOK_EMAIL'),
            password=os.getenv('OUTLOOK_PASSWORD'),
            scopes=['Mail.Read', 'User.Read']
        )
    
    if "access_token" in result:
        print("✓ Successfully obtained access token")
        token = result["access_token"]
        
        # Test endpoints
        endpoints = [
            {
                'url': 'https://graph.microsoft.com/v1.0/me/messages',
                'name': 'Mail Access (Mail.Read)',
                'permission': 'Mail.Read'
            },
            {
                'url': 'https://graph.microsoft.com/v1.0/me',
                'name': 'User Profile (User.Read)',
                'permission': 'User.Read'
            }
        ]
        
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        print("\nTesting API endpoints:")
        for endpoint in endpoints:
            try:
                response = requests.get(endpoint['url'], headers=headers)
                if response.status_code == 200:
                    print(f"✓ {endpoint['name']}: Access granted")
                    if endpoint['name'] == 'Mail Access (Mail.Read)':
                        emails = response.json().get('value', [])
                        print(f"  Found {len(emails)} emails")
                else:
                    print(f"✗ {endpoint['name']}: Access denied (Status: {response.status_code})")
                    error_message = response.json().get('error', {}).get('message', 'Unknown error')
                    print(f"  Error: {error_message}")
                    print("\nTroubleshooting suggestions:")
                    print("1. Double-check the email and password are correct")
                    print("2. Verify MFA is not required for this account")
                    print("3. Check if the app has admin consent for these permissions")
            except Exception as e:
                print(f"✗ {endpoint['name']}: Error occurred: {str(e)}")
    else:
        print("✗ Failed to obtain token")
        print(f"Error: {result.get('error_description', 'Unknown error')}")
        print("\nTroubleshooting suggestions:")
        print("1. Verify your Azure AD app registration settings:")
        print("   - Check if the client ID and secret are correct")
        print("   - Verify the app has the correct redirect URIs")
        print("   - Make sure the app has been granted admin consent")
        print("2. Check if the user account:")
        print("   - Has MFA disabled (MFA will prevent this authentication method)")
        print("   - Has the correct licenses assigned")
        print("   - Is not blocked or restricted")
        
    print("\nNote: IMAP.AccessAsUser.All permission is for IMAP protocol access only.")

if __name__ == "__main__":
    verify_permissions()
