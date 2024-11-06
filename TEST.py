import msal
import requests
import os

def verify_permissions():
    """Verify Microsoft Graph API permissions and access using ROPC flow"""
    
    # Print environment variables check
    env_vars = ['AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'AZURE_TENANT_ID', 'USER_EMAIL', 'USER_PASSWORD']
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
    app = msal.PublicClientApplication(
        os.getenv('AZURE_CLIENT_ID'),
        authority=authority
    )
    
    print("\nAttempting to acquire token using ROPC flow...")
    
    # Try to acquire token using ROPC flow
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
                        print(f"  Found {len(emails)} emails in the inbox")
                else:
                    print(f"✗ {endpoint['name']}: Access denied (Status: {response.status_code})")
                    error_message = response.json().get('error', {}).get('message', 'Unknown error')
                    print(f"  Error: {error_message}")
            except Exception as e:
                print(f"✗ {endpoint['name']}: Error occurred: {str(e)}")
    else:
        print("✗ Failed to obtain token")
        print(f"Error: {result.get('error_description', 'Unknown error')}")
        
    print("\nNote: IMAP.AccessAsUser.All permission cannot be tested directly via Microsoft Graph API.")

if __name__ == "__main__":
    verify_permissions()
