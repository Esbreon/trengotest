import msal
import requests
import os

def verify_permissions():
    """Verify Microsoft Graph API permissions and access"""
    
    # Print environment variables (without showing actual values)
    env_vars = ['AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'AZURE_TENANT_ID']
    print("\nEnvironment Variables Check:")
    for var in env_vars:
        value = os.getenv(var)
        if value:
            masked_value = value[:4] + '*' * (len(value) - 4)
            print(f"✓ {var} is set (value: {masked_value})")
        else:
            print(f"✗ {var} is not set")
    
    # Get OAuth2 token
    authority = f"https://login.microsoftonline.com/{os.getenv('AZURE_TENANT_ID')}"
    app = msal.ConfidentialClientApplication(
        os.getenv('AZURE_CLIENT_ID'),
        authority=authority,
        client_credential=os.getenv('AZURE_CLIENT_SECRET')
    )
    
    print("\nTesting permissions:")
    
    # Test getting token with .default scope
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" in result:
        print("✓ Successfully obtained access token")
        token = result["access_token"]
        
        # Test endpoints based on granted permissions
        endpoints = [
            {
                'url': f'https://graph.microsoft.com/v1.0/users/{os.getenv("USER_ID", "developer@boostix.nl")}/messages',
                'name': 'Mail Access (Mail.Read)',
                'permission': 'Mail.Read'
            },
            {
                'url': f'https://graph.microsoft.com/v1.0/users/{os.getenv("USER_ID", "developer@boostix.nl")}',
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
                else:
                    print(f"✗ {endpoint['name']}: Access denied (Status: {response.status_code})")
                    error_message = response.json().get('error', {}).get('message', 'Unknown error')
                    print(f"  Error: {error_message}")
                    print("\nTroubleshooting tips:")
                    print("1. Verify the permission is granted as 'Delegated' in Azure Portal")
                    print("2. Check if admin consent has been granted")
                    print("3. Verify the user account has the necessary mailbox permissions")
            except Exception as e:
                print(f"✗ {endpoint['name']}: Error occurred: {str(e)}")
    else:
        print(f"✗ Failed to obtain token")
        print(f"Error: {result.get('error_description', 'Unknown error')}")
        
    # Note about IMAP access
    print("\nNote: IMAP.AccessAsUser.All permission cannot be tested directly via Microsoft Graph API.")
    print("This permission is used for IMAP protocol access, not REST API access.")

if __name__ == "__main__":
    verify_permissions()
