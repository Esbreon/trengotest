from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File
from io import BytesIO
import pandas as pd
import os

def get_sharepoint_data():
    """Haalt data op uit SharePoint Excel bestand"""
    try:
        print("Start ophalen SharePoint data...")
        
        # Pas de juiste SharePoint site URL aan
        site_url = "https://boostix.sharepoint.com/sites/BoostiX"
        client_id = os.environ.get('SHAREPOINT_CLIENT_ID')
        client_secret = os.environ.get('SHAREPOINT_CLIENT_SECRET')

        # Check for environment variables
        if not client_id or not client_secret:
            raise Exception("Client ID of Client Secret ontbreekt")

        print("Verbinden met SharePoint...")
        credentials = ClientCredential(client_id, client_secret)
        ctx = ClientContext(site_url).with_credentials(credentials)

        # Gebruik de juiste relatieve URL
        relative_url = "/sites/BoostiX/Gedeelde documenten/Projecten/06. Fixzed/4. Data/TestingTrengo/trengotest.xlsx"
        
        print(f"Ophalen bestand: {relative_url}")
        
        # Probeer het bestand op te halen met de juiste methode
        response = ctx.web.get_file_by_server_relative_url(relative_url).download(BytesIO()).execute_query()
        
        # Lees het bestand in een DataFrame
        bytes_file_obj = BytesIO(response.content)
        df = pd.read_excel(bytes_file_obj)
        print(f"Data opgehaald. Aantal rijen: {len(df)}")
        print("Voorbeeld van opgehaalde data:")
        print(df.head())
        return df
    
    except Exception as e:
        print(f"Fout bij ophalen SharePoint data: {str(e)}")
        raise

# Test de functie
get_sharepoint_data()
