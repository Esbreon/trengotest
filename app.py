import os
import requests
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
from io import BytesIO

def get_sharepoint_data():
    """Haalt data op uit SharePoint Excel bestand"""
    try:
        print("Start ophalen SharePoint data...")
        
        # Pas de juiste SharePoint site URL aan
        site_url = "https://boostix.sharepoint.com/sites/BoostiX"
        client_id = os.environ.get('SHAREPOINT_CLIENT_ID')
        client_secret = os.environ.get('SHAREPOINT_CLIENT_SECRET')
        print(client_id)
        print(client_secret)
        print("Verbinden met SharePoint...")
        credentials = ClientCredential(client_id, client_secret)
        ctx = ClientContext(site_url).with_credentials(credentials)
        
        # Relatieve URL van het bestand (pas dit pad aan op basis van je SharePoint structuur)
        relative_url = "/sites/BoostiX/Gedeelde%20documenten/Projecten/06.%20Fixzed/4.%20Data/TestingTrengo/trengotest.xlsx"
        
        print(f"Ophalen bestand: {relative_url}")
        
        # Haal het bestand op
        response = File.open_binary(ctx, relative_url)
        bytes_file_obj = BytesIO()
        bytes_file_obj.write(response.content)
        bytes_file_obj.seek(0)
        
        df = pd.read_excel(bytes_file_obj)
        print(f"Data opgehaald. Aantal rijen: {len(df)}")
        print("Voorbeeld van opgehaalde data:")
        print(df.head())
        return df
    
    except Exception as e:
        print(f"Fout bij ophalen SharePoint data: {str(e)}")
        print(f"Client ID aanwezig: {'Ja' if client_id else 'Nee'}")
        print(f"Client Secret aanwezig: {'Ja' if client_secret else 'Nee'}")
        raise

def send_whatsapp_message(naam, dag, tijdvak):
    """Verstuurt WhatsApp bericht via Trengo"""
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    # Bijgewerkt telefoonnummer
    PHONE_NUMBER = "31653610195"
    
    payload = {
        "recipient_phone_number": PHONE_NUMBER,
        "hsm_id": 181327,
        "params": [
            {
                "type": "body",
                "key": "{{1}}",
                "value": str(naam)
            },
            {
                "type": "body",
                "key": "{{2}}",
                "value": str(dag)
            },
            {
                "type": "body",
                "key": "{{3}}",
                "value": str(tijdvak)
            }
        ]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": "Bearer " + os.environ.get('TRENGO_API_KEY')
    }
    
    try:
        print(f"Versturen bericht naar {PHONE_NUMBER} voor {naam}...")
        response = requests.post(url, json=payload, headers=headers)
        print(f"Response van Trengo: {response.text}")
        return response.json()
    
    except Exception as e:
        print(f"Fout bij versturen bericht: {str(e)}")
        raise

def process_data():
    """Hoofdfunctie die data ophaalt en berichten verstuurt"""
    print(f"\n=== Start nieuwe verwerking: {datetime.now()} ===")
    
    try:
        # Haal data op
        df = get_sharepoint_data()
        
        if df.empty:
            print("Geen data gevonden om te verwerken")
            return
        
        # Verwerk elke rij
        for index, row in df.iterrows():
            try:
                print(f"\nVerwerken rij {index + 1}: {row['naam']}")
                send_whatsapp_message(
                    naam=row['naam'],
                    dag=row['dag'],
                    tijdvak=row['tijdvak']
                )
                print(f"Bericht verstuurd voor {row['naam']}")
                
            except Exception as e:
                print(f"Fout bij verwerken rij {index}: {str(e)}")
                continue
    
    except Exception as e:
        print(f"Algemene fout: {str(e)}")

# Start direct één verwerking
print("Start eerste verwerking...")
process_data()

# Schedule volgende verwerkingen
scheduler = BlockingScheduler()
scheduler.add_job(process_data, 'interval', minutes=30)

if __name__ == "__main__":
    print("\nStarting scheduler...")
    scheduler.start()
