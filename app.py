import os
import requests
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler

# Trengo configuratie
TRENGO_URL = "https://app.trengo.com/api/v2/wa_sessions"
TRENGO_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiNDk5NjQ4M2JjMDE3Y2FlYzI4ZTExYTFlN2E0NTYyODkzNmIzZTEyNzRmZDllMThmZDIxNDZkYTA1OTJiNGUyZDRiODMyMTYyZjE2OWE1NDMiLCJpYXQiOjE3Mjk0MzMwNDYsIm5iZiI6MTcyOTQzMzA0NiwiZXhwIjo0ODUzNDg0MjQ2LCJzdWIiOiI3Mjc3MzQiLCJzY29wZXMiOltdLCJhZ2VuY3lfaWQiOjMxOTc2N30.bYiBRdH_tSt3uHUSGTFANJBhSfjZo-1hRzN9SHjQf4VB4NqxsXcaFg2wZXSGlKfvMgQ10X3KG1JtbJZoDRfuUA"
TEMPLATE_ID = 181327

def get_sharepoint_data():
    """Haalt data op uit SharePoint lijst"""
    try:
        # SharePoint configuratie
        site_url = "https://boostix.sharepoint.com/s/BoostiX"
        client_id = os.environ.get('SHAREPOINT_CLIENT_ID')
        client_secret = os.environ.get('SHAREPOINT_CLIENT_SECRET')
        
        # Verbind met SharePoint
        credentials = ClientCredential(client_id, client_secret)
        ctx = ClientContext(site_url).with_credentials(credentials)
        
        # Haal lijst op (pas de lijst naam aan)
        list_title = "trengotest"
        target_list = ctx.web.lists.get_by_title(list_title)
        items = target_list.items.get().execute_query()
        
        # Converteer naar pandas DataFrame
        data = []
        for item in items:
            data.append({
                'naam': item.properties.get('Title', ''),  # SharePoint gebruikt 'Title' als standaard
                'dag': item.properties.get('Dag', ''),
                'tijdvak': item.properties.get('Tijdvak', ''),
                'telefoonnummer': item.properties.get('Telefoonnummer', '')
            })
        
        return pd.DataFrame(data)
    
    except Exception as e:
        print(f"Fout bij ophalen SharePoint data: {str(e)}")
        raise

def send_whatsapp_message(nummer, naam, dag, tijdvak):
    """Verstuurt WhatsApp bericht via Trengo"""
    payload = {
        "recipient_phone_number": nummer,
        "hsm_id": TEMPLATE_ID,
        "params": [
            {
                "type": "body",
                "key": "{{1}}",
                "value": naam
            },
            {
                "type": "body",
                "key": "{{2}}",
                "value": dag
            },
            {
                "type": "body",
                "key": "{{3}}",
                "value": tijdvak
            }
        ]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {TRENGO_TOKEN}"
    }
    
    try:
        response = requests.post(TRENGO_URL, json=payload, headers=headers)
        response.raise_for_status()
        print(f"Bericht verstuurd naar {nummer}: {response.json()}")
        return response.json()
    
    except Exception as e:
        print(f"Fout bij versturen bericht naar {nummer}: {str(e)}")
        raise

def process_data():
    """Hoofdfunctie die data ophaalt en berichten verstuurt"""
    print(f"Start verwerking: {datetime.now()}")
    
    try:
        # Haal data op
        df = get_sharepoint_data()
        print(f"Data opgehaald uit SharePoint. Aantal rijen: {len(df)}")
        
        # Verwerk elke rij
        for index, row in df.iterrows():
            try:
                nummer = row['telefoonnummer']
                # Zorg dat nummer in correct formaat staat
                if not nummer.startswith('31'):
                    nummer = '31' + nummer.lstrip('0')
                
                send_whatsapp_message(
                    nummer=nummer,
                    naam=row['naam'],
                    dag=row['dag'],
                    tijdvak=row['tijdvak']
                )
                
            except Exception as e:
                print(f"Fout bij verwerken rij {index}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Algemene fout in process_data: {str(e)}")

# Schedule de job
scheduler = BlockingScheduler()
scheduler.add_job(process_data, 'interval', minutes=30)

if __name__ == "__main__":
    print("Starting scheduler...")
    scheduler.start()
