import os
import requests
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler

# Airtable configuratie
AIRTABLE_BASE_ID = os.environ.get('AIRTABLE_BASE_ID')
AIRTABLE_TABLE_NAME = os.environ.get('AIRTABLE_TABLE_NAME')  # Bijvoorbeeld: "Projecten"
AIRTABLE_API_KEY = os.environ.get('AIRTABLE_API_KEY')

def get_airtable_data():
    """Haalt data op uit Airtable."""
    try:
        print("Start ophalen Airtable data...")
        
        url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_NAME}"
        
        headers = {
            "Authorization": f"Bearer {AIRTABLE_API_KEY}",
            "Content-Type": "application/json"
        }
        
        print("Verbinden met Airtable...")
        
        response = requests.get(url, headers=headers)
        
        if response.status_code != 200:
            raise Exception(f"Fout bij ophalen data: {response.status_code} - {response.text}")

        data = response.json()
        
        # Converteer records naar een DataFrame
        df = pd.json_normalize(data['records'])
        
        print(f"Data opgehaald. Aantal rijen: {len(df)}")
        print("Voorbeeld van opgehaalde data:")
        print(df.head())
        return df
    
    except Exception as e:
        print(f"Fout bij ophalen Airtable data: {str(e)}")
        raise

def send_whatsapp_message(naam, dag, tijdvak):
    """Verstuurt WhatsApp bericht via Trengo."""
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
                "value": str(fields.naam)
            },
            {
                "type": "body",
                "key": "{{2}}",
                "value": str(fields.dag)
            },
            {
                "type": "body",
                "key": "{{3}}",
                "value": str(fields.tijdvak)
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
    """Hoofdfunctie die data ophaalt en berichten verstuurt."""
    print(f"\n=== Start nieuwe verwerking: {datetime.now()} ===")
    
    try:
        # Haal data op
        df = get_airtable_data()
        
        if df.empty:
            print("Geen data gevonden om te verwerken")
            return
        
        # Verwerk elke rij
        for index, row in df.iterrows():
            try:
                print(f"\nVerwerken rij {index + 1}: {row['fields']['naam']}")
                send_whatsapp_message(
                    naam=row['fields']['naam'],
                    dag=row['fields']['dag'],
                    tijdvak=row['fields']['tijdvak']
                )
                print(f"Bericht verstuurd voor {row['fields']['naam']}")
                
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
