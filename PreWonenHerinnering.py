import os
import requests
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler

# Airtable configuratie
AIRTABLE_BASE_ID = os.environ.get('AIRTABLE_BASE_ID')
AIRTABLE_TABLE_NAME = os.environ.get('AIRTABLE_TABLE_NAME')  # Bijvoorbeeld: "Projecten"
AIRTABLE_API_KEY = os.environ.get('AIRTABLE_API_KEY')

# WhatsApp template configuratie
WHATSAPP_TEMPLATE_ID = os.environ.get('WHATSAPP_TEMPLATE_ID')

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

def format_phone_number(phone):
    """Formatteert telefoonnummer naar het juiste formaat voor Trengo."""
    # Verwijder alle niet-numerieke karakters
    phone = ''.join(filter(str.isdigit, str(phone)))
    
    # Als het nummer met een 0 begint, vervang dit door 31
    if phone.startswith('0'):
        phone = '31' + phone[1:]
    # Als het nummer nog niet met 31 begint, voeg het toe
    elif not phone.startswith('31'):
        phone = '31' + phone
    
    return phone

def send_whatsapp_message(naam, monteur, dagnaam, datum, begintijd, eindtijd, reparatieduur, taaknummer, mobielnummer):
    """Verstuurt WhatsApp bericht via Trengo."""
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    # Formatteer het telefoonnummer
    formatted_phone = format_phone_number(mobielnummer)
    
    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": WHATSAPP_TEMPLATE_ID,
        "params": [
            {
                "type": "body",
                "key": "{{1}}",
                "value": str(naam)
            },
            {
                "type": "body",
                "key": "{{2}}",
                "value": str(monteur)
            },
            {
                "type": "body",
                "key": "{{3}}",
                "value": str(dagnaam)
            },
            {
                "type": "body",
                "key": "{{4}}",
                "value": str(datum)
            },
            {
                "type": "body",
                "key": "{{5}}",
                "value": str(monteur)
            },
            {
                "type": "body",
                "key": "{{6}}",
                "value": str(begintijd)
            },
            {
                "type": "body",
                "key": "{{7}}",
                "value": str(eindtijd)
            },
            {
                "type": "body",
                "key": "{{8}}",
                "value": str(reparatieduur)
            },
            {
                "type": "body",
                "key": "{{9}}",
                "value": str(taaknummer)
            },
            {
                "type": "body",
                "key": "{{10}}",
                "value": str(mobielnummer)
            }
        ]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": "Bearer " + os.environ.get('TRENGO_API_KEY')
    }
    
    try:
        print(f"Versturen bericht naar {formatted_phone} voor {naam}...")
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
                print(f"\nVerwerken rij {index + 1}: {row['fields.Naam klant']}")
                
                # Controleer of het telefoonnummer aanwezig is
                if 'fields.Mobielnummer' not in row or pd.isna(row['fields.Mobielnummer']):
                    print(f"Geen telefoonnummer gevonden voor {row['fields.Naam klant']}, deze rij wordt overgeslagen")
                    continue
                
                send_whatsapp_message(
                    naam=row['fields.Naam klant'],
                    monteur=row['fields.Monteur'],
                    dagnaam=row['fields.Dagnaam'],
                    datum=row['fields.Datum'],
                    begintijd=row['fields.Begintijd'],
                    eindtijd=row['fields.Eindtijd'],
                    reparatieduur=row['fields.Reparatieduur'],
                    taaknummer=row['fields.Taaknummer'],
                    mobielnummer=row['fields.Mobielnummer']
                )
                print(f"Bericht verstuurd voor {row['fields.Naam klant']}")
                
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
