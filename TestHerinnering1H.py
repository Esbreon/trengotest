import os
import requests
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler

# Airtable configuratie
AIRTABLE_BASE_ID = os.environ.get('AIRTABLE_BASE_ID')
AIRTABLE_TABLE_NAME = os.environ.get('AIRTABLE_T1H')
AIRTABLE_API_KEY = os.environ.get('AIRTABLE_API_KEY')
WHATSAPP_TEMPLATE_ID = os.environ.get('WHATSAPP_TEMPLATE_ID_G_TEST')

def format_date(date_str):
    """Formatteert datum naar dd MMM yy formaat met Nederlandse maandafkortingen."""
    try:
        nl_month_abbr = {
            1: 'jan', 2: 'feb', 3: 'mrt', 4: 'apr', 5: 'mei', 6: 'jun',
            7: 'jul', 8: 'aug', 9: 'sep', 10: 'okt', 11: 'nov', 12: 'dec'
        }
        
        if isinstance(date_str, datetime):
            date_obj = date_str
        else:
            try:
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                try:
                    date_obj = datetime.strptime(date_str, '%d/%m/%Y')
                except ValueError:
                    date_obj = datetime.strptime(date_str, '%d-%m-%Y')
        
        day = date_obj.day
        month = date_obj.month
        year = str(date_obj.year)[2:]
        
        return f"{day} {nl_month_abbr[month]} {year}"
    
    except Exception as e:
        print(f"Fout bij formatteren datum {date_str}: {str(e)}")
        return date_str

def delete_airtable_record(record_id):
    """Deletes a record from Airtable."""
    url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_NAME}/{record_id}"
    headers = {
        "Authorization": f"Bearer {AIRTABLE_API_KEY}",
        "Content-Type": "application/json"
    }
    
    response = requests.delete(url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Fout bij verwijderen record: {response.status_code} - {response.text}")
    print(f"Record {record_id} succesvol verwijderd")

def get_airtable_data():
    """Haalt data op uit Airtable."""
    try:
        print("Start ophalen Airtable data...")
        
        url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_NAME}"
        headers = {
            "Authorization": f"Bearer {AIRTABLE_API_KEY}",
            "Content-Type": "application/json"
        }
        
        response = requests.get(url, headers=headers)
        
        if response.status_code != 200:
            raise Exception(f"Fout bij ophalen data: {response.status_code} - {response.text}")
        data = response.json()
        
        df = pd.json_normalize(data['records'])
        df['fields.Taaknummer'] = df['fields.Taaknummer'].astype(int)
        
        print(f"Data opgehaald. Aantal rijen: {len(df)}")
        return df
    
    except Exception as e:
        print(f"Fout bij ophalen Airtable data: {str(e)}")
        raise

def format_phone_number(phone):
    """Formatteert telefoonnummer naar het juiste formaat voor Trengo."""
    phone = ''.join(filter(str.isdigit, str(phone)))
    if phone.startswith('0'):
        phone = '31' + phone[1:]
    elif not phone.startswith('31'):
        phone = '31' + phone
    return phone

def send_whatsapp_message(naam, monteur, dagnaam, datum, begintijd, eindtijd, reparatieduur, taaknummer, mobielnummer):
    """Verstuurt WhatsApp bericht via Trengo."""
    url = "https://app.trengo.com/api/v2/wa_sessions"
    formatted_phone = format_phone_number(mobielnummer)
    formatted_date = format_date(datum)
    
    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": WHATSAPP_TEMPLATE_ID,
        "params": [
            {"type": "body", "key": "{{1}}", "value": str(naam)},
            {"type": "body", "key": "{{2}}", "value": str(monteur)},
            {"type": "body", "key": "{{3}}", "value": str(dagnaam)},
            {"type": "body", "key": "{{4}}", "value": formatted_date},
            {"type": "body", "key": "{{5}}", "value": str(monteur)},
            {"type": "body", "key": "{{6}}", "value": str(begintijd)},
            {"type": "body", "key": "{{7}}", "value": str(eindtijd)},
            {"type": "body", "key": "{{8}}", "value": str(reparatieduur)},
            {"type": "body", "key": "{{9}}", "value": str(taaknummer)}
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
        df = get_airtable_data()
        
        if df.empty:
            print("Geen data gevonden om te verwerken")
            return
        
        for index, row in df.iterrows():
            try:
                print(f"\nVerwerken rij {index + 1}: {row['fields.Naam bewoner']}")
                
                if 'fields.Mobielnummer' not in row or pd.isna(row['fields.Mobielnummer']):
                    print(f"Geen telefoonnummer gevonden voor {row['fields.Naam bewoner']}, deze rij wordt overgeslagen")
                    continue
                
                # Send message
                send_whatsapp_message(
                    naam=row['fields.Naam bewoner'],
                    monteur=row['fields.Monteur'],
                    dagnaam=row['fields.Dagnaam'],
                    datum=row['fields.Datum bezoek'],
                    begintijd=row['fields.Begintijd'],
                    eindtijd=row['fields.Eindtijd'],
                    reparatieduur=row['fields.Reparatieduur'],
                    taaknummer=row['fields.Taaknummer'],
                    mobielnummer=row['fields.Mobielnummer']
                )
                
                # Delete record after successful send
                delete_airtable_record(row['id'])
                print(f"Bericht verstuurd en record verwijderd voor {row['fields.Naam bewoner']}")
                
            except Exception as e:
                print(f"Fout bij verwerken rij {index}: {str(e)}")
                continue
    
    except Exception as e:
        print(f"Algemene fout: {str(e)}")

print("Start eerste verwerking...")
process_data()

scheduler = BlockingScheduler()
scheduler.add_job(process_data, 'interval', minutes=30)

if __name__ == "__main__":
    print("\nStarting scheduler...")
    scheduler.start()
