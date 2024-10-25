import os
import requests
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler

# Airtable configuration
AIRTABLE_BASE_ID = os.environ.get('AIRTABLE_BASE_ID')
AIRTABLE_TABLE_NAME = os.environ.get('AIRTABLE_T4H')
AIRTABLE_API_KEY = os.environ.get('AIRTABLE_API_KEY')
WHATSAPP_TEMPLATE_ID = os.environ.get('WHATSAPP_TEMPLATE_ID_B_TEST')

def get_airtable_data():
    """Fetches data from Airtable."""
    try:
        print("Starting Airtable data fetch...")
        
        url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_NAME}"
        
        headers = {
            "Authorization": f"Bearer {AIRTABLE_API_KEY}",
            "Content-Type": "application/json"
        }
        
        response = requests.get(url, headers=headers)
        
        if response.status_code != 200:
            raise Exception(f"Error fetching data: {response.status_code} - {response.text}")
        
        data = response.json()
        df = pd.json_normalize(data['records'])
        print(f"Data retrieved. Row count: {len(df)}")
        return df
    
    except Exception as e:
        print(f"Error fetching Airtable data: {str(e)}")
        raise

def format_phone_number(phone):
    """Formats phone number for Trengo."""
    phone = ''.join(filter(str.isdigit, str(phone)))
    if phone.startswith('0'):
        phone = '31' + phone[1:]
    elif not phone.startswith('31'):
        phone = '31' + phone
    return phone

def send_whatsapp_message(naam_klant, datum, tijdvak, reparatieduur, mobielnummer):
    """Sends WhatsApp message via Trengo with the new template format."""
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    formatted_phone = format_phone_number(mobielnummer)
    
    # Format parameters according to the new template
    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": WHATSAPP_TEMPLATE_ID,
        "params": [
            {
                "type": "body",
                "key": "{{1}}",
                "value": str(naam_klant)  # Name from first column
            },
            {
                "type": "body",
                "key": "{{2}}",
                "value": str(datum)  # Date from second column
            },
            {
                "type": "body",
                "key": "{{3}}",
                "value": str(tijdvak)  # Time slot from third column
            },
            {
                "type": "body",
                "key": "{{4}}",
                "value": str(reparatieduur)  # Duration from fourth column
            }
        ]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": "Bearer " + os.environ.get('TRENGO_API_KEY')
    }
    
    try:
        print(f"Sending message to {formatted_phone} for {naam_klant}...")
        response = requests.post(url, json=payload, headers=headers)
        print(f"Trengo response: {response.text}")
        return response.json()
    
    except Exception as e:
        print(f"Error sending message: {str(e)}")
        raise

def process_data():
    """Main function to fetch data and send messages."""
    print(f"\n=== Starting new processing: {datetime.now()} ===")
    
    try:
        df = get_airtable_data()
        
        if df.empty:
            print("No data found to process")
            return
        
        for index, row in df.iterrows():
            try:
                print(f"\nProcessing row {index + 1}: {row['fields.Naam klant']}")
                
                if 'fields.Mobielnummer' not in row or pd.isna(row['fields.Mobielnummer']):
                    print(f"No phone number found for {row['fields.Naam klant']}, skipping row")
                    continue
                
                send_whatsapp_message(
                    naam_klant=row['fields.Naam klant'],
                    datum=row['fields.Datum'],
                    tijdvak=row['fields.Tijdvak'],
                    reparatieduur=row['fields.Reparatieduur'],
                    mobielnummer=row['fields.Mobielnummer']
                )
                print(f"Message sent for {row['fields.Naam klant']}")
                
            except Exception as e:
                print(f"Error processing row {index}: {str(e)}")
                continue
    
    except Exception as e:
        print(f"General error: {str(e)}")

# Start initial processing
print("Starting first processing...")
process_data()

# Schedule future processing
scheduler = BlockingScheduler()
scheduler.add_job(process_data, 'interval', minutes=30)

if __name__ == "__main__":
    print("\nStarting scheduler...")
    scheduler.start()
