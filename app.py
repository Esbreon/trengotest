import os
import requests
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler

# Airtable configuration
AIRTABLE_BASE_ID = os.environ.get('AIRTABLE_BASE_ID')
AIRTABLE_TABLE_NAME = os.environ.get('AIRTABLE_TABLE_NAME')
AIRTABLE_API_KEY = os.environ.get('AIRTABLE_API_KEY')

def get_airtable_data():
    """Fetches data from Airtable."""
    try:
        print("Starting Airtable data fetch...")
        
        url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_NAME}"
        
        headers = {
            "Authorization": f"Bearer {AIRTABLE_API_KEY}",
            "Content-Type": "application/json"
        }
        
        print("Connecting to Airtable...")
        
        response = requests.get(url, headers=headers)
        
        if response.status_code != 200:
            raise Exception(f"Error fetching data: {response.status_code} - {response.text}")
        data = response.json()
        
        # Convert records to DataFrame
        df = pd.json_normalize(data['records'])
        
        print(f"Data retrieved. Row count: {len(df)}")
        print("Sample of retrieved data:")
        print(df.head())
        return df
    
    except Exception as e:
        print(f"Error fetching Airtable data: {str(e)}")
        raise

def format_phone_number(phone):
    """Formats phone number to the correct format for Trengo."""
    if not phone:  # Check for None or empty string
        return None
        
    # Remove all non-numeric characters
    phone = ''.join(filter(str.isdigit, str(phone)))
    
    # Validate phone number length
    if len(phone) < 9:  # Basic validation for minimum length
        return None
        
    # If number starts with 0, replace with 31
    if phone.startswith('0'):
        phone = '31' + phone[1:]
    # If number doesn't start with 31, add it
    elif not phone.startswith('31'):
        phone = '31' + phone
    
    return phone

def send_whatsapp_message(naam, dag, tijdvak, telefoon):
    """Sends WhatsApp message via Trengo."""
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    # Format the phone number
    formatted_phone = format_phone_number(telefoon)
    if not formatted_phone:
        raise ValueError(f"Invalid phone number for {naam}: {telefoon}")
    
    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": 181327,
        "params": [
            {
                "type": "body",
                "key": "{{1}}",
                "value": str(naam)  # Fixed: Using the parameter naam instead of Naam klant
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
            },
            {
                "type": "body",
                "key": "{{4}}",
                "value": "60 minuten"  # Added default repair duration
            }
        ]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }
    
    try:
        print(f"Sending message to {formatted_phone} for {naam}...")
        response = requests.post(url, json=payload, headers=headers)
        
        if response.status_code not in (200, 201):
            raise Exception(f"Error from Trengo API: {response.status_code} - {response.text}")
            
        print(f"Message sent successfully to {naam}")
        return response.json()
    
    except Exception as e:
        print(f"Error sending message: {str(e)}")
        raise

def process_data():
    """Main function that fetches data and sends messages."""
    print(f"\n=== Starting new processing cycle: {datetime.now()} ===")
    
    try:
        # Fetch data
        df = get_airtable_data()
        
        if df.empty:
            print("No data found to process")
            return
        
        # Process each row
        for index, row in df.iterrows():
            try:
                naam = row.get('fields.Naam klant')
                print(f"\nProcessing row {index + 1}: {naam}")
                
                # Check if phone number exists
                if 'fields.Mobielnummer' not in row or pd.isna(row['fields.Mobielnummer']):
                    print(f"No phone number found for {naam}, skipping this row")
                    continue
                
                # Validate required fields
                required_fields = {
                    'Naam': naam,
                    'Datum': row.get('fields.Datum'),
                    'Tijdvak': row.get('fields.Tijdvak'),
                    'Mobielnummer': row.get('fields.Mobielnummer')
                }
                
                if any(not value for value in required_fields.values()):
                    missing = [k for k, v in required_fields.items() if not v]
                    print(f"Missing required fields for {naam}: {', '.join(missing)}")
                    continue
                
                send_whatsapp_message(
                    naam=naam,
                    dag=row['fields.Datum'],
                    tijdvak=row['fields.Tijdvak'],
                    telefoon=row['fields.Mobielnummer']
                )
                
            except Exception as e:
                print(f"Error processing row {index}: {str(e)}")
                continue
    
    except Exception as e:
        print(f"General error: {str(e)}")

def main():
    # Start initial processing
    print("Starting initial processing...")
    process_data()

    # Schedule future processing
    scheduler = BlockingScheduler()
    scheduler.add_job(process_data, 'interval', minutes=30)

    print("\nStarting scheduler...")
    scheduler.start()

if __name__ == "__main__":
    main()
