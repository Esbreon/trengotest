import os
import requests
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler

AIRTABLE_BASE_ID = os.environ.get('AIRTABLE_BASE_ID')
AIRTABLE_TABLE_NAME = os.environ.get('AIRTABLE_V4H')
AIRTABLE_API_KEY = os.environ.get('AIRTABLE_API_KEY')
WHATSAPP_TEMPLATE_ID = os.environ.get('WHATSAPP_TEMPLATE_ID_VESTEDA_4H')

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
        raise Exception(f"Error deleting record: {response.status_code} - {response.text}")
    print(f"Record {record_id} successfully deleted")

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

def send_whatsapp_message(naam_bewoner, datum, tijdvak, reparatieduur, mobielnummer):
    """Sends WhatsApp message via Trengo with the new template format."""
    url = "https://app.trengo.com/api/v2/wa_sessions"
    formatted_phone = format_phone_number(mobielnummer)
    formatted_date = format_date(datum)
    
    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": WHATSAPP_TEMPLATE_ID,
        "params": [
            {"type": "body", "key": "{{1}}", "value": str(naam_bewoner)},
            {"type": "body", "key": "{{2}}", "value": formatted_date},
            {"type": "body", "key": "{{3}}", "value": str(tijdvak)},
            {"type": "body", "key": "{{4}}", "value": str(reparatieduur)}
        ]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": "Bearer " + os.environ.get('TRENGO_API_KEY')
    }
    
    try:
        print(f"Sending message to {formatted_phone} for {naam_bewoner}...")
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
                print(f"\nProcessing row {index + 1}: {row['fields.Naam bewoner']}")
                
                if 'fields.Mobielnummer' not in row or pd.isna(row['fields.Mobielnummer']):
                    print(f"No phone number found for {row['fields.Naam bewoner']}, skipping row")
                    continue
                
                # Send message
                send_whatsapp_message(
                    naam_bewoner=row['fields.Naam bewoner'],
                    datum=row['fields.Datum bezoek'],
                    tijdvak=row['fields.Tijdvak'],
                    reparatieduur=row['fields.Reparatieduur'],
                    mobielnummer=row['fields.Mobielnummer']
                )
                
                # Delete record after successful send
                delete_airtable_record(row['id'])
                print(f"Message sent and record deleted for {row['fields.Naam bewoner']}")
                
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
