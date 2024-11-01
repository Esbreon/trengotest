import os
import requests
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler

AIRTABLE_BASE_ID = os.environ.get('AIRTABLE_BASE_ID')
AIRTABLE_TABLE_NAME = os.environ.get('AIRTABLE_PW1H')
AIRTABLE_API_KEY = os.environ.get('AIRTABLE_API_KEY')
WHATSAPP_TEMPLATE_ID = os.environ.get('WHATSAPP_TEMPLATE_ID_PW_1H')

def format_date(date_str):
    """Formatteert datum naar dd MMM yy formaat met Nederlandse maandafkortingen."""
    try:
        # Nederlandse maandafkortingen
        nl_month_abbr = {
            1: 'jan',
            2: 'feb',
            3: 'mrt',
            4: 'apr',
            5: 'mei',
            6: 'jun',
            7: 'jul',
            8: 'aug',
            9: 'sep',
            10: 'okt',
            11: 'nov',
            12: 'dec'
        }
        
        # Eerst proberen te parsen als datetime object
        if isinstance(date_str, datetime):
            date_obj = date_str
        else:
            # Probeer verschillende datumformaten die uit Airtable kunnen komen
            try:
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                try:
                    date_obj = datetime.strptime(date_str, '%d/%m/%Y')
                except ValueError:
                    date_obj = datetime.strptime(date_str, '%d-%m-%Y')
        
        # Haal dag en maandnummer op
        day = date_obj.day
        month = date_obj.month
        year = str(date_obj.year)[2:]  # Laatste 2 cijfers van het jaar
        
        # Formatteer met Nederlandse maandafkorting
        formatted_date = f"{day} {nl_month_abbr[month]} {year}"
        return formatted_date
    
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
