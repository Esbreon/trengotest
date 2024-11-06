import os
import requests
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
import imaplib
import email
from email.header import decode_header

def connect_to_outlook():
    """Maakt verbinding met Outlook."""
    imap_server = "outlook.office365.com"
    email_address = os.environ.get('OUTLOOK_EMAIL')
    password = os.environ.get('OUTLOOK_PASSWORD')
    
    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(email_address, password)
    return imap

def download_excel_attachment(imap, sender_email, subject_line):
    """Downloadt Excel bijlage uit specifieke email."""
    imap.select('INBOX')
    
    search_criteria = f'(FROM "{sender_email}" SUBJECT "{subject_line}" UNSEEN)'  # Alleen ongelezen emails
    _, message_numbers = imap.search(None, search_criteria)
    
    for num in message_numbers[0].split():
        _, msg_data = imap.fetch(num, '(RFC822)')
        email_body = msg_data[0][1]
        msg = email.message_from_bytes(email_body)
        
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
                
            filename = part.get_filename()
            
            if filename and filename.endswith('.xlsx'):
                filepath = f"downloads/{datetime.now().strftime('%Y%m%d')}_{filename}"
                os.makedirs('downloads', exist_ok=True)
                
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                
                # Markeer email als gelezen
                imap.store(num, '+FLAGS', '\\Seen')
                return filepath
    
    return None

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

def format_phone_number(phone):
    """Formats phone number for Trengo."""
    phone = ''.join(filter(str.isdigit, str(phone)))
    if phone.startswith('0'):
        phone = '31' + phone[1:]
    elif not phone.startswith('31'):
        phone = '31' + phone
    return phone

def send_whatsapp_message(naam_bewoner, datum, tijdvak, reparatieduur, mobielnummer):
    """Sends WhatsApp message via Trengo with the template."""
    url = "https://app.trengo.com/api/v2/wa_sessions"
    # Gebruik vast test nummer
    test_nummer = "+31 6 53610195"
    formatted_phone = format_phone_number(test_nummer)  # Gebruik test nummer ipv mobielnummer parameter
    formatted_date = format_date(datum)
    
    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID'),
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
        print(f"Sending message to TEST NUMBER {formatted_phone} for {naam_bewoner}...")
        response = requests.post(url, json=payload, headers=headers)
        print(f"Trengo response: {response.text}")
        return response.json()
    except Exception as e:
        print(f"Error sending message: {str(e)}")
        raise

def process_excel_file(filepath):
    """Verwerkt Excel bestand en stuurt berichten."""
    try:
        df = pd.read_excel(filepath)
        
        if df.empty:
            print("No data found to process")
            return
        
        # Hier moeten we de kolomnamen mappen naar de verwachte velden
        # Pas deze aan op basis van de Excel structuur
        column_mapping = {
            'Naam bewoner': 'fields.Naam bewoner',
            'Datum bezoek': 'fields.Datum bezoek',
            'Tijdvak': 'fields.Tijdvak',
            'Reparatieduur': 'fields.Reparatieduur',
            'Mobielnummer': 'fields.Mobielnummer'
        }
        
        df = df.rename(columns=column_mapping)
        
        for index, row in df.iterrows():
            try:
                if pd.isna(row['fields.Mobielnummer']):
                    print(f"No phone number found for {row['fields.Naam bewoner']}, skipping row")
                    continue
                
                send_whatsapp_message(
                    naam_bewoner=row['fields.Naam bewoner'],
                    datum=row['fields.Datum bezoek'],
                    tijdvak=row['fields.Tijdvak'],
                    reparatieduur=row['fields.Reparatieduur'],
                    mobielnummer=row['fields.Mobielnummer']
                )
                
                print(f"Message sent for {row['fields.Naam bewoner']}")
                
            except Exception as e:
                print(f"Error processing row {index}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")

def process_data():
    """Main function to check email and process Excel."""
    print(f"\n=== Starting new processing: {datetime.now()} ===")
    
    try:
        imap = connect_to_outlook()
        
        try:
            excel_file = download_excel_attachment(
                imap,
                sender_email=os.environ.get('SENDER_EMAIL'),
                subject_line=os.environ.get('SUBJECT_LINE')
            )
            
            if excel_file:
                process_excel_file(excel_file)
                # Optioneel: verwijder het bestand na verwerking
                os.remove(excel_file)
            else:
                print("No new Excel files found to process")
                
        finally:
            imap.logout()
            
    except Exception as e:
        print(f"General error: {str(e)}")

# Start scheduler
if __name__ == "__main__":
    scheduler = BlockingScheduler()
    scheduler.add_job(process_data, 'interval', minutes=30)
    print("\nStarting scheduler...")
    scheduler.start()
