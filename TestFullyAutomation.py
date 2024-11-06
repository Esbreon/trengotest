import os
import sys
import requests
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
import imaplib
import email
from email.header import decode_header

def connect_to_outlook():
    """Maakt verbinding met Outlook."""
    print("Verbinding maken met Outlook...")
    imap_server = "outlook.office365.com"
    email_address = os.environ.get('OUTLOOK_EMAIL')
    password = os.environ.get('OUTLOOK_PASSWORD')
    
    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(email_address, password)
    print("Verbinding met Outlook succesvol")
    return imap

def download_excel_attachment(imap, sender_email, subject_line):
    """Downloadt Excel bijlage uit specifieke email."""
    print(f"\nZoeken naar emails van {sender_email} met onderwerp '{subject_line}'...")
    imap.select('INBOX')
    
    search_criteria = f'(FROM "{sender_email}" SUBJECT "{subject_line}" UNSEEN)'
    _, message_numbers = imap.search(None, search_criteria)
    
    if not message_numbers[0]:
        print("Geen nieuwe emails gevonden")
        return None
    
    print("Nieuwe email(s) gevonden, bijlage controleren...")
    
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
                
                print(f"Excel bijlage gevonden: {filename}")
                print(f"Opslaan als: {filepath}")
                
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                
                imap.store(num, '+FLAGS', '\\Seen')
                print("Email gemarkeerd als gelezen")
                return filepath
    
    print("Geen Excel bijlage gevonden in nieuwe emails")
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
        print(f"Versturen WhatsApp bericht naar TEST NUMMER {formatted_phone} voor {naam_bewoner}...")
        print(f"Bericht details: Datum={formatted_date}, Tijdvak={tijdvak}, Reparatieduur={reparatieduur}")
        response = requests.post(url, json=payload, headers=headers)
        print(f"Trengo response: {response.text}")
        return response.json()
    except Exception as e:
        print(f"Fout bij versturen bericht: {str(e)}")
        raise

def process_excel_file(filepath):
    """Verwerkt Excel bestand en stuurt berichten."""
    try:
        print(f"\nVerwerken Excel bestand: {filepath}")
        df = pd.read_excel(filepath)
        
        if df.empty:
            print("Geen data gevonden in Excel bestand")
            return
        
        print(f"Aantal rijen gevonden: {len(df)}")
        print(f"Kolommen in bestand: {', '.join(df.columns)}")
        
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
                print(f"\nVerwerken rij {index + 1}/{len(df)}")
                if pd.isna(row['fields.Mobielnummer']):
                    print(f"Geen telefoonnummer gevonden voor {row['fields.Naam bewoner']}, deze overslaan")
                    continue
                
                send_whatsapp_message(
                    naam_bewoner=row['fields.Naam bewoner'],
                    datum=row['fields.Datum bezoek'],
                    tijdvak=row['fields.Tijdvak'],
                    reparatieduur=row['fields.Reparatieduur'],
                    mobielnummer=row['fields.Mobielnummer']
                )
                
                print(f"Bericht verstuurd voor {row['fields.Naam bewoner']}")
                
            except Exception as e:
                print(f"Fout bij verwerken rij {index}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Fout bij verwerken Excel bestand: {str(e)}")

def process_data():
    """Main function to check email and process Excel."""
    print(f"\n=== Start nieuwe verwerking: {datetime.now()} ===")
    
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
                print(f"\nVerwijderen tijdelijk bestand: {excel_file}")
                os.remove(excel_file)
            else:
                print("Geen nieuwe Excel bestanden gevonden om te verwerken")
                
        finally:
            print("Outlook verbinding sluiten")
            imap.logout()
            
    except Exception as e:
        print(f"Algemene fout: {str(e)}")

# Start script
if __name__ == "__main__":
    print("\n=== ENVIRONMENT CHECK ===")
    required_vars = [
        'OUTLOOK_EMAIL', 
        'OUTLOOK_PASSWORD', 
        'SENDER_EMAIL', 
        'SUBJECT_LINE',
        'WHATSAPP_TEMPLATE_ID',
        'TRENGO_API_KEY'
    ]
    
    missing_vars = [var for var in required_vars if not os.environ.get(var)]
    if missing_vars:
        print(f"ERROR: Missende environment variables: {', '.join(missing_vars)}")
        sys.exit(1)
    
    print("Alle environment variables zijn ingesteld")
    
    print("\n=== EERSTE TEST ===")
    print("Handmatige test uitvoeren...")
    try:
        process_data()
        print("Handmatige test compleet")
    except Exception as e:
        print(f"Fout tijdens handmatige test: {str(e)}")
    
    print("\n=== SCHEDULER STARTEN ===")
    print("Automatische verwerking elke 30 minuten wordt gestart...")
    scheduler = BlockingScheduler()
    scheduler.add_job(process_data, 'interval', minutes=30)
    
    try:
        print("Scheduler draait nu... (Ctrl+C om te stoppen)")
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        print('\nScript gestopt')
