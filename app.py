import os
import pandas as pd
import requests
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from io import BytesIO
from apscheduler.schedulers.blocking import BlockingScheduler
from datetime import datetime

# Initialize scheduler
scheduler = BlockingScheduler()

def read_sharepoint_excel():
    """
    Leest Excel bestand van SharePoint
    """
    try:
        # Haal configuratie uit environment variables
        SITE_URL = os.environ['SHAREPOINT_SITE_URL']
        EXCEL_PATH = os.environ['SHAREPOINT_EXCEL_PATH']
        CLIENT_ID = os.environ['SHAREPOINT_CLIENT_ID']
        CLIENT_SECRET = os.environ['SHAREPOINT_CLIENT_SECRET']

        auth = AuthenticationContext(SITE_URL)
        auth.acquire_token_for_app(client_id=CLIENT_ID, client_secret=CLIENT_SECRET)
        ctx = ClientContext(SITE_URL, auth)
        
        response = File.open_binary(ctx, EXCEL_PATH)
        bytes_file_obj = BytesIO()
        bytes_file_obj.write(response.content)
        bytes_file_obj.seek(0)
        
        return pd.read_excel(bytes_file_obj)
    except Exception as e:
        print(f"Fout bij lezen SharePoint bestand: {str(e)}")
        raise

def send_whatsapp_template(naam, dag, tijdvak):
    """
    Verstuurt WhatsApp template via Trengo API
    """
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    # Vast telefoonnummer voor testing
    TEST_PHONE = "31611341059"
    
    payload = {
        "recipient_phone_number": TEST_PHONE,
        "hsm_id": 181327,
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
        "Authorization": f"Bearer {os.environ['TRENGO_API_TOKEN']}"
    }
    
    response = requests.post(url, json=payload, headers=headers)
    return response.json()

@scheduler.scheduled_job('interval', minutes=30)
def process_excel():
    """
    Hoofdfunctie die elke 30 minuten wordt uitgevoerd
    """
    print(f"Start verwerking: {datetime.now()}")
    
    try:
        # Lees data van SharePoint
        df = read_sharepoint_excel()
        print("Excel bestand succesvol gelezen van SharePoint")
        
        # Verwerk elke rij
        for index, row in df.iterrows():
            try:
                response = send_whatsapp_template(
                    naam=row['naam'],
                    dag=row['dag'],
                    tijdvak=row['tijdvak']
                )
                
                print(f"Bericht verstuurd voor rij {index + 1}: {response}")
                
            except Exception as e:
                print(f"Fout bij versturen van rij {index + 1}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Algemene fout: {str(e)}")

if __name__ == "__main__":
    print("Scheduler starting...")
    scheduler.start()
