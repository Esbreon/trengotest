import os
import pyodbc
import pandas as pd
import requests  
from flask import Flask, jsonify

app = Flask(__name__)

def get_sql_connection():
    """Establish connection to SQL Server using FreeTDS."""
    conn = pyodbc.connect(
        DRIVER='/usr/lib/x86_64-linux-gnu/odbc/libtdsodbc.so', 
        SERVER=os.getenv('SQL_SERVER'),
        DATABASE=os.getenv('SQL_DATABASE'),
        UID=os.getenv('SQL_USER'),
        PWD=os.getenv('SQL_PASSWORD'),
        PORT=1433,
        TDS_Version='7.4'
    )
    return conn

def fetch_data_from_sql():
    """Fetch data from the SQL view."""
    query = "SELECT RawContactFormattedPhoneTrengo AS phone_number, RelationFullName AS name FROM Fixzed_Staging.LIST_Feedback WHERE CanSendFeedback = 1"
    conn = get_sql_connection()
    df = pd.read_sql(query, conn)
    conn.close()
    return df

def format_phone_number(phone):
    """Ensure phone number is properly formatted."""
    if pd.isna(phone):
        return None
    phone = str(phone).strip()
    return phone if phone.startswith('+') else f'+{phone}'

def send_whatsapp_message(name, phone_number):
    """Simulates sending WhatsApp message via Trengo (always in test mode)."""
    test_mode = True  
    formatted_phone = format_phone_number(phone_number)
    if not formatted_phone:
        print(f"Invalid phone number for {name}")
        return

    payload = {
        "recipient_phone_number": formatted_phone,
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_FB_PW'),
        "params": [
            {"type": "body", "key": "{{1}}", "value": str(name)}
        ]
    }

    if test_mode:
        print(f"[TEST MODE] WhatsApp message **would** be sent to {formatted_phone} for {name}.")
        print(f"[TEST MODE] Payload: {payload}")
    else:
        print("Foutje")
        #url = "https://app.trengo.com/api/v2/wa_sessions"
        #headers = {
        #    "accept": "application/json",
        #    "content-type": "application/json",
        #    "Authorization": "Bearer " + os.environ.get('TRENGO_API_KEY')
        #}
        #try:
        #    response = requests.post(url, json=payload, headers=headers)
        #    response.raise_for_status()
        #    print(f"WhatsApp message sent to {formatted_phone} for {name}")
        #except requests.exceptions.HTTPError as e:
        #    print(f"Error sending message: {str(e)}")
        #    if e.response is not None:
        #        print(f"Response body: {e.response.text}")
        #    raise

def process_feedback_requests(test_mode=True):
    """Fetch eligible tasks and print WhatsApp messages in test mode."""
    print("Fetching feedback requests from SQL view...")
    df = fetch_data_from_sql()
    
    if df.empty:
        print("No eligible feedback requests found.")
        return
    
    print(f"Found {len(df)} tasks eligible for feedback requests.")
    for index, row in df.iterrows():
        try:
            send_whatsapp_message(
                name=row.get('name', 'Resident'),
                phone_number=row.get('phone_number'),
                test_mode=test_mode
            )
        except Exception as e:
            print(f"Error processing task {index}: {str(e)}")

def main():
    """Main execution function."""
    process_feedback_requests(test_mode=True) 

if __name__ == "__main__":
    main()
