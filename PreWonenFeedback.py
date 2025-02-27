import os
import pyodbc
import pandas as pd
from flask import Flask, jsonify

app = Flask(__name__)

def get_sql_connection():
    """Establish connection to SQL Server using FreeTDS."""
    conn = pyodbc.connect(
        DRIVER='{ODBC Driver 17 for SQL Server}',  # Use FreeTDS-compatible driver
        SERVER=os.getenv('SQL_SERVER'),
        DATABASE=os.getenv('SQL_DATABASE'),
        UID=os.getenv('SQL_USER'),
        PWD=os.getenv('SQL_PASSWORD')
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

@app.route('/process_feedback_requests', methods=['GET'])
def process_feedback_requests():
    """Fetch eligible tasks and return WhatsApp messages as JSON."""
    df = fetch_data_from_sql()

    if df.empty:
        return jsonify({"message": "No eligible feedback requests found."})

    result = []
    for _, row in df.iterrows():
        formatted_phone = format_phone_number(row.get('phone_number'))
        if not formatted_phone:
            result.append({"name": row.get('name', 'Resident'), "status": "Invalid phone number"})
        else:
            result.append({"name": row.get('name', 'Resident'), "phone_number": formatted_phone, "status": "Ready to send"})
    
    return jsonify(result)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
