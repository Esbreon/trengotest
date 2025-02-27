import os
import pyodbc
import pandas as pd

def get_sql_connection():
    """Establish connection to SQL Server using FreeTDS."""
    conn = pyodbc.connect(
        DRIVER='{SQL Server}',
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

def print_whatsapp_message(name, phone_number):
    """Prints the WhatsApp message instead of sending it."""
    formatted_phone = format_phone_number(phone_number)
    if not formatted_phone:
        print(f"Invalid phone number for {name}")
        return
    
    print(f"[TEST] Would send WhatsApp message to {formatted_phone} for {name}")

def process_feedback_requests():
    """Fetch eligible tasks and print WhatsApp messages."""
    print("Fetching feedback requests from SQL view...")
    df = fetch_data_from_sql()
    
    if df.empty:
        print("No eligible feedback requests found.")
        return
    
    print(f"Found {len(df)} tasks eligible for feedback requests.")
    for index, row in df.iterrows():
        try:
            print_whatsapp_message(
                name=row.get('name', 'Resident'),
                phone_number=row.get('phone_number')
            )
        except Exception as e:
            print(f"Error processing task {index}: {str(e)}")

def main():
    """Main execution function."""
    process_feedback_requests()

if __name__ == "__main__":
    main()
