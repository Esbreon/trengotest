import os
import sys
import math
import json
import base64
import requests
import pandas as pd
from datetime import datetime
import msal

# === CONFIGURATION ===
CUSTOM_FIELDS = {
    "locatie": 613776,
    "element": 618192,
    "defect": 618193,
    "werkbonnummer": 618194,
    "binnen_of_buiten": 618205
}
PER_PAGE = 25
MAX_TICKETS = 500

REQUIRED_ENV_VARS = [
    'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'AZURE_TENANT_ID',
    'OUTLOOK_EMAIL', 'OUTLOOK_PASSWORD', 'TEST_EMAIL', 'SUBJECT_LINE_PW_BEVESTIGING',
    'WHATSAPP_TEMPLATE_ID_TEST_BEVESTIGING', 'TRENGO_API_KEY'
]

class OutlookClient:
    def __init__(self):
        self.client_id = os.getenv('AZURE_CLIENT_ID')
        self.client_secret = os.getenv('AZURE_CLIENT_SECRET')
        self.tenant_id = os.getenv('AZURE_TENANT_ID')
        self.username = os.getenv('OUTLOOK_EMAIL')
        self.password = os.getenv('OUTLOOK_PASSWORD')
        self.app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )

    def get_token(self):
        scopes = ['https://graph.microsoft.com/Mail.Read']
        result = self.app.acquire_token_by_username_password(
            username=self.username,
            password=self.password,
            scopes=scopes
        )
        if "access_token" not in result:
            raise Exception(f"Token acquisition failed: {result.get('error_description')}")
        return result["access_token"]

    def download_excel_attachment(self, sender_email, subject_line):
        token = self.get_token()
        headers = {'Authorization': f'Bearer {token}'}
        filter_query = f"from/emailAddress/address eq '{sender_email}' and subject eq '{subject_line}' and isRead eq false"
        url = 'https://graph.microsoft.com/v1.0/me/messages?$filter=' + filter_query

        response = requests.get(url, headers=headers)
        response.raise_for_status()
        messages = response.json().get('value', [])

        for msg in messages:
            if not msg.get("hasAttachments"):
                continue
            att_url = f"https://graph.microsoft.com/v1.0/me/messages/{msg['id']}/attachments"
            att_resp = requests.get(att_url, headers=headers)
            att_resp.raise_for_status()
            for att in att_resp.json().get("value", []):
                if att.get("name", "").endswith(".xlsx"):
                    filepath = f"downloads/{datetime.now().strftime('%Y%m%d')}_{att['name']}"
                    os.makedirs('downloads', exist_ok=True)
                    with open(filepath, 'wb') as f:
                        f.write(base64.b64decode(att['contentBytes']))
                    requests.patch(f"https://graph.microsoft.com/v1.0/me/messages/{msg['id']}", headers=headers, json={"isRead": True})
                    return filepath
        return None


def fetch_recent_trengo_tickets():
    token = os.getenv("TRENGO_API_KEY")
    tickets_by_werkbon = {}
    next_url = f"https://app.trengo.com/api/v2/tickets?page=1&per_page={PER_PAGE}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    page = 1
    ticket_count = 0

    while next_url and ticket_count < MAX_TICKETS:
        print(f"Retrieved page {page}")
        resp = requests.get(next_url, headers=headers)
        resp.raise_for_status()
        payload = resp.json()
        for ticket in payload.get("data", []):
            ticket_count += 1
            if ticket_count > MAX_TICKETS:
                break
            custom_fields = ticket.get("custom_field_values", [])
            ticket_id = ticket.get("id")
            status = ticket.get("status", "")
            archived = ticket.get("archived_at") is not None
            is_open = status in ["OPEN", "ASSIGNED", "CLOSED"] or archived
            for field in custom_fields:
                if str(field.get("custom_field_id")) == str(CUSTOM_FIELDS["werkbonnummer"]):
                    werkbon = str(field.get("value", "")).strip()
                    if werkbon:
                        tickets_by_werkbon.setdefault(werkbon, []).append({"ticket_id": ticket_id, "is_open": is_open})
        next_url = payload.get("links", {}).get("next")
        page += 1
    return tickets_by_werkbon


def safe_str(val):
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()

def format_date(date_str):
    try:
        if pd.isna(date_str):
            return ""
        for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
            try:
                dt = datetime.strptime(str(date_str), fmt)
                break
            except: continue
        else:
            return str(date_str)
        nl_months = ['januari','februari','maart','april','mei','juni','juli','augustus','september','oktober','november','december']
        return f"{dt.day} {nl_months[dt.month - 1]} {dt.year}"
    except: return str(date_str)

def format_phone(phone):
    if pd.isna(phone): return None
    phone = str(phone).strip()
    if phone.endswith('.0'): phone = phone.split('.')[0]
    return phone

def send_message_with_ticket(ticket_id, params):
    url = f"https://app.trengo.com/api/v2/messages"
    headers = {"Authorization": f"Bearer {os.getenv('TRENGO_API_KEY')}", "Content-Type": "application/json"}
    payload = {
        "ticket_id": ticket_id,
        "message_type": "whatsapp_template",
        "hsm_id": os.getenv('WHATSAPP_TEMPLATE_ID_TEST_BEVESTIGING'),
        "params": params
    }
    requests.post(url, json=payload, headers=headers).raise_for_status()
    print(f"✅ Message appended to existing ticket {ticket_id}")

def send_new_whatsapp_message(phone, params):
    url = "https://app.trengo.com/api/v2/wa_sessions"
    headers = {"Authorization": f"Bearer {os.getenv('TRENGO_API_KEY')}", "Content-Type": "application/json"}
    payload = {"recipient_phone_number": phone, "hsm_id": os.getenv('WHATSAPP_TEMPLATE_ID_TEST_BEVESTIGING'), "params": params}
    resp = requests.post(url, json=payload, headers=headers)
    resp.raise_for_status()
    print(f"✅ New WhatsApp message sent to {phone}")
    return resp.json().get("message", {}).get("ticket_id")

def set_custom_fields(ticket_id, fields):
    headers = {"Authorization": f"Bearer {os.getenv('TRENGO_API_KEY')}", "Content-Type": "application/json"}
    for fid, val in fields.items():
        requests.post(f"https://app.trengo.com/api/v2/tickets/{ticket_id}/custom_fields", json={"custom_field_id": fid, "value": val}, headers=headers).raise_for_status()

def process_excel_file(filepath, ticket_lookup):
    df = pd.read_excel(filepath)
    print(f"📄 Rijen in Excel: {len(df)}")
    for _, row in df.iterrows():
        werkbon = safe_str(row['Werkbonnummer'])
        mobiel = format_phone(row['Mobielnummer'])
        if not mobiel or not werkbon:
            continue

        params = [
            {"type": "body", "key": "{{1}}", "value": safe_str(row['Naam bewoner'])},
            {"type": "body", "key": "{{2}}", "value": safe_str(row['Taaktype'])},
            {"type": "body", "key": "{{3}}", "value": safe_str(row['Dag'])},
            {"type": "body", "key": "{{4}}", "value": format_date(row['Datum bezoek'])},
            {"type": "body", "key": "{{5}}", "value": safe_str(row['Tijdvak'])},
            {"type": "body", "key": "{{6}}", "value": safe_str(row['Reparatieduur'])},
            {"type": "body", "key": "{{7}}", "value": safe_str(row['DP Nummer'])}
        ]

        open_tickets = [t for t in ticket_lookup.get(werkbon, []) if t['is_open']]
        if len(open_tickets) == 1:
            send_message_with_ticket(open_tickets[0]['ticket_id'], params)
        else:
            new_ticket_id = send_new_whatsapp_message(mobiel, params)
            if new_ticket_id:
                set_custom_fields(new_ticket_id, {
                    CUSTOM_FIELDS['locatie']: safe_str(row['Locatie']),
                    CUSTOM_FIELDS['element']: safe_str(row['Element']),
                    CUSTOM_FIELDS['defect']: safe_str(row['Defect']),
                    CUSTOM_FIELDS['werkbonnummer']: werkbon,
                    CUSTOM_FIELDS['binnen_of_buiten']: safe_str(row['Binnen of buiten'])
                })

def main():
    missing = [v for v in REQUIRED_ENV_VARS if not os.getenv(v)]
    if missing:
        print(f"❌ Missende ENV vars: {', '.join(missing)}")
        sys.exit(1)

    print("🔄 Ophalen recent Trengo tickets...")
    ticket_lookup = fetch_recent_trengo_tickets()
    print(f"✅ {sum(len(t) for t in ticket_lookup.values())} tickets opgehaald voor {len(ticket_lookup)} unieke werkbonnummers")

    outlook = OutlookClient()
    sender = os.getenv('TEST_EMAIL')
    subject = os.getenv('SUBJECT_LINE_PW_BEVESTIGING')
    filepath = outlook.download_excel_attachment(sender, subject)

    if filepath:
        try:
            process_excel_file(filepath, ticket_lookup)
        finally:
            os.remove(filepath)
            print(f"🧹 Tijdelijk bestand verwijderd: {filepath}")
    else:
        print("📭 Geen nieuwe Excel bijlage gevonden")

if __name__ == "__main__":
    main()
