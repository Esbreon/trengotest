import os
import sys
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
        filter_query = (
            f"from/emailAddress/address eq '{sender_email}' "
            f"and subject eq '{subject_line}' and isRead eq false"
        )
        url = 'https://graph.microsoft.com/v1.0/me/messages'
        params = {'$filter': filter_query, '$select': 'id,hasAttachments'}

        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        for msg in resp.json().get('value', []):
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
                    # mark as read
                    requests.patch(
                        f"https://graph.microsoft.com/v1.0/me/messages/{msg['id']}",
                        headers=headers,
                        json={"isRead": True}
                    )
                    return filepath
        return None

def fetch_recent_trengo_tickets():
    token = os.getenv("TRENGO_API_KEY")
    tickets_by_werkbon = {}
    next_url = f"https://app.trengo.com/api/v2/tickets?page=1&per_page={PER_PAGE}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    page = 1
    count = 0

    while next_url and count < MAX_TICKETS:
        resp = requests.get(next_url, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        for t in data.get("data", []):
            count += 1
            if count > MAX_TICKETS:
                break
            # only need werkbon filter here
            cf = t.get("custom_field_values", [])
            for f in cf:
                if str(f.get("custom_field_id")) == str(CUSTOM_FIELDS["werkbonnummer"]):
                    wb = str(f.get("value", "")).strip()
                    if wb:
                        tickets_by_werkbon.setdefault(wb, []).append(t["id"])
        next_url = data.get("links", {}).get("next")
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
        for fmt in ('%Y-%m-%d','%d/%m/%Y','%d-%m-%Y'):
            try:
                d = datetime.strptime(str(date_str), fmt)
                nl = ['januari','februari','maart','april','mei','juni','juli',
                      'augustus','september','oktober','november','december']
                return f"{d.day} {nl[d.month-1]} {d.year}"
            except: pass
        return str(date_str)
    except:
        return str(date_str)

def format_phone(phone):
    if pd.isna(phone):
        return None
    p = str(phone).strip()
    return p.split('.')[0] if p.endswith('.0') else p

def send_new_whatsapp_message(phone, params):
    url = "https://app.trengo.com/api/v2/wa_sessions"
    headers = {
        "Authorization": f"Bearer {os.getenv('TRENGO_API_KEY')}",
        "Content-Type": "application/json"
    }
    payload = {
        "recipient_phone_number": phone,
        "hsm_id": os.getenv('WHATSAPP_TEMPLATE_ID_TEST_BEVESTIGING'),
        "params": params
    }
    r = requests.post(url, json=payload, headers=headers)
    r.raise_for_status()
    tid = r.json().get("message", {}).get("ticket_id")
    print(f"✅ New ticket {tid} created for {phone}")
    return tid

def set_custom_fields(ticket_id, fields):
    headers = {
        "Authorization": f"Bearer {os.getenv('TRENGO_API_KEY')}",
        "Content-Type": "application/json"
    }
    for fid, val in fields.items():
        r = requests.post(
            f"https://app.trengo.com/api/v2/tickets/{ticket_id}/custom_fields",
            json={"custom_field_id": fid, "value": val},
            headers=headers
        )
        r.raise_for_status()

def merge_tickets(main_id, merge_ids):
    url = f"https://app.trengo.com/api/v2/tickets/{main_id}/merge"
    headers = {
        "Authorization": f"Bearer {os.getenv('TRENGO_API_KEY')}",
        "Content-Type": "application/json"
    }
    payload = {"ticket_ids": merge_ids}
    r = requests.post(url, json=payload, headers=headers)
    r.raise_for_status()
    print(f"🔀 Merged tickets {merge_ids} into {main_id}")

def process_excel_file(filepath, ticket_lookup):
    df = pd.read_excel(filepath)
    print(f"📄 Rows in Excel: {len(df)}")
    for _, row in df.iterrows():
        wb = safe_str(row['Werkbonnummer'])
        phone = format_phone(row['Mobielnummer'])
        if not (wb and phone):
            continue

        # build template params
        params = [
            {"type":"body","key":"{{1}}","value":safe_str(row['Naam bewoner'])},
            {"type":"body","key":"{{2}}","value":safe_str(row['Taaktype'])},
            {"type":"body","key":"{{3}}","value":safe_str(row['Dag'])},
            {"type":"body","key":"{{4}}","value":format_date(row['Datum bezoek'])},
            {"type":"body","key":"{{5}}","value":safe_str(row['Tijdvak'])},
            {"type":"body","key":"{{6}}","value":safe_str(row['Reparatieduur'])},
            {"type":"body","key":"{{7}}","value":wb}
        ]

        existing = ticket_lookup.get(wb, [])
        # 1 existing → we will merge
        new_tid = send_new_whatsapp_message(phone, params)
        if new_tid:
            set_custom_fields(new_tid, {
                CUSTOM_FIELDS['locatie']: safe_str(row['Locatie']),
                CUSTOM_FIELDS['element']: safe_str(row['Element']),
                CUSTOM_FIELDS['defect']: safe_str(row['Defect']),
                CUSTOM_FIELDS['werkbonnummer']: wb,
                CUSTOM_FIELDS['binnen_of_buiten']: safe_str(row['Binnen of buiten'])
            })
            if len(existing) == 1:
                merge_tickets(existing[0], [new_tid])

def main():
    missing = [v for v in REQUIRED_ENV_VARS if not os.getenv(v)]
    if missing:
        print("❌ Missing ENV vars:", missing)
        sys.exit(1)

    # 1) fetch lookup
    lookup = fetch_recent_trengo_tickets()
    print(f"🔍 Found {sum(len(v) for v in lookup.values())} tickets for {len(lookup)} werkbonnummers")

    # 2) download Excel
    out = OutlookClient()
    f = out.download_excel_attachment(os.getenv('TEST_EMAIL'), os.getenv('SUBJECT_LINE_PW_BEVESTIGING'))
    if not f:
        print("📭 No new Excel attachment found")
        return

    # 3) process & merge logic
    try:
        process_excel_file(f, lookup)
    finally:
        os.remove(f)
        print("🧹 Deleted", f)

if __name__=="__main__":
    main()
