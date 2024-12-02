import os
import requests
import time

def send_initial_template_message():
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    payload = {
        "recipient_phone_number": "+31653610195",
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_PLAN'),
        "params": [
            {"type": "body", "key": "{{1}}", "value": "Tris"}
        ]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }
    
    response = requests.post(url, json=payload, headers=headers)
    response.raise_for_status()
    print(f"Trengo response: {response.text}")
    return response.json().get('message', {}).get('ticket_id')

def send_followup_message(ticket_id):
    url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}/messages"
    
    payload = {
        "body": "Bedankt voor uw reactie! Hier is de link om een afspraak in te plannen: [planningslink]",
        "type": "whatsapp"
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }
    
    response = requests.post(url, json=payload, headers=headers)
    response.raise_for_status()
    print("Follow-up message sent successfully")

def get_latest_customer_message(ticket_id):
    url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}/messages"
    
    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    messages = response.json().get('data', [])
    
    customer_messages = [m for m in messages if m.get('from_contact')]
    return customer_messages[-1] if customer_messages else None

def monitor_ticket_response(ticket_id):
    latest_message = get_latest_customer_message(ticket_id)
    
    if latest_message:
        response_text = latest_message.get('body', '').strip()
        
        if response_text == 'Ja':
            print("'Ja' response received, sending follow-up...")
            send_followup_message(ticket_id)
            return True
        elif response_text == 'Nee':
            print("'Nee' response received")
            return True
            
    return False

if __name__ == "__main__":
    ticket_id = send_initial_template_message()
    print(f"Monitoring ticket ID: {ticket_id}")
    
    if ticket_id:
        max_attempts = 30  # 5 minutes with 10-second intervals
        for attempt in range(max_attempts):
            if monitor_ticket_response(ticket_id):
                break
            print(f"Waiting for response... (attempt {attempt + 1}/{max_attempts})")
            time.sleep(10)
    else:
        print("No ticket ID received, cannot monitor responses")
