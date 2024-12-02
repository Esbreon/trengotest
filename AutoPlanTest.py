import os
import requests
import time

def send_initial_template_message():
    """
    Sends the initial WhatsApp template message to our test number.
    Returns the ticket ID which we'll use to track the conversation.
    """
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    # Set up the template message for our test user
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

def update_planning_url_field(ticket_id):
    """
    Updates the custom field 'URL Zelfstandig plannen' on the ticket with value 'test'.
    This happens after we receive a positive response.
    """
    url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}"
    
    payload = {
        "custom_fields": {
            "Locatie": "test"
        }
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }
    
    response = requests.patch(url, json=payload, headers=headers)
    response.raise_for_status()
    print("Updated 'URL Zelfstandig plannen' field to 'test'")

def get_latest_customer_message(ticket_id):
    """
    Retrieves the most recent message from the customer in this conversation.
    """
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
    """
    Checks for customer responses and updates the custom field if a positive response is received.
    """
    latest_message = get_latest_customer_message(ticket_id)
    
    if latest_message:
        response_text = latest_message.get('body', '').strip()
        
        if "Ja" in response_text:
            print(f"Positive response received: '{response_text}'")
            update_planning_url_field(ticket_id)
            return True
        elif "Nee" in response_text:
            print(f"Negative response received: '{response_text}'")
            return True
            
    return False

if __name__ == "__main__":
    # Send the initial template message and get started
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
