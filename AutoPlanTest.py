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
    
    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        print(f"Trengo response: {response.text}")
        return response.json().get('message', {}).get('ticket_id')
    except requests.exceptions.RequestException as e:
        print(f"Error sending template message: {str(e)}")
        return None

def update_planning_url_field(ticket_id):
    """
    Updates the Locatie custom field (ID: 613776) on the ticket.
    This happens after we receive a positive response.
    
    The function uses Trengo's dedicated custom fields endpoint:
    POST /api/v2/tickets/{ticket_id}/custom_fields
    """
    # Notice how we construct the URL - it's specifically for custom fields
    url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}/custom_fields"
    
    # This is the exact payload structure Trengo expects
    payload = {
        "custom_field_id": 613776,
        "value": "test"
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }
    
    try:
        # Notice we're using POST instead of PATCH here
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        print("Successfully updated 'Locatie' field to 'test'")
        # Print response for debugging purposes
        print(f"Update response: {response.text}")
        return True
    except requests.exceptions.RequestException as e:
        print(f"Error updating custom field: {str(e)}")
        if hasattr(e, 'response') and e.response is not None:
            print(f"Response content: {e.response.text}")
        return False
        
def get_latest_customer_message(ticket_id):
    """
    Retrieves the most recent message from the customer in this conversation.
    """
    url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}/messages"
    
    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        messages = response.json().get('data', [])
        
        # Filter for customer messages and get the most recent one
        customer_messages = [m for m in messages if m.get('from_contact')]
        return customer_messages[-1] if customer_messages else None
    except requests.exceptions.RequestException as e:
        print(f"Error getting latest message: {str(e)}")
        return None

def monitor_ticket_response(ticket_id):
    """
    Checks for customer responses and updates the custom field if a positive response is received.
    Includes detailed logging to help diagnose issues.
    """
    print("\nChecking for new messages...")
    latest_message = get_latest_customer_message(ticket_id)
    
    if latest_message:
        response_text = latest_message.get('body', '').strip().lower()
        print(f"Found message: '{response_text}'")
        
        if "ja" in response_text:
            print(f"Positive response received: '{response_text}'")
            if update_planning_url_field(ticket_id):
                print("Successfully processed positive response")
            else:
                print("Failed to update custom field after positive response")
            return True
        elif "nee" in response_text:
            print(f"Negative response received: '{response_text}'")
            return True
        else:
            print(f"Message received but not a clear yes/no: '{response_text}'")
            
    else:
        print("No new messages found")
    
    return False

if __name__ == "__main__":
    # Send the initial template message and get started
    ticket_id = send_initial_template_message()
    print(f"Monitoring ticket ID: {ticket_id}")
    
    if ticket_id:
        max_attempts = 30  # 5 minutes with 10-second intervals
        success = False
        
        for attempt in range(max_attempts):
            try:
                if monitor_ticket_response(ticket_id):
                    success = True
                    break
                print(f"Waiting for response... (attempt {attempt + 1}/{max_attempts})")
                time.sleep(10)
            except Exception as e:
                print(f"Error during monitoring: {str(e)}")
                break
        
        if not success:
            print("Monitoring period ended without receiving a definitive response")
    else:
        print("No ticket ID received, cannot monitor responses")
