import os
import requests
import time

def send_initial_template_message():
    """
    Sends the initial template message with Ja/Nee buttons using fixed test data.
    The message will always be sent to the same number with the same name.
    Returns the ticket ID for tracking the conversation.
    """
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    # Fixed test data
    phone_number = "+31653610195"
    name = "Tris"
    
    # Template message with our fixed test name
    payload = {
        "recipient_phone_number": phone_number,
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_PLAN'),
        "params": [
            {"type": "body", "key": "{{1}}", "value": name}
        ]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }
    
    try:
        print(f"Sending initial message to {phone_number} for {name}...")
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        
        response_data = response.json()
        ticket_id = response_data.get('ticket_id')
        print(f"Initial message sent successfully. Ticket ID: {ticket_id}")
        return ticket_id
        
    except Exception as e:
        print(f"Error sending initial message: {str(e)}")
        print(f"Response content: {response.text if 'response' in locals() else 'No response'}")
        raise

def get_latest_customer_message(ticket_id):
    """
    Gets the most recent customer message from a specific ticket.
    Specifically looks for button responses ('Ja' or 'Nee').
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
        
        # Filter for customer messages and get the latest one
        customer_messages = [m for m in messages if m.get('from_contact')]
        latest_message = customer_messages[-1] if customer_messages else None
        
        if latest_message:
            print(f"Latest customer message: {latest_message.get('body')}")
        else:
            print("No customer messages found yet")
            
        return latest_message
        
    except Exception as e:
        print(f"Error getting latest message: {str(e)}")
        print(f"Response content: {response.text if 'response' in locals() else 'No response'}")
        raise

def send_followup_message(ticket_id):
    """
    Sends a follow-up message on an existing ticket when 'Ja' is received.
    """
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
    
    try:
        print(f"Sending follow-up message on ticket {ticket_id}...")
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        print("Follow-up message sent successfully")
        
    except Exception as e:
        print(f"Error sending follow-up message: {str(e)}")
        print(f"Response content: {response.text if 'response' in locals() else 'No response'}")
        raise

def monitor_ticket_response(ticket_id):
    """
    Monitors a specific ticket for a 'Ja' or 'Nee' response.
    Returns True if a follow-up message was sent.
    """
    print(f"\nChecking ticket {ticket_id} for response...")
    
    try:
        latest_message = get_latest_customer_message(ticket_id)
        
        if latest_message:
            response_text = latest_message.get('body', '').strip()
            
            if response_text == 'Ja':
                print("'Ja' response received, sending follow-up...")
                send_followup_message(ticket_id)
                return True
            elif response_text == 'Nee':
                print("'Nee' response received, no follow-up needed")
                return True
            else:
                print(f"Received response: '{response_text}' - waiting for Ja/Nee")
                
        return False
        
    except Exception as e:
        print(f"Error monitoring ticket: {str(e)}")
        raise

# Simple test execution
if __name__ == "__main__":
    # Check for required environment variables
    if not all([
        os.environ.get('TRENGO_API_KEY'),
        os.environ.get('WHATSAPP_TEMPLATE_ID_PLAN')
    ]):
        print("ERROR: Missing required environment variables")
        exit(1)
    
    try:
        # Send initial message with hardcoded data
        ticket_id = send_initial_template_message()
        
        if ticket_id:
            print("\nWaiting for response...")
            
            # Monitor for response (check every 10 seconds for 5 minutes)
            max_attempts = 30  # 5 minutes with 10-second intervals
            for attempt in range(max_attempts):
                if monitor_ticket_response(ticket_id):
                    break
                print(f"Waiting for response... (attempt {attempt + 1}/{max_attempts})")
                time.sleep(10)
                
    except Exception as e:
        print(f"Test failed: {str(e)}")
