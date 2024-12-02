import os
import requests
import time
import json
from pathlib import Path

class SingleRunWhatsAppSender:
    """
    A class to manage WhatsApp template message sending that ensures single execution
    through persistent state management.
    """
    
    def __init__(self):
        """Initialize the sender with configuration and state management"""
        self.state_file = Path('whatsapp_sender_state.json')
        self.headers = {
            "accept": "application/json",
            "content-type": "application/json",
            "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
        }
        
    def _load_state(self):
        """
        Load the execution state from file.
        Returns a dictionary with execution history.
        """
        if self.state_file.exists():
            try:
                with open(self.state_file, 'r') as f:
                    return json.load(f)
            except json.JSONDecodeError:
                print("Warning: State file corrupted, creating new state")
                return {"last_execution": None, "ticket_id": None}
        return {"last_execution": None, "ticket_id": None}
    
    def _save_state(self, ticket_id):
        """
        Save the execution state to prevent duplicate runs.
        
        Args:
            ticket_id (str): The ticket ID from the successful execution
        """
        state = {
            "last_execution": time.time(),
            "ticket_id": ticket_id
        }
        with open(self.state_file, 'w') as f:
            json.dump(state, f)
    
    def send_template_message(self):
        """
        Sends the WhatsApp template message and updates custom field.
        Ensures this operation only happens once through state management.
        
        Returns:
            str or None: The ticket ID if successful, None if already run or on failure
        """
        # Check if we've already run this operation
        state = self._load_state()
        if state["ticket_id"]:
            print(f"Message already sent. Previous ticket ID: {state['ticket_id']}")
            return None
            
        # API endpoint for WhatsApp sessions
        url = "https://app.trengo.com/api/v2/wa_sessions"
        
        # Template message configuration
        template_payload = {
            "recipient_phone_number": "+31611341059",
            "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_PLAN'),
            "params": [
                {"type": "body", "key": "{{1}}", "value": "Dirkteur"}
            ]
        }
        
        try:
            # Send template message
            template_response = requests.post(url, json=template_payload, headers=self.headers)
            template_response.raise_for_status()
            print("Template message sent successfully")
            
            # Extract ticket ID
            ticket_id = template_response.json().get('message', {}).get('ticket_id')
            
            if not ticket_id:
                print("No ticket ID received in template response")
                return None
                
            print(f"Ticket ID received: {ticket_id}")
            
            # Update custom field
            custom_field_url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}/custom_fields"
            custom_field_payload = {
                "custom_field_id": 618842,
                "value": "https://www.fixzed.nl/"
            }
            
            # Send custom field update request
            field_response = requests.post(
                custom_field_url,
                json=custom_field_payload,
                headers=self.headers
            )
            field_response.raise_for_status()
            print("Custom field updated successfully")
            
            # Save state to prevent future runs
            self._save_state(ticket_id)
            
            return ticket_id
            
        except requests.exceptions.RequestException as e:
            print(f"Error occurred: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response content: {e.response.text}")
            return None

def main():
    """
    Main function that coordinates the single-run WhatsApp message process
    with proper environment validation.
    """
    print("Starting WhatsApp template message process...")
    
    # Verify environment variables
    required_vars = ['TRENGO_API_KEY', 'WHATSAPP_TEMPLATE_ID_PLAN']
    missing_vars = [var for var in required_vars if not os.environ.get(var)]
    
    if missing_vars:
        print(f"Error: Missing required environment variables: {', '.join(missing_vars)}")
        return
    
    # Initialize and run the sender
    sender = SingleRunWhatsAppSender()
    ticket_id = sender.send_template_message()
    
    if ticket_id:
        print(f"Process completed successfully. Ticket ID: {ticket_id}")
    else:
        print("Process did not complete - either it was already run or encountered an error")

if __name__ == "__main__":
    main()
