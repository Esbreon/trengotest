import os
import requests
import time

def send_initial_template_message():
    """
    Sends the initial WhatsApp template message and immediately updates the custom field.
    This function handles both the template sending and field update in one sequential flow.
    
    Returns:
        str or None: The ticket ID if successful, None if any step fails
    """
    # API endpoint for creating new WhatsApp sessions
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    # Set up the template message payload with our test user information
    template_payload = {
        "recipient_phone_number": "+31653610195",
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_PLAN'),
        "params": [
            {"type": "body", "key": "{{1}}", "value": "Tris"}
        ]
    }
    
    # Headers required for all Trengo API requests
    # We use the same headers for both operations to maintain consistency
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }
    
    try:
        # Step 1: Send the template message
        template_response = requests.post(url, json=template_payload, headers=headers)
        template_response.raise_for_status()  # This will raise an exception for error status codes
        print("Template message sent successfully")
        
        # Extract the ticket ID from the response
        # We need this ID to update the custom field
        ticket_id = template_response.json().get('message', {}).get('ticket_id')
        
        if ticket_id:
            print(f"Ticket ID received: {ticket_id}")
            
            # Step 2: Update the custom field
            # We construct the URL for the custom fields endpoint using the ticket ID
            custom_field_url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}/custom_fields"
            
            # Payload for updating the custom field
            # Using the specific custom field ID (613776) for 'Locatie'
            custom_field_payload = {
                "custom_field_id": 613776,
                "value": "test"
            }
            
            # Send the request to update the custom field
            field_response = requests.post(
                custom_field_url, 
                json=custom_field_payload, 
                headers=headers
            )
            field_response.raise_for_status()
            print("Custom field updated successfully")
            
            return ticket_id
        else:
            print("No ticket ID received in template response")
            return None
            
    except requests.exceptions.RequestException as e:
        # Comprehensive error handling to catch any API-related issues
        print(f"Error occurred: {str(e)}")
        # If we have a response object, print its content for debugging
        if hasattr(e, 'response') and e.response is not None:
            print(f"Response content: {e.response.text}")
        return None

def main():
    """
    Main function to orchestrate the process of sending the template
    and updating the custom field.
    """
    print("Starting WhatsApp template message process...")
    
    # Verify that required environment variables are set
    if not os.environ.get('TRENGO_API_KEY'):
        print("Error: TRENGO_API_KEY environment variable is not set")
        return
    
    if not os.environ.get('WHATSAPP_TEMPLATE_ID_PLAN'):
        print("Error: WHATSAPP_TEMPLATE_ID_PLAN environment variable is not set")
        return
    
    # Send the template and update the custom field
    ticket_id = send_initial_template_message()
    
    if ticket_id:
        print(f"Process completed successfully. Ticket ID: {ticket_id}")
    else:
        print("Process failed to complete successfully")

if __name__ == "__main__":
    main()
