import os
import requests
import time
import base64

def create_encoded_custom_field(location="fixzed-a", email="", planregel=""):
    """
    Creates a base64 encoded string from the custom field parameters.
    
    Args:
        location (str): Location identifier (default: "fixzed-a")
        email (str): User's email address
        planregel (str): Plan rule identifier
    
    Returns:
        str: Base64 encoded string of the combined parameters
    """
    # Combine the parameters with commas
    combined_string = f"{location},{email},{planregel}"
    # Convert to bytes and encode to base64
    encoded_bytes = base64.b64encode(combined_string.encode('utf-8'))
    # Convert bytes back to string for API transmission
    return encoded_bytes.decode('utf-8')

def send_initial_template_message(email, planregel):
    """
    Sends the initial WhatsApp template message and immediately updates the custom field.
    This function handles both the template sending and field update in one sequential flow.
    
    Args:
        email (str): User's email address
        planregel (str): Plan rule identifier
    
    Returns:
        str or None: The ticket ID if successful, None if any step fails
    """
    # API endpoint for creating new WhatsApp sessions
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    # Set up the template message payload with our test user information
    # Now including email and planregel in the template parameters
    template_payload = {
        "recipient_phone_number": "+31653610195",
        "hsm_id": os.environ.get('WHATSAPP_TEMPLATE_ID_PLAN'),
        "params": [
            {"type": "body", "key": "{{1}}", "value": "Tris"},
            {"type": "body", "key": "{{2}}", "value": email},
            {"type": "body", "key": "{{3}}", "value": planregel}
        ]
    }
    
    # Headers required for all Trengo API requests
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {os.environ.get('TRENGO_API_KEY')}"
    }
    
    try:
        # Step 1: Send the template message
        template_response = requests.post(url, json=template_payload, headers=headers)
        template_response.raise_for_status()
        print("Template message sent successfully")
        
        # Extract the ticket ID from the response
        ticket_id = template_response.json().get('message', {}).get('ticket_id')
        
        if ticket_id:
            print(f"Ticket ID received: {ticket_id}")
            
            # Step 2: Update the custom field
            custom_field_url = f"https://app.trengo.com/api/v2/tickets/{ticket_id}/custom_fields"
            
            # Create the base64 encoded custom field value
            encoded_value = create_encoded_custom_field(
                email=email,
                planregel=planregel
            )
            
            # Payload for updating the custom field with encoded value
            custom_field_payload = {
                "custom_field_id": 618842,
                "value": encoded_value
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
        print(f"Error occurred: {str(e)}")
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
    
    # Example usage with the provided email and planregel
    email = "tristan@boostix.nl"
    planregel = "1111"
    
    # Send the template and update the custom field
    ticket_id = send_initial_template_message(email, planregel)
    
    if ticket_id:
        print(f"Process completed successfully. Ticket ID: {ticket_id}")
    else:
        print("Process failed to complete successfully")

if __name__ == "__main__":
    main()
