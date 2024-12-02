import os
import requests
import json
from typing import Dict, Any
import sys

def get_custom_fields() -> Dict[str, Any]:
    """
    Retrieves all custom fields from the Trengo API.
    
    Returns:
        Dict containing the API response with custom fields data
    
    Raises:
        requests.exceptions.RequestException: If the API request fails
        KeyError: If the API response is missing expected data
    """
    url = "https://app.trengo.com/api/v2/custom_fields"
    
    # Attempt to get API key from environment variables
    api_key = os.environ.get('TRENGO_API_KEY')
    if not api_key:
        raise ValueError("TRENGO_API_KEY environment variable is not set")
    
    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error making request to Trengo API: {str(e)}")
        raise

def display_custom_fields(fields_data: Dict[str, Any]) -> None:
    """
    Displays custom fields information in a readable format.
    
    Args:
        fields_data: Dictionary containing the custom fields data from Trengo
    """
    try:
        fields = fields_data.get('data', [])
        
        if not fields:
            print("No custom fields found.")
            return
            
        print("\nTrengo Custom Fields:")
        print("-" * 80)
        print(f"{'ID':<8} {'Name':<30} {'Type':<15} {'Required':<10}")
        print("-" * 80)
        
        for field in fields:
            field_id = field.get('id', 'N/A')
            name = field.get('name', 'N/A')
            field_type = field.get('type', 'N/A')
            required = "Yes" if field.get('required', False) else "No"
            
            print(f"{field_id:<8} {name:<30} {field_type:<15} {required:<10}")
            
        print("-" * 80)
        print(f"\nTotal custom fields: {len(fields)}")
        
    except KeyError as e:
        print(f"Error parsing fields data: {str(e)}")
        raise

def main():
    """
    Main function to fetch and display Trengo custom fields.
    """
    try:
        # Fetch the custom fields
        fields_data = get_custom_fields()
        
        # Display the fields in a formatted way
        display_custom_fields(fields_data)
        
        # Optionally save to a JSON file for reference
        with open('trengo_custom_fields.json', 'w') as f:
            json.dump(fields_data, f, indent=2)
            print("\nCustom fields data saved to 'trengo_custom_fields.json'")
            
    except Exception as e:
        print(f"\nError: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
