import requests

def test_whatsapp():
    """Test functie om één WhatsApp bericht te versturen"""
    url = "https://app.trengo.com/api/v2/wa_sessions"
    
    payload = {
        "recipient_phone_number": "31611341059",  # Test nummer
        "hsm_id": 181327,
        "params": [
            {
                "type": "body",
                "key": "{{1}}",
                "value": "Test Naam"
            },
            {
                "type": "body",
                "key": "{{2}}",
                "value": "Test Dag"
            },
            {
                "type": "body",
                "key": "{{3}}",
                "value": "Test Tijdvak"
            }
        ]
    }
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiNDk5NjQ4M2JjMDE3Y2FlYzI4ZTExYTFlN2E0NTYyODkzNmIzZTEyNzRmZDllMThmZDIxNDZkYTA1OTJiNGUyZDRiODMyMTYyZjE2OWE1NDMiLCJpYXQiOjE3Mjk0MzMwNDYsIm5iZiI6MTcyOTQzMzA0NiwiZXhwIjo0ODUzNDg0MjQ2LCJzdWIiOiI3Mjc3MzQiLCJzY29wZXMiOltdLCJhZ2VuY3lfaWQiOjMxOTc2N30.bYiBRdH_tSt3uHUSGTFANJBhSfjZo-1hRzN9SHjQf4VB4NqxsXcaFg2wZXSGlKfvMgQ10X3KG1JtbJZoDRfuUA"
    }
    
    try:
        print("Versturen van test bericht...")
        response = requests.post(url, json=payload, headers=headers)
        print(f"Status code: {response.status_code}")
        print(f"Response body: {response.text}")
        return response.json()
    except Exception as e:
        print(f"Fout bij versturen: {str(e)}")

if __name__ == "__main__":
    print("Start test verzending...")
    test_whatsapp()
    print("Test verzending voltooid.")
