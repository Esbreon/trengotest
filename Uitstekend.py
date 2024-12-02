from flask import Flask, jsonify
import os
import requests
import json

app = Flask(__name__)

class TrengoClient:
    def __init__(self):
        self.api_key = os.getenv('TRENGO_API_KEY')  # Zorg ervoor dat deze omgevingsvariabele is ingesteld
        self.base_url = "https://app.trengo.com/api/v2"

    def get_data(self):
        """Haal gegevens op uit Trengo (bijvoorbeeld aantal tickets)."""
        url = f"{self.base_url}/tickets"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }

        try:
            response = requests.get(url, headers=headers)

            if response.status_code == 200:
                tickets = response.json()
                return len(tickets)  # Aantal tickets
            else:
                print(f"Fout bij ophalen data uit Trengo: {response.status_code}")
                return None
        except requests.exceptions.RequestException as e:
            print(f"Fout bij het maken van de API-aanroep naar Trengo: {e}")
            return None

def create_json_for_smiirl(ticket_count):
    """Zet de opgehaalde data om naar een JSON die geschikt is voor de Smiirl counter."""
    data = {
        "count": ticket_count  # De data die we naar de Smiirl counter sturen
    }
    return json.dumps(data)

@app.route('/counter', methods=['GET'])
def counter():
    """API endpoint die JSON retourneert voor de Smiirl counter."""
    trengo_client = TrengoClient()
    ticket_count = trengo_client.get_data()
    
    if ticket_count is not None:
        return create_json_for_smiirl(ticket_count), 200
    else:
        return jsonify({"error": "Data kon niet worden opgehaald uit Trengo"}), 500

if __name__ == '__main__':
    # Haal de juiste poort op voor Heroku (via de omgevingsvariabele PORT)
    port = int(os.environ.get('PORT', 5000))  # Gebruik 5000 lokaal als fallback
    app.run(host='0.0.0.0', port=port)  # Luister op 0.0.0.0 zodat Heroku het kan bereiken
