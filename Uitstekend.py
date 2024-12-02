from flask import Flask, jsonify
import os
import requests
import json

app = Flask(__name__)

class TrengoClient:
    def __init__(self):
        self.api_key = os.getenv('TRENGO_API_KEY')
        self.base_url = "https://app.trengo.com/api/v2"

    def get_data(self):
        """Haal gegevens op uit Trengo (bijvoorbeeld aantal tickets)."""
        url = f"{self.base_url}/tickets"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }

        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            tickets = response.json()
            return len(tickets)  # Aantal tickets
        else:
            print(f"Fout bij ophalen data uit Trengo: {response.status_code}")
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
        return create_json_for_smiirl(ticket_count)
    else:
        return jsonify({"error": "Data kon niet worden opgehaald uit Trengo"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)  # Je kunt Heroku configureren om dit te draaien
