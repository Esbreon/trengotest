from flask import Flask, jsonify, request
import os
import requests
import json
import logging
import sys

# Setup basic logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Voeg een basis route toe om te testen
@app.route('/')
def home():
    logger.debug("Home route accessed")
    return "App is running"

# Voeg error logging toe aan je bestaande route
@app.route('/counter', methods=['GET'])
def counter():
    """API endpoint die JSON retourneert voor de Smiirl counter."""
    logger.debug("Counter route accessed")
    
    try:
        trengo_client = TrengoClient()
        ticket_count = trengo_client.get_data()
        
        if ticket_count is not None:
            return create_json_for_smiirl(ticket_count), 200
        else:
            logger.error("Kon geen data ophalen uit Trengo")
            return jsonify({"error": "Data kon niet worden opgehaald uit Trengo"}), 500
    except Exception as e:
        logger.error(f"Error in counter route: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    logger.info(f"Starting app on port {port}")
    app.run(host='0.0.0.0', port=port)
