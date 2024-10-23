import os
import requests
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File
import pandas as pd
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
from io import BytesIO
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class SharePointConnection:
    def __init__(self):
        # Load credentials from environment variables
        self.site_url = os.environ.get('SHAREPOINT_SITE_URL', 'https://boostix.sharepoint.com/sites/BoostiX')
        self.client_id = os.environ.get('SHAREPOINT_CLIENT_ID')
        self.client_secret = os.environ.get('SHAREPOINT_CLIENT_SECRET')
        
        if not all([self.client_id, self.client_secret]):
            raise ValueError("Missing required environment variables for SharePoint authentication")

    def get_context(self):
        """Create authenticated SharePoint context"""
        try:
            auth_context = AuthenticationContext(self.site_url)
            auth_context.acquire_token_for_app(
                client_id=self.client_id, 
                client_secret=self.client_secret
            )
            return ClientContext(self.site_url, auth_context)
        except Exception as e:
            logger.error(f"Failed to create SharePoint context: {str(e)}")
            raise

    def get_sharepoint_data(self, file_path):
        """Retrieve data from SharePoint Excel file"""
        try:
            ctx = self.get_context()
            logger.info(f"Retrieving file: {file_path}")
            
            response = File.open_binary(ctx, file_path)
            
            if not response.ok:
                logger.error(f"Failed to retrieve file. Status: {response.status_code}")
                raise Exception(f"SharePoint file retrieval failed: {response.status_code}")
                
            bytes_file_obj = BytesIO(response.content)
            df = pd.read_excel(bytes_file_obj)
            
            logger.info(f"Successfully retrieved {len(df)} rows of data")
            return df
            
        except Exception as e:
            logger.error(f"Error retrieving SharePoint data: {str(e)}")
            raise

class WhatsAppMessenger:
    def __init__(self):
        self.api_key = os.environ.get('TRENGO_API_KEY')
        if not self.api_key:
            raise ValueError("Missing Trengo API key")
        
        self.base_url = "https://app.trengo.com/api/v2/wa_sessions"
        self.headers = {
            "accept": "application/json",
            "content-type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }

    def send_message(self, phone_number, naam, dag, tijdvak):
        """Send WhatsApp message via Trengo"""
        payload = {
            "recipient_phone_number": phone_number,
            "hsm_id": 181327,
            "params": [
                {"type": "body", "key": "{{1}}", "value": str(naam)},
                {"type": "body", "key": "{{2}}", "value": str(dag)},
                {"type": "body", "key": "{{3}}", "value": str(tijdvak)}
            ]
        }

        try:
            logger.info(f"Sending message to {phone_number} for {naam}")
            response = requests.post(
                self.base_url, 
                json=payload, 
                headers=self.headers,
                timeout=30  # Add timeout
            )
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Error sending WhatsApp message: {str(e)}")
            raise

class DataProcessor:
    def __init__(self):
        self.sharepoint = SharePointConnection()
        self.messenger = WhatsAppMessenger()
        
    def process_data(self):
        """Main processing function"""
        logger.info(f"Starting new processing cycle: {datetime.now()}")
        
        try:
            # Configure your SharePoint file path
            file_path = "/sites/BoostiX/Gedeelde%20documenten/Projecten/06.%20Fixzed/4.%20Data/TestingTrengo/trengotest.xlsx"
            df = self.sharepoint.get_sharepoint_data(file_path)
            
            if df.empty:
                logger.info("No data found to process")
                return
            
            # Validate required columns
            required_columns = ['naam', 'dag', 'tijdvak', 'telefoon']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Missing required columns: {missing_columns}")
            
            # Process each row
            for index, row in df.iterrows():
                try:
                    # Clean phone number format
                    phone = str(row['telefoon']).strip()
                    if not phone.startswith('31'):
                        phone = '31' + phone.lstrip('0')
                        
                    self.messenger.send_message(
                        phone_number=phone,
                        naam=row['naam'],
                        dag=row['dag'],
                        tijdvak=row['tijdvak']
                    )
                    logger.info(f"Successfully processed message for {row['naam']}")
                    
                except Exception as e:
                    logger.error(f"Error processing row {index}: {str(e)}")
                    continue
                    
        except Exception as e:
            logger.error(f"General processing error: {str(e)}")
            raise

def main():
    # Create processor instance
    processor = DataProcessor()
    
    # Run initial processing
    logger.info("Starting initial processing...")
    processor.process_data()
    
    # Set up scheduler
    scheduler = BlockingScheduler()
    scheduler.add_job(processor.process_data, 'interval', minutes=30)
    
    logger.info("Starting scheduler...")
    scheduler.start()

if __name__ == "__main__":
    main()
