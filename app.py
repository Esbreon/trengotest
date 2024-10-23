import os
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File
import pandas as pd
from io import BytesIO
import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def test_sharepoint_access():
    try:
        # Basic setup
        site_url = "https://boostix.sharepoint.com/sites/BoostiX"
        client_id = os.environ.get('SHAREPOINT_CLIENT_ID')
        client_secret = os.environ.get('SHAREPOINT_CLIENT_SECRET')
        
        logger.info(f"Attempting connection with client_id: {client_id[:4]}...")
        
        # Simpler authentication method
        credentials = ClientCredential(client_id, client_secret)
        ctx = ClientContext(site_url).with_credentials(credentials)
        
        # Test basic connection
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        logger.info(f"Successfully connected to site: {web.properties['Title']}")
        
        # Try to list root folder contents first
        root_folder = ctx.web.default_document_library().root_folder
        ctx.load(root_folder)
        ctx.execute_query()
        logger.info(f"Root folder name: {root_folder.properties['Name']}")
        
        # Try both path formats
        file_paths = [
            "/sites/BoostiX/Gedeelde%20documenten/Projecten/06.%20Fixzed/4.%20Data/TestingTrengo/trengotest.xlsx",
            "/Gedeelde%20documenten/Projecten/06.%20Fixzed/4.%20Data/TestingTrengo/trengotest.xlsx"
        ]
        
        for path in file_paths:
            try:
                logger.info(f"\nTrying path: {path}")
                file = ctx.web.get_file_by_server_relative_url(path)
                ctx.load(file)
                ctx.execute_query()
                logger.info(f"Success! File size: {file.properties['Length']} bytes")
                
                # Try to read the file
                response = File.open_binary(ctx, path)
                bytes_file_obj = BytesIO()
                bytes_file_obj.write(response.content)
                bytes_file_obj.seek(0)
                
                df = pd.read_excel(bytes_file_obj)
                logger.info(f"Successfully read Excel file. Row count: {len(df)}")
                return df
                
            except Exception as e:
                logger.error(f"Failed with this path: {str(e)}")
                continue
                
        raise Exception("Could not access file with any path format")
        
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        logger.error("Full error details:", exc_info=True)
        raise

if __name__ == "__main__":
    print("Testing SharePoint access...")
    test_sharepoint_access()
