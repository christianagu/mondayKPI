import os
import requests
from dotenv import load_dotenv
from googleapiclient.discovery import build
from google.oauth2 import service_account

load_dotenv()  # Load environment variables

class GoogleBase:
    def __init__(self, presentation_id):
    # Call Google Slides
        SERVICE_ACCOUNT_FILE = 'integrationskpi-93d104cec721.json'
        SCOPES = ['https://www.googleapis.com/auth/presentations']
        creds = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        self.service = build('slides', 'v1', credentials=creds)
        self.presentation_id = presentation_id
    
    def send_request(self, slide_request):
        try:
            response = self.service.presentations().batchUpdate(
                presentationId=self.presentation_id, body={'requests': slide_request}
            ).execute()
            return response  # Ensure that a response is returned
        except Exception as error:  # Broad catch for demonstration; specify exceptions as needed
            print(f'Error sending request: {error}')
            raise


class GoogleSlides(GoogleBase):
    def get_slide(self):
        return self.service.presentations().get(presentationId=self.presentation_id).execute()
    
    async def create_request(self, slide_request):
        return self.send_request(slide_request)

    def update_slide(self, presentation_id, request_body):
        return self.send_request(f'/v1/presentations/{presentation_id}:batchUpdate', 'POST', request_body)
