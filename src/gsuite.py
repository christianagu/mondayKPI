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

    # Get slide data
    def get_slide(self):
        return self.service.presentations().get(presentationId=self.presentation_id).execute()
    

    # Standard create request
    async def create_request(self, slide_request):
        return self.send_request(slide_request)
    
    # Create slide request
    async def create_slide_request(self, slide_object_id, insertion_index, predefined_layout):
        slide_request = [
            {
                'createSlide': {
                    'objectId': slide_object_id,  # Google Slides will generate an ID if not specified
                    'insertionIndex': insertion_index,
                    'slideLayoutReference': {
                            'predefinedLayout': predefined_layout
                        }
                    }
            }
        ]
        return self.send_request(slide_request) 

    # insert title requests
    async def insert_title_request(self, title_placeholder_object_id, title_text, title_insertion_index):
        insert_title_requests = [
        {
            'insertText': {
                'objectId': title_placeholder_object_id,
                'text': title_text,
                'insertionIndex': title_insertion_index
            }
        }
    ]
        return self.send_request(insert_title_requests)

    # Create table request
    async def create_table_request(self, table_object_id, slide_object_id):
        # Create  table requests
        create_table_requests = [
            {
                'createTable': {
                    'objectId': table_object_id, # The name of the table
                    'elementProperties': {
                        'pageObjectId': slide_object_id, # the slide object id
                        'size': {
                            'height': {'magnitude': 2500000, 'unit': 'EMU'},
                            'width': {'magnitude': 8000000, 'unit': 'EMU'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 311700,
                            'translateY': 1100000,
                            #'translateY': 1157225,
                            'unit': 'EMU'
                        }
                    },
                    'rows': 6,
                    'columns': 5
                }
            }
        ]
        return self.send_request(create_table_requests)
