from dotenv import load_dotenv
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaFileUpload
load_dotenv()  # Load environment variables
from googleapiclient.errors import HttpError

class GoogleBase:
    def __init__(self, google_scopes,
                 google_app, google_app_version, presentation_id):
    # Call Google Slides
        SERVICE_ACCOUNT_FILE = 'integrationskpi-93d104cec721.json'
        self.creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=google_scopes)
        self.service = build(google_app, google_app_version, credentials=self.creds)
    

class GoogleDrive(GoogleBase):
    async def create_presentation(self, presentation_name, parent_folder_id):
        """Create a Google Slides presentation in a specific Drive folder and return its ID and name."""
        file_metadata = {
            'name': presentation_name,
            'mimeType': 'application/vnd.google-apps.presentation',
            'parents': [parent_folder_id]  # Specify the folder ID here
        }
        file = self.service.files().create(body=file_metadata,
                                    fields='id, name').execute()
        
        # Return a dictionary with the presentation name as key and ID as value
        print(file.get('name'), file.get('id'))
        return file
    
    async def delete_file(self, file_id):
        """Delete a file from Google Drive identified by file_id."""
        try:
            self.service.files().delete(fileId=file_id).execute()
            print(f"File {file_id} deleted successfully.")
        except HttpError as error:
            print(f"An error occurred: {error}")

    async def list_folders(self):
        """Lists the first 100 folders the Service Account has access to."""
        try:
            results = self.service.files().list(
                pageSize=100,
                fields="nextPageToken, files(id, name)",
                q="mimeType = 'application/vnd.google-apps.folder' and trashed = false"
            ).execute()
            folders = results.get('files', [])

            if not folders:
                print("No folders found.")
            else:
                print("Folders:")
                for folder in folders:
                    print(f"{folder['name']} ({folder['id']})")
            return folders
        except HttpError as error:
            print(f"An error occurred: {error}")

    async def list_files(self):
        try:
            # Call the Drive v3 API
            results = (
                self.service.files()
                .list(pageSize=10, fields="nextPageToken, files(id, name)")
                .execute()
            )
            items = results.get('files', [])
            if not items:
                print("No files found.")
                return
            print("Files:")
            for item in items:
                print(f"{item['name']} ({item['id']})")

            return items

        except HttpError as error:
            # TODO(developer) - Handle errors from drive API.
            print(f"An error occurred: {error}")

    async def add_permissions(self, file_id, email, transfer_ownership, role='writer', type='user'):
        if not email:  # Check if email is None or empty
            print("Email address is required for user or group permissions.")
            return

        permission = {
            'type': type,
            'role': role,
            'emailAddress': email
        }
        try:
            self.service.permissions().create(
                fileId=file_id,
                body=permission,
                fields='id',
                sendNotificationEmail= True if transfer_ownership else False,
                transferOwnership = transfer_ownership,
            ).execute()
            print(f"Permission added: {email} as {role}")
        except HttpError as error:
            print(f"An error occurred: {error}")

    async def check_for_existing_docs(self, data, existing_folder, existing_file):
        try:
            if existing_file:
                items = data
                if not items:
                    print("No files found.")
                    return
                for item in items:
                    if(existing_file['name'] == item['name'] or existing_file['id'] == item['id']):
                        print(f"Existing File found: {item['name']} ({item['id']})")
                        return item
            elif existing_folder:
                items = data
                if not items:
                    print("No Folders found.")
                    return
                for item in items:
                    if(existing_folder['name'] == item['name'] or existing_folder['id'] == item['id']):
                        print(f"Existing Folder found: {item['name']} ({item['id']})")
                        return item

        except HttpError as error:
            # TODO(developer) - Handle errors from drive API.
            print(f"An error occurred: {error}")

class GoogleSlides(GoogleBase):
        def __init__(self, google_scopes, google_app, google_app_version, presentation_id):
            super().__init__(google_scopes, google_app, google_app_version)
            self.presentation_id = presentation_id

    # Get slide data
    def get_slide(self):
        return self.service.presentations().get(presentationId=self.presentation_id).execute()
    
    # Util to help with finding the title_place_holder. Should maybe be in gsuites.py
    def find_title_placeholder(self, create_slide_response, slide_index):
        # Get the ID of the newly created slide
        slide_object_id = create_slide_response.get('replies')[0].get('createSlide').get('objectId')
        print(f"Created slide ID: {slide_object_id}")
        # Assuming 'service' and 'presentation_id' are already defined
        slide = self.get_slide().get('slides')[int(slide_index)]
        
        # Find the title placeholder
        for obj in slide['pageElements']:
            if obj['shape']['shapeType'] == 'TEXT_BOX' and 'placeholder' in obj['shape']:
                if obj['shape']['placeholder']['type'] == 'TITLE':
                    title_placeholder_object_id = obj['objectId']
                    break

        return title_placeholder_object_id
    
    # Standard create request
    async def batch_update_request(self, slide_request):
        batch_request = self.service.presentations().batchUpdate(
                presentationId=self.presentation_id, body={'requests': slide_request}
            ).execute()
        return batch_request
    
    # Create slide request
    async def create_slide(self, slide_object_id, insertion_index, predefined_layout):
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
        return await self.batch_update_request(slide_request) 

    # insert title requests
    async def insert_title(self, title_placeholder_object_id, title_text, title_insertion_index):
        insert_title_requests = [
        {
            'insertText': {
                'objectId': title_placeholder_object_id,
                'text': title_text,
                'insertionIndex': title_insertion_index
            }
        }
    ]
        return await self.batch_update_request(insert_title_requests)

    # Create table request
    async def create_table(self, table_object_id, slide_object_id):
        # Create  table requests
        create_tables = [
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
        return await self.batch_update_request(create_tables)
    
    async def setup_slides(self, slide_object_id, insertion_index, predefined_layout):
        ###
        """Create Quarterly Slides"""
        ###
        # Slide Parameters

        ## Create slide function called using parameters above
        create_slide_response = await self.create_slide(slide_object_id, insertion_index, predefined_layout)
        
        # Title parameters
        slide_index = '1'
        title_placeholder_object_id = self.find_title_placeholder(create_slide_response, slide_index)
        title_text = 'Quarter by Region'
        title_insertion_index = 0
        

        # Insert the title into the newly created slide
        create_title = await self.insert_title(title_placeholder_object_id, title_text, title_insertion_index)
        print("Title inserted.")
        
        # Table parameters
        monthly_table_by_region = 'q1RegionTable'

        # Create table with passed parameters
        create_table_response = await self.create_table(monthly_table_by_region, slide_object_id)

        ###
        """Create Monthly Slides"""
        ###
        # Slide Parameters
        slide_object_id = 'MONTH_BY_REGION'
        insertion_index = '2'
        predefined_layout = 'TITLE_ONLY'

        ## Create slide function called using parameters above
        create_slide_response = await self.create_slide(slide_object_id, insertion_index, predefined_layout)
        
        # Title parameters
        slide_index = '2'
        title_placeholder_object_id = self.find_title_placeholder(create_slide_response, slide_index)
        title_text = 'Quarter 1 Breakdown'
        title_insertion_index = 0
        
        # Insert the title into the newly created slide
        create_title = await self.insert_title(title_placeholder_object_id, title_text, title_insertion_index)
        print("Title inserted.")
        
        # Table Parameters
        monthly_table_by_region = 'q1MonthlyRegionTable'

        # Create table with passed parameters
        create_table_response = await self.create_table(monthly_table_by_region, slide_object_id)
        print("Table Created.")