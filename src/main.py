import asyncio
from monday import MondayBoards
from gsuite import GoogleSlides, GoogleDrive
from dotenv import load_dotenv
from googleapiclient.discovery import build
from google.oauth2 import service_account
import time
import os
import datetime


"""
Read the README!
"""

# Monday
# Gather our project data and format KPI data
async def process_monday_data(monday_projects):
        try:
            os.system('cls')
        except Exception as e:
            print(f'An error occurred: {e}')
            time.sleep(4)
            os.system('clear')
        
        print("Starting...\n")
        print("Getting Project Boards from Monday GQL query...\n\n")
        gathered_project_boards = await monday_projects.get_project_board()

        # Group our projects by region to then be used in our next function to group by frequency
        print("\nGrouping our projects by region...\n")
        projects_by_region = await monday_projects.group_projects_by_region(gathered_project_boards)
        print("Complete...\n")
        

        print("\nGrouping our projects by Frequency...\n")
        # Using our grouped region data to then group by frequency (Monthly/Quarterly)
        projects_by_frequency = await monday_projects.group_projects_by_month(projects_by_region)
        print("Complete...\n")

        print("\nGathering KPI Stats for Slides Presentation...\n")
        kpi_by_month, kpi_by_quarter = gather_kpi_stats = await monday_projects.gather_kpi_stats(projects_by_frequency)
        print("Complete...\n")

        return kpi_by_month, kpi_by_quarter

# Set up our slides by using gsuite.py functions:
# - create slides for each data set from monday
# - Create titles for each slide
# - Create tables for each slide
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



# Insert data into our google slides tables that were created
async def kpi_to_slides(google_slides, kpi_by_month, kpi_by_quarter):
    slideByRegionHeaders = [
            [None,"EMEA", "APAC", "NA", "Overall"],
            ["Signed"],
            ["Started"],
            ["Canceled"],
            ["On Hold"],
            ["Completed"],
        ]
        
        # Define the mapping of row titles in slideByRegionHeaders to keys in kpi data
    title_to_key_map = {
            1: 'projects_signed',
            2: 'projects_started',
            3: 'canceled_projects',
            4: 'paused_projects',
            5: 'projects_completed',
        }
    
    data_sets = [kpi_by_month, kpi_by_quarter]
    
    for data in data_sets:
        for data_time_cadence, regions_data in data.items():
            # Initialize storage for overall counts in the quarter
            overall_counts = {key: 0 for key in title_to_key_map.values()}
            
            for title_index, data_key in title_to_key_map.items():
                # Initialize the list for this row with the title included
                row_data = [slideByRegionHeaders[title_index][0]]
                
                for region in slideByRegionHeaders[0][1:]:  # Skip the first None value
                    if region == "Overall":
                        # Append the overall counts instead of getting them from the data
                        row_data.append(str(overall_counts[data_key]))
                    else:
                        # Fetch and append the data for each region
                        region_data = regions_data.get(region, {})
                        count = region_data.get(data_key, 0)
                        row_data.append(str(count))
                        # Update the overall counts
                        overall_counts[data_key] += count
                
                # Replace the old row data in slideByRegionHeaders with the new row_data
                slideByRegionHeaders[title_index] = row_data
        
            requests = []
        # Generating requests to populate the table
        for row_index, row in enumerate(slideByRegionHeaders):
            for col_index, text in enumerate(row):
                if 'Q1' in data_time_cadence:
                    # Each cell is targeted by row and column for text insertion
                    requests.append({
                        'insertText': {
                            'objectId': 'q1RegionTable',
                            'cellLocation': {
                                'rowIndex': row_index,
                                'columnIndex': col_index,
                            },
                            'text': text,
                            'insertionIndex':0, # Insert at the beginning of the cell
                        }
                    })
                else:
                    requests.append({
                        'insertText': {
                            'objectId': 'q1MonthlyRegionTable',
                            'cellLocation': {
                                'rowIndex': row_index,
                                'columnIndex': col_index,
                            },
                            'text': text,
                            'insertionIndex':0, # Insert at the beginning of the cell
                        }
                    })
        response = await google_slides.batch_update_request(requests)
        print("Table populated.")

async def main():

    try:
        # Define the scopes
        drive_scopes = {
            'api_drive_scope': ['https://www.googleapis.com/auth/drive'],
            'google_app': 'drive',
            'google_app_version': 'v3'
        }
        presentation_scopes = {
            'api_drive_scope': ['https://www.googleapis.com/auth/presentations'],
            'google_app': 'slides',
            'google_app_version': 'v1'
        }
        google_drive = GoogleDrive(drive_scopes['api_drive_scope'], drive_scopes['google_app'], drive_scopes['google_app_version'], None)
        monday_projects = MondayBoards()

        folder_id = os.getenv('GOOGLE_SLIDES_FOLDER_ID')
        folder_name = os.getenv('GOOGLE_SLIDES_FOLDER_NAME')
        user_email = os.getenv('GOOGLE_EMAIL')

        presentation_name = str(datetime.now().year) + ' Data'
        
        existing_folder = {'name': folder_name, 'id': '' }
        existing_file = {'name': presentation_name, 'id': '' }
        #delete_id_flag = '1H6ZKXixXYExGSPrkqjSyspgbgnkDbFhDYZFIEuZw_OA'
        #await google_drive.delete_file(delete_id_flag)
        folders = await google_drive.list_folders()


        print(existing_folder)
        compared_folder = await google_drive.check_for_existing_docs(folders, existing_folder, None)
        print('Compared Folder: ', compared_folder)
        if compared_folder:
            print("Folder exists....\nGathering file data")
            #folders = await google_drive.list_folders()

            file_id = '1_ERXogtNIjuAJ9Mhn7MqGThjftC_eVz_TmSCzJSWZfA'
            await google_drive.delete_file(file_id)


            files = await google_drive.list_files()
            compared_file = await google_drive.check_for_existing_docs(files, None, existing_file)
            print(compared_file)
            if compared_file:
                print("We have an existing file.\nChecking if data needs to be updated.")
            else:
                print("No existing file found. Creating Presentation File")
                created_file_obj = await google_drive.create_presentation(presentation_name, folder_id)
                await google_drive.add_permissions(created_file_obj.get('id'), user_email, False, role='writer', type='user')
                
                print("setting up slides")
                google_slides = GoogleSlides(presentation_scopes['api_drive_scope'], presentation_scopes['google_app'], presentation_scopes['google_app_version'], created_file_obj.get('id') )
                
                
                slide_object_id = 'Q1_BY_REGION'
                insertion_index = '1'
                predefined_layout = 'TITLE_ONLY'

                await google_slides.setup_slides(slide_object_id, insertion_index, predefined_layout)
                print("\nslides finished")
                kpi_by_month, kpi_by_quarter = await process_monday_data(monday_projects)
                files = await google_drive.list_files()
                data_to_tables = await kpi_to_slides(google_slides, kpi_by_month, kpi_by_quarter)
                
        else:
            pass

    except Exception as e:
        print(f'An error occurred: {e}')

if __name__ == '__main__':
    asyncio.run(main())
