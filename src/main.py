import asyncio
from googleapiclient.discovery import build
from monday import MondayBoards
from gsuite import GoogleSlides
from google.oauth2 import service_account
import time
import os

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

# Util to help with finding the title_place_holder. Should maybe be in gsuites.py
def find_title_placeholder(google_slides, create_slide_response, slide_index):
    # Get the ID of the newly created slide
    slide_object_id = create_slide_response.get('replies')[0].get('createSlide').get('objectId')
    print(f"Created slide ID: {slide_object_id}")
    # Assuming 'service' and 'presentation_id' are already defined
    slide = google_slides.get_slide().get('slides')[int(slide_index)]
    
    # Find the title placeholder
    for obj in slide['pageElements']:
        if obj['shape']['shapeType'] == 'TEXT_BOX' and 'placeholder' in obj['shape']:
            if obj['shape']['placeholder']['type'] == 'TITLE':
                title_placeholder_object_id = obj['objectId']
                break

    return title_placeholder_object_id


# Set up our slides by using gsuite.py functions:
# - create slides for each data set from monday
# - Create titles for each slide
# - Create tables for each slide
async def setup_slides(google_slides):
    ###
    """Create Quarterly Slides"""
    ###
    # Slide Parameters
    slide_object_id = 'Q1_BY_REGION'
    insertion_index = '1'
    predefined_layout = 'TITLE_ONLY'

    ## Create slide function called using parameters above
    create_slide_response = await google_slides.create_slide_request(slide_object_id, insertion_index, predefined_layout)
    
    # Title parameters
    slide_index = '1'
    title_placeholder_object_id = find_title_placeholder(google_slides, create_slide_response, slide_index)
    title_text = 'Quarter by Region'
    title_insertion_index = 0
    

    # Insert the title into the newly created slide
    create_title = await google_slides.insert_title_request(title_placeholder_object_id, title_text, title_insertion_index)
    print("Title inserted.")
    
    # Table parameters
    monthly_table_by_region = 'q1RegionTable'

    # Create table with passed parameters
    create_table_response = await google_slides.create_table_request(monthly_table_by_region, slide_object_id)



    ###
    """Create Monthly Slides"""
    ###
    # Slide Parameters
    slide_object_id = 'MONTH_BY_REGION'
    insertion_index = '2'
    predefined_layout = 'TITLE_ONLY'

    ## Create slide function called using parameters above
    create_slide_response = await google_slides.create_slide_request(slide_object_id, insertion_index, predefined_layout)
    
    # Title parameters
    slide_index = '2'
    title_placeholder_object_id = find_title_placeholder(google_slides, create_slide_response, slide_index)
    title_text = 'Quarter 1 Breakdown'
    title_insertion_index = 0
    
    # Insert the title into the newly created slide
    create_title = await google_slides.insert_title_request(title_placeholder_object_id, title_text, title_insertion_index)
    print("Title inserted.")
    
    # Table Parameters
    monthly_table_by_region = 'q1MonthlyRegionTable'

    # Create table with passed parameters
    create_table_response = await google_slides.create_table_request(monthly_table_by_region, slide_object_id)
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
        response = await google_slides.create_request(requests)
        print("Table populated.")

async def main():

    try:
        monday_projects = MondayBoards()
        #presentation_id = '1UNK_F66Uc6WMB4EBQ9FYdao_GXU4E6pKTL5jkGQ1tQ4'
        presentation_id = input("Insert presentation ID...\n")
        google_slides = GoogleSlides(presentation_id)
        
        kpi_by_month, kpi_by_quarter = await process_monday_data(monday_projects)
        await setup_slides(google_slides)
        await kpi_to_slides(google_slides, kpi_by_month, kpi_by_quarter)    

    except Exception as e:
        print(f'An error occurred: {e}')

if __name__ == '__main__':
    asyncio.run(main())
