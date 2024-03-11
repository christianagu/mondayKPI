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
async def gather_and_map_data(monday_projects):
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
        projects_by_frequency = await monday_projects.group_projects_by_frequency(projects_by_region)
        print("Complete...\n")

        print("\nGathering KPI Stats for Slides Presentation...\n")
        kpi_by_month, kpi_by_quarter = gather_kpi_stats = await monday_projects.gather_kpi_stats(projects_by_frequency)
        print("Complete...\n")

        return kpi_by_month, kpi_by_quarter


async def create_update_slides(google_slides):
    slide_request = [
            {
                'createSlide': {
                    'objectId': 'Q1_BY_REGION',  # Google Slides will generate an ID if not specified
                    'insertionIndex': '1',
                    'slideLayoutReference': {
                        'predefinedLayout': 'TITLE_ONLY'
                    }
                }
            }
        ]

    create_slide_response = await google_slides.create_request(slide_request)        
    # Get the ID of the newly created slide
    slide_object_id = create_slide_response.get('replies')[0].get('createSlide').get('objectId')
    print(f"Created slide ID: {slide_object_id}")
    # Assuming 'service' and 'presentation_id' are already defined
    slide = google_slides.get_slide().get('slides')[1]
    
    # Find the title placeholder
    for obj in slide['pageElements']:
        if obj['shape']['shapeType'] == 'TEXT_BOX' and 'placeholder' in obj['shape']:
            if obj['shape']['placeholder']['type'] == 'TITLE':
                title_placeholder_object_id = obj['objectId']
                break
    # Insert the title into the newly created slide
    insert_title_requests = [
        {
            'insertText': {
                'objectId': title_placeholder_object_id,
                'text': 'Quarter 1 by Region',
                'insertionIndex': 0
            }
        }
    ]
    create_title = await google_slides.create_request(insert_title_requests)
    print("Title inserted.")
    
    # Create  table requests
    create_table_requests = [
        {
            'createTable': {
                'objectId': 'q1RegionTable',
                'elementProperties': {
                    'pageObjectId': slide_object_id,
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
    create_table_response = await google_slides.create_request(create_table_requests)
    print("Table Created.")

async def kpi_to_matrix(google_slides, kpi_by_month, kpi_by_quarter):
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

    for quarter, regions_data in kpi_by_quarter.items():
        print(f"Quarter: {quarter}\n")
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
            # Each cell is targeted by row and column for text insertion
            requests.append({
                'insertText': {
                    'objectId': 'q1RegionTable',
                    'cellLocation': {
                        'rowIndex': row_index,
                        'columnIndex': col_index,
                    },
                    'text': text,
                    'insertionIndex': 0, # Insert at the beginning of the cell
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
        kpi_by_month, kpi_by_quarter = await gather_and_map_data(monday_projects)
        
        await create_update_slides(google_slides)
        await kpi_to_matrix(google_slides, kpi_by_month, kpi_by_quarter)
        

    except Exception as e:
        print(f'An error occurred: {e}')

if __name__ == '__main__':
    asyncio.run(main())
