import asyncio
import os
from gql import Client, gql
from gql.transport.aiohttp import AIOHTTPTransport
from dotenv import load_dotenv
from datetime import datetime
import json

load_dotenv()  # Load environment variables

def create_json_file(filename, data):
    with open(f'{filename}', 'w') as file:
            json.dump(data, file, indent=4)

class MondayBase:
    def __init__(self):
        self.endpoint = 'https://api.monday.com/v2/'
        transport = AIOHTTPTransport(
            url=self.endpoint,
            headers={
                'Authorization': f'Bearer {os.getenv("MONDAY_API_KEY")}',
                'API-Version': '2024-01'
            }
        )
        self.client = Client(transport=transport, fetch_schema_from_transport=True)

        
    async def send_request(self, query):
        try:
            return await self.client.execute_async(query)
        except Exception as e:
            print(e)
            raise

class MondayBoards(MondayBase):

    
    # Main function to query our data from Monday
    # After gathering our data, we enrich our project objects with data from the column_value object
    # At the end we delete our column value since it's no longer needed
    async def get_project_board(self):
        grouped_data = {}  # Object to store data grouped by group titles
        
        # Queries
        fields_to_gather = {
            'int_manager': 'Int Mgr',
            'csm': 'csm',
            'project_status': 'status',
            'int_type': 'int type',
            'data_points': 'data point',
            'project_creation_date': 'proj creation',
            'late_project': 'late? (w)',
            'country': 'country',
            'start_date': 'start date',
            'due_date': 'due date',
            'updated_less_than_week': 'updated <1w?',
            'erp': 'erp (old)'
        }

        groups_query = gql("""
        query GetGroups {
            boards(ids: 498075709) {
                name
                groups {
                    id
                    title
                }
            }
        }
        """)

        items_query_template = """
        query GetItemsByGroup($groupId: String!, $cursor: String) {
            boards(ids: 498075709) {
                groups(ids: [$groupId]) {
                    title
                    items_page(limit: 50, cursor: $cursor) {
                        items {
                            id
                            name
                            created_at
                            updated_at
                            column_values {
                                column {
                                    title
                                }
                                text
                                type
                                value
                            }
                        }
                        cursor
                    }
                }
            }
        }
        """

        # Execute the groups query
        groups_response = await self.send_request(groups_query)
        groups = [group for board in groups_response['boards'] for group in board['groups']]

        # Iterate through each group and fetch items
        for group in groups:
            new_cursor = None
            grouped_data[group['title']] = []
            while True:
                items_query = gql(items_query_template)
                response = await self.client.execute_async(items_query, variable_values={"groupId": group['id'], "cursor": new_cursor})
                group_items = [item for board in response['boards'] for g in board['groups'] for item in g['items_page']['items']]
                grouped_data[group['title']].extend(group_items)
                new_cursor = response['boards'][0]['groups'][0]['items_page'].get('cursor')
                if not new_cursor:
                    break

        # Process the gathered data
        for group_title, items in grouped_data.items():
            for item in items:
                item['closed_date'] = None  # Initialize closed_date
                for column_value in item['column_values']:
                    column_title = column_value['column']['title'].lower()  # Lowercase for consistency
                    # Attempt to parse the JSON string in 'value' for all column_values
                    if column_value['value']:
                        try:
                            # Parse the JSON string into a Python object
                            value_obj = json.loads(column_value['value'])
                            column_value['value'] = value_obj  # Update with parsed object
                            
                            # Specifically handle the 'Status' column
                            if column_title == 'status' and column_value['text'].lower() == 'completed':
                                # Check if 'changed_at' is present in the parsed JSON
                                if 'changed_at' in value_obj:
                                    item['closed_date'] = value_obj['changed_at']
                                    
                        except (TypeError, json.JSONDecodeError):
                            # If 'value' is not a string or not valid JSON, ignore
                            pass

                    # Handle other fields based on 'column_title' and mapping logic
                    key = next((k for k, v in fields_to_gather.items() if v.lower() == column_title), None)
                    if key:
                        item[key] = column_value['text']
                        
                        # Example of adjusting the region based on the 'country' field
                        if key == 'country':
                            country = column_value['text'].lower()
                            if country in ['united states', 'canada', 'mexico', 'brazil', 'colombia']:
                                item['region'] = 'NA'
                            elif country in ['australia', 'china', 'india', 'indonesia', 'mongolia', 'hong kong']:
                                item['region'] = 'APAC'
                            else:
                                item['region'] = 'EMEA'
                del item['column_values']

        create_json_file('raw_data/new_grouped_data.json',grouped_data)

        return grouped_data


    # Groups projects by region
    async def group_projects_by_region(self, grouped_project_boards):
        projects_by_region = {
            'NA': {
                 'Open Projects': [],
                 'Closed Projects': [],
                 'Backlog': [],
                },
            'APAC': {
                 'Open Projects': [],
                 'Closed Projects': [],
                 'Backlog': [],
            },
            'EMEA': {
                 'Open Projects': [],
                 'Closed Projects': [],
                 'Backlog': [],
            },
        }
        try: 
            # Add projects to their regions
            for group_title, items in grouped_project_boards.items():
                        for item in items:
                            if item.get('region') and item['region'] == 'NA':
                                projects_by_region['NA'][group_title].append(item)
                            elif item.get('region') and item['region'] == 'APAC':
                                projects_by_region['APAC'][group_title].append(item)
                            elif item.get('region') and item['region'] == 'EMEA':
                                projects_by_region['EMEA'][group_title].append(item)
        
        except Exception as e:
            print(f'An error occurred: {e}')

        create_json_file('raw_data/projects_by_region.json',projects_by_region)

        return projects_by_region
    


    # Helper function to add projects to our object by month/region/kpi
    def add_to_projects_by_frequency(self, projects_by_monthly_freq, month, region, status, item):
        # Initialize the month if it doesn't exist
        if month not in projects_by_monthly_freq:
            projects_by_monthly_freq[month] = {'NA': {'projects_signed': [], 'projects_started': [], 'projects_completed': [], 'canceled_projects': [], 'paused_projects': []},
                                            'APAC': {'projects_signed': [], 'projects_started': [], 'projects_completed': [], 'canceled_projects': [], 'paused_projects': []},
                                            'EMEA': {'projects_signed': [], 'projects_started': [], 'projects_completed': [], 'canceled_projects': [], 'paused_projects': []}}
        
        # Append the item to the appropriate list based on the status
        if status.lower() == 'in progress':
            projects_by_monthly_freq[month][region]['projects_started'].append(item)
        if status.lower() == 'on hold':
            projects_by_monthly_freq[month][region]['paused_projects'].append(item)
        if status.lower() == 'canceled':
            projects_by_monthly_freq[month][region]['canceled_projects'].append(item)
        if status.lower() == 'completed':
            projects_by_monthly_freq[month][region]['projects_completed'].append(item)
        if status is not None:
            projects_by_monthly_freq[month][region]['projects_signed'].append(item)

    # Helper function to sort projects by Month (Feb, Jan, March) -> (Jan, Feb, March)
    def sort_projects_by_frequency(self, projects_by_monthly_freq):
        # Sorting the months requires understanding that they are not just strings but represent dates
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        sorted_projects_by_frequency = {month: projects_by_monthly_freq.get(month, {'NA': {}, 'APAC': {}, 'EMEA': {}})
                                        for month in month_order if month in projects_by_monthly_freq}
        return sorted_projects_by_frequency

    def data_by_int_manager(self, sorted_projects_by_frequency):
        int_manager_by_month = {}
        int_manager_by_month_count = {}

        try: 
            for month, project_frequency in sorted_projects_by_frequency.items():
                            if month not in int_manager_by_month:
                                int_manager_by_month[month] = {}
                                int_manager_by_month_count[month] = {}
                            for region, group_title in project_frequency.items():
                                if region not in sorted_projects_by_frequency[month][region]:
                                    for project_type, projects in group_title.items():
                                            count = 1
                                            for project in projects:
                                                if project['int_manager'] not in int_manager_by_month[month] and project['int_manager'] is not None:
                                                    int_manager_by_month[month][project['int_manager']] = {}
                                                    int_manager_by_month_count[month][project['int_manager']] = {}
                                                if project_type not in int_manager_by_month[month][project['int_manager']]:
                                                    int_manager_by_month[month][project['int_manager']][project_type] = []
                                                    int_manager_by_month_count[month][project['int_manager']][project_type] = count
                                                int_manager_by_month[month][project['int_manager']][project_type].append(project)
                                                int_manager_by_month_count[month][project['int_manager']][project_type] += count
            return int_manager_by_month, int_manager_by_month_count
        except Exception as e:
            print(e)


    # Group projects by frequency (Monthly / Quarterly)
    async def group_projects_by_month(self, grouped_project_boards):
        projects_by_monthly_freq = {}

        try:
            # NA
            # Projects Started, On Hold, Completed&Canceled, Signed
            for region, group_title in grouped_project_boards.items():
                        for title, items in group_title.items():
                            
                            # Projects Started
                            if title == 'Open Projects':
                                for item in items:
                                    project_status = item['project_status']

                                    # Projects with 'in progress' status
                                    if item['project_status'].lower() == 'in progress':
                                        project_created_year = item['start_date'].split('-')[0] if item['start_date'] else None
                                        start_parsed_date = datetime.strptime(item['start_date'],"%Y-%m-%d")
                                        # Format the datetime object to only show the full month name
                                        project_start_month = start_parsed_date.strftime('%B')
                                        if project_created_year == '2024':
                                            self.add_to_projects_by_frequency(projects_by_monthly_freq, project_start_month, region, project_status, item)

                        # On Hold Projects
                                    if item['project_status'].lower() == 'on hold':
                                        project_created_year = item['start_date'].split('-')[0] if item['start_date'] else None
                                        start_parsed_date = datetime.strptime(item['start_date'],"%Y-%m-%d")
                                        # Format the datetime object to only show the full month name
                                        project_start_month = start_parsed_date.strftime('%B')
                                        if project_created_year == '2024':
                                            self.add_to_projects_by_frequency(projects_by_monthly_freq, project_start_month, region, project_status, item)

                            # Completed and Canceled projects
                            elif title == 'Closed Projects':
                                for item in items:
                                    project_status = item['project_status']

                                    # Check for canceled projects
                                    if item['project_status'].lower() == 'canceled':
                                        project_created_year = item['project_creation_date'].split('-')[0] if item['project_creation_date'] else None
                                        recent_updated_parsed_date = datetime.strptime(item['updated_at'],"%Y-%m-%dT%H:%M:%SZ")
                                        # Format the datetime object to only show the full month name
                                        project_closed_month = recent_updated_parsed_date.strftime('%B')
                                        if project_created_year == '2024':
                                            self.add_to_projects_by_frequency(projects_by_monthly_freq, project_closed_month, region, project_status, item)
                                    
                                    # Check for Completed Projects
                                    if item['project_status'].lower() == 'completed':
                                        project_closed_year = item['closed_date'].split('-')[0] if item['project_creation_date'] else None
                                        recent_updated_parsed_date = datetime.strptime(item['updated_at'],"%Y-%m-%dT%H:%M:%SZ")
                                        # Format the datetime object to only show the full month name
                                        project_closed_month = recent_updated_parsed_date.strftime('%B')
                                        if project_closed_year == '2024':
                                            self.add_to_projects_by_frequency(projects_by_monthly_freq, project_closed_month, region, project_status, item)
                            
                            # Signed Projects
                            elif title == 'Open Projects' or 'Backlog' or 'Closed Projects':
                                for item in items:
                                    project_status = item['project_status']
                                    # check for projects matching canceled status
                                    project_created_year = item['project_creation_date'].split('-')[0] if item['project_creation_date'] else None
                                    recent_updated_parsed_date = datetime.strptime(item['created_at'],"%Y-%m-%dT%H:%M:%SZ")
                                    # Format the datetime object to only show the full month name
                                    project_closed_month = recent_updated_parsed_date.strftime('%B')
                                    if project_created_year == '2024':
                                        self.add_to_projects_by_frequency(projects_by_monthly_freq, project_closed_month, region, project_status, item)
            
            # After populating projects_by_frequency, sort it before writing to JSON
            sorted_projects_by_frequency = self.sort_projects_by_frequency(projects_by_monthly_freq)

        except Exception as e:
            print(f'An error occurred: {e}')

        create_json_file('raw_data/projects_by_monthly_freq.json',sorted_projects_by_frequency)
        return sorted_projects_by_frequency
    
    # Gather KPI stats for projects from previous objects
    async def gather_kpi_stats(self, sorted_projects_by_frequency):
        kpi_by_month = {}
        kpi_by_quarter = {}
        int_manager_by_month = {}
        int_manager_by_quarter_count = {}
        int_manager_by_month_count = {}
        month_count = 0

        try:
            int_manager_by_month, int_manager_by_month_count = self.data_by_int_manager(sorted_projects_by_frequency)
            

            print('int_manager_by_month\n',int_manager_by_month)
            for month, project_frequency in sorted_projects_by_frequency.items():
                        if month not in kpi_by_month:
                            kpi_by_month[month] = {}
                        for region, group_title in project_frequency.items():
                            if region not in kpi_by_month[month]:
                                kpi_by_month[month][region] = {}
                            kpi_by_month[month][region]['projects_started'] = len((sorted_projects_by_frequency[month][region]['projects_started']))
                            kpi_by_month[month][region]['canceled_projects'] = len((sorted_projects_by_frequency[month][region]['canceled_projects']))
                            kpi_by_month[month][region]['projects_signed'] = len((sorted_projects_by_frequency[month][region]['projects_signed']))
                            kpi_by_month[month][region]['paused_projects'] = len((sorted_projects_by_frequency[month][region]['paused_projects']))        
                            kpi_by_month[month][region]['projects_completed'] = len((sorted_projects_by_frequency[month][region]['projects_completed']))

            for month, project_frequency in sorted_projects_by_frequency.items():
                month_count += 1
                # Determine the current quarter based on the month count
                quarter = f"Q{str((month_count - 1) // 3 + 1)}"
                
                if quarter not in kpi_by_quarter:
                    kpi_by_quarter[quarter] = {}
                
                for region, group_title in project_frequency.items():
                    if region not in kpi_by_quarter[quarter]:
                        kpi_by_quarter[quarter][region] = {'projects_started': 0, 'canceled_projects': 0, 'projects_signed': 0, 'paused_projects': 0, 'projects_completed': 0}
                    
                    # Now safely update the counts
                    kpi_by_quarter[quarter][region]['projects_started'] += len(project_frequency[region]['projects_started'])
                    kpi_by_quarter[quarter][region]['canceled_projects'] += len(project_frequency[region]['canceled_projects'])
                    kpi_by_quarter[quarter][region]['projects_signed'] += len(project_frequency[region]['projects_signed'])
                    kpi_by_quarter[quarter][region]['paused_projects'] += len(project_frequency[region]['paused_projects'])
                    kpi_by_quarter[quarter][region]['projects_completed'] += len(project_frequency[region]['projects_completed'])
                
                # Reset the month count after each quarter
                if month_count % 3 == 0:
                    month_count = 0

            month_count = 0
            for month, project_frequency in int_manager_by_month.items():
                month_count += 1
                # Determine the current quarter based on the month count
                quarter = f"Q{str((month_count - 1) // 3 + 1)}"
                
                if quarter not in int_manager_by_quarter_count:
                    int_manager_by_quarter_count[quarter] = {}
                
                for int_manager, group_title in project_frequency.items():
                    if int_manager not in int_manager_by_quarter_count[quarter]:
                        int_manager_by_quarter_count[quarter][int_manager] = {'projects_started': 0, 'canceled_projects': 0, 'projects_signed': 0, 'paused_projects': 0, 'projects_completed': 0}
                    for type, project_types in group_title.items():
                        if type not in project_frequency[int_manager]:
                            project_frequency[int_manager][type] = 0
                        int_manager_by_quarter_count[quarter][int_manager][type] += len(project_frequency[int_manager][type]) if project_frequency[int_manager][type] else 0
                   
                
                # Reset the month count after each quarter
                if month_count % 3 == 0:
                    month_count = 0
    
        except Exception as e:
            print(f'An error occured: {e}')
        
        create_json_file('raw_data/kpi_by_quarter.json',kpi_by_quarter)
        create_json_file('raw_data/kpi_by_month.json',kpi_by_month)
        create_json_file('raw_data/int_manager_by_month.json',int_manager_by_month)
        create_json_file('raw_data/int_manager_by_quarter_count.json',int_manager_by_quarter_count)
        create_json_file('raw_data/int_manager_by_month_count.json',int_manager_by_month_count)

        return kpi_by_month, kpi_by_quarter, int_manager_by_quarter_count, int_manager_by_month_count

