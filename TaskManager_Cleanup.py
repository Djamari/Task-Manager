import config as cfg
from notion_client import Client
from tqdm import tqdm
import json
from datetime import date, datetime
import os
import numpy as np
import pandas as pd
import xlsxwriter
from funcs import get_all_items, empty_notes
import excel2img

print("Cleaning up the Archived tasked...")

### Get items
notion = Client(auth=cfg.authenticator)
filter_items = {'and': [{'property':'Archive', 'checkbox':{'equals': True}},
                  {'property':'Cleaned', 'checkbox':{'equals': False}}
                  ]}
items_tasks = get_all_items(notion, database_id=cfg.ID_DB_Tasks, filter=filter_items)


### Clean up latest archived tasks
for item_task in tqdm(items_tasks):
    print("Cleaning " + str(item_task['properties']['Name']['title']))

    # Check for date
    if item_task['properties']['Date planned']['date'] is None:
        date_task = item_task['last_edited_time'].split('T')[0]
    else:
        date_task = item_task['properties']['Date planned']['date']['start']

    # Get icon based on notes content
    notes = notion.blocks.children.list(item_task['id'])
    if empty_notes(notes):
        emoji = '🟩'
    else:
        emoji = '📒'


    # Update task properties
    id = item_task['id']
    notion.pages.update(id, icon={'emoji':emoji}, properties={
        'Date planned': {'start': date_task, 'end': None, 'time_zone': None},
        'Cleaned': True
    })

### Log current planning ###
print("Logging the current global planning...")
filename_data = 'planning_log/planning_data.json'
if not os.path.isfile(filename_data):
    data_log = {'log_dates': []}
else:
    with open(filename_data) as f:
        data_log = json.loads(f.read())

# Gather all info; only include milestones that are not done yet
filter_milestones = {'or': [{'property':'Status', 'select':{'does_not_equal': 'Done'}},

                  ]}
items_milestones = get_all_items(notion, database_id=cfg.ID_DB_Milestones, filter=filter_milestones)
items_projects = get_all_items(notion, database_id=cfg.ID_DB_Projects)
items_stages = get_all_items(notion, database_id=cfg.ID_DB_Stages)

# Keep track of stages and their dates
stage_dates_to_be_added = {}
today_string = date.today().strftime('%Y-%m-%d')

data_visualization = {'start': {}, 'end': {}}

# Loop through all milestones
for milestone in items_milestones:
    # Get project id and name
    project_id = milestone['properties']['Project']['relation'][0]['id']
    project_name = [project['properties']['Name']['title'][0]['text']['content'] for project in items_projects if project['id'] == project_id][0]
    project_id = 'project_'  + project_id

    # Check if project exists in data. Create it if not.
    if project_id not in  data_log.keys():
        data_log[project_id] = {'project_name': project_name}
    else:
        # Update project name in case it has changed since last time
        data_log[project_id]['project_name'] = project_name


    # Get Stage id
    stage_id = milestone['properties']['Stage']['relation'][0]['id']
    stage = [stage for stage in items_stages if stage['id'] == stage_id][0]

    # Check if Stage as meaningful milestons
    if stage['properties']['Date_earliest']['rollup']['date'] is None:
        continue

    # Get stage name and dates
    stage_name = stage['properties']['Name']['title'][0]['text']['content']
    dates = stage['properties']['Date_F']['formula']['date']
    stage_dates = {'start': dates['start'],'end': dates['end'], 'log': today_string}
    stage_id =  'stage_' + stage_id

    # Store stage date
    stage_dates_to_be_added[stage_id] = stage_dates

    # Check if Stage id exists in data. Create it if not.
    if stage_id not in data_log[project_id].keys():
        data_log[project_id][stage_id] = {'stage_name': stage_name, 'stage_planned_periods': []}
    else:
        # Update stage name in case it has changed since last time
        data_log[project_id][stage_id]['stage_name'] = stage_name


    # Get milestone id, name, and dates
    milestone_id = 'milestone_' + milestone['id']
    milestone_name = milestone['properties']['Milestone']['title'][0]['text']['content']
    milestone_dates = milestone['properties']['Period']['date']
    milestone_dates = {'start': milestone_dates['start'],'end': milestone_dates['end'], 'log': today_string}

    # Check if Milestone ID exists in data. Create it if not.
    if not milestone_id in data_log[project_id][stage_id].keys():
        data_log[project_id][stage_id][milestone_id] = {'milestone_name': milestone_name, 'milestone_planned_periods': [milestone_dates]}
    else:
        # Update milestone name in case it has changed since last time
        data_log[project_id][stage_id][milestone_id]['milestone_name'] = milestone_name

        # Store date info
        data_log[project_id][stage_id][milestone_id]['milestone_planned_periods'].append(milestone_dates)

    # Add necessary data to visualization dictionary
    if project_name not in data_visualization['start'].keys():
        data_visualization['start'][project_name] = {}
        data_visualization['end'][project_name] = {}
    if stage_name not in data_visualization['start'][project_name].keys():
        data_visualization['start'][project_name][stage_name] = {}
        data_visualization['end'][project_name][stage_name] = {}
    data_visualization['start'][project_name][stage_name][milestone_name] = datetime.strptime(milestone_dates['start'], '%Y-%m-%d')
    data_visualization['end'][project_name][stage_name][milestone_name] = datetime.strptime(milestone_dates['end'], '%Y-%m-%d')


# Store all stage dates
for stage_id in stage_dates_to_be_added.keys():
    project_id = 'project_' + [stage['properties']['Project']['relation'][0]['id'] for stage in items_stages if stage['id'] in stage_id][0]
    data_log[project_id][stage_id]['stage_planned_periods'].append(stage_dates_to_be_added[stage_id])


# Update json file
with open(filename_data, "w") as outfile:
    json.dump(data_log, outfile, indent=4)

## Create "image" of current planning
excel_filename = 'planning_log/planning_visualization.xlsx'

# Open (or create) Excel file
workbook = xlsxwriter.Workbook(excel_filename)

# Craete new sheet with today's date as name
worksheet = workbook.add_worksheet(date.today().strftime('%b %d, %Y'))

# Gather date info to figure out the order
dates_flattened_start = {}
dates_flattened_end = {}
for project_name in data_visualization['start'].keys():
    for stage_name in data_visualization['start'][project_name].keys():
        for milestone_name in data_visualization['start'][project_name][stage_name].keys():
            dates_flattened_start[project_name + stage_name + milestone_name] = data_visualization['start'][project_name][stage_name][milestone_name]
            dates_flattened_end[project_name + stage_name + milestone_name] = data_visualization['end'][project_name][stage_name][milestone_name]
min_date = min(dates_flattened_start.values())
max_date = max(dates_flattened_end.values())

# Prepare formatting
format_bold_center_border = workbook.add_format(
    {
        "bold": 1,
        "align": "center",
        "valign": "vcenter",
        "top": 1,
        "bottom": 1,
        "left": 1,
        "right": 1
    }
)
format_bold_left_border = workbook.add_format(
    {
        "bold": 1,
        "align": "left",
        "valign": "vcenter",
        "top": 1,
        "bottom": 1,
        "left": 1,
        "right": 1
    }
)

colors =  ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
formats_color = []
for c in colors:
    formats_color.append(workbook.add_format({"fg_color": c}))


# Create first row: years
years = pd.date_range(min_date,max_date, freq='M').strftime("%Y").tolist()
months_and_years = pd.date_range(min_date,max_date, freq='M').strftime("%b-%Y").tolist() # for later use
start_idx = 3
for year in range(int(years[0]), int(years[-1]) + 1):
    number_of_months = np.sum(np.asarray(years) == str(year))
    cell_start = xlsxwriter.utility.xl_col_to_name(start_idx) + "1"
    cell_end = xlsxwriter.utility.xl_col_to_name(start_idx + number_of_months - 1) + "1"
    worksheet.merge_range(cell_start + ":" + cell_end, year, format_bold_center_border)

    start_idx += number_of_months

# Create 2nd row: some headings and months
row_header = ['Project', 'Stage', 'Milestone'] + pd.date_range(min_date,max_date, freq='M').strftime("%b").tolist()
worksheet.write_row('A2', row_header, format_bold_center_border)

# Add project/stage/milestone info
sorted_milestones_start = dict(sorted(dates_flattened_start.items(), key=lambda item: item[1]))
row_idx_start = 3
milestone_idx = 0
visited = dict((key, False) for key in sorted_milestones_start.keys())
project_names = list(data_visualization['start'].keys())

#while np.sum(list(visited.values())) != len(list(visited.values())):
keys = list(sorted_milestones_start.keys())
rows_end_project = []
for milestone_idx in range(len(sorted_milestones_start.keys())):
    key = list(sorted_milestones_start.keys())[milestone_idx]

    # If not yet visited, add all date of this project
    if not visited[key]:
        # Get all keys related ot this project
        project_name = [name for name in project_names if name in key][0]
        keys_this_project = [key for key in keys if project_name in key]
        nr_of_rows = len(keys_this_project)

        # Add project name
        project_txt = ''.join([i if ord(i) < 128 else ' ' for i in project_name]).strip()
        if nr_of_rows > 1:
            worksheet.merge_range("A" + str(row_idx_start) + ":A" + str(row_idx_start + nr_of_rows - 1), project_txt, format_bold_center_border)
        else:
            worksheet.write("A" + str(row_idx_start), project_txt, format_bold_center_border)

        # Store end of previous project
        rows_end_project.append(row_idx_start-2)

        # Get stage order
        stage_names = list(data_visualization['start'][project_name].keys())
        stage_names_ordered = []
        for key in keys_this_project:
            for stage_name in stage_names:
                if stage_name in key:
                    stage_names_ordered.append(stage_name)
        stage_names_ordered = list(dict.fromkeys(stage_names_ordered))

        # Loop through stages
        row_idx_start_stage = row_idx_start
        color_idx = 0
        for stage_name in stage_names_ordered:

            keys_this_stage = [key for key in keys_this_project if stage_name in key]
            nr_of_milestones = len(keys_this_stage)

            # Add stage name
            stage_txt = ''.join([i if ord(i) < 128 else ' ' for i in stage_name]).strip()
            if nr_of_milestones > 1:
                worksheet.merge_range("B" + str(row_idx_start_stage) + ":B" + str(row_idx_start_stage + nr_of_milestones - 1), stage_txt, format_bold_left_border)
            else:
                worksheet.write("B" + str(row_idx_start_stage), stage_txt, format_bold_left_border)

            # Loop through Milestones
            current_row = row_idx_start_stage
            for key_this_milestone in keys_this_stage:
                # Add Milestone name
                milestone_txt = ''.join([i if ord(i) < 128 else ' ' for i in str.replace(key_this_milestone, project_name + stage_name, '')]).strip()
                worksheet.write('C' + str(current_row), milestone_txt, format_bold_left_border)

                # Select date range
                month_start = dates_flattened_start[key_this_milestone].strftime("%b-%Y")
                month_end = dates_flattened_end[key_this_milestone].strftime("%b-%Y")
                start_column = np.where(np.asarray(months_and_years) == month_start)[0][0] + 3
                end_column = np.where(np.asarray(months_and_years) == month_end)[0][0] + 3

                # Format this range
                for column in range(start_column, end_column + 1):
                    worksheet.write(xlsxwriter.utility.xl_col_to_name(column) + str(current_row), "", formats_color[color_idx])

                # update row index
                current_row += 1

                # Flag as visited
                visited[key_this_milestone] = True


            # Update index
            row_idx_start_stage += nr_of_milestones
            color_idx += 1


        # Update starting row
        row_idx_start += nr_of_rows


# Loop through all cells underneath the months and adjust format; fill with space
column_start = 3
column_end = column_start + len(months_and_years)
row_start = 2
row_end = row_start + len(dates_flattened_start.keys())


possible_keys = ["bold", "align", "valign","top","bottom","left","right", "fg_color"]
for row in range(row_start,row_end):
    for col in range(column_start,column_end):
        # Get current format
        format_dict = {}
        if row in worksheet.table.keys():
            if col in worksheet.table[row].keys():
                for key, value in worksheet.table[row][col].format.__dict__.items():
                    if key in possible_keys:
                        format_dict[key] = value

        # Add top and bottom border if this cell is coloured
        if "fg_color" in format_dict.keys():
            format_dict['top'] = 1
            format_dict['bottom'] = 1

        # Add right border if end of table or end of year
        if col == column_end -1:
            format_dict['right'] = 1
        if 'Dec' in months_and_years[col - column_start]:
            format_dict['right'] = 1


        # Add bottom border if end of project or and of table
        if row == row_end - 1:
            format_dict['bottom'] = 1
        if row in rows_end_project:
            format_dict['bottom'] = 1


        # Add red borders if column is today
        if datetime.today().strftime("%b-%Y") == months_and_years[col - column_start]:
            format_dict['left'] = 5
            format_dict['right'] = 5
            format_dict['left_color'] = 'red'
            format_dict['right_color'] = 'red'
            if row == row_start:
                format_dict['top'] = 5
                format_dict['top_color'] = 'red'
            if row == row_end -1:
                format_dict['bottom'] = 5
                format_dict['bottom_color'] = 'red'

        # Update format
        format_wb = workbook.add_format(format_dict)
        worksheet.write(row, col, " ", format_wb)


# Adjust column widths
worksheet.autofit()
worksheet.set_column(2, 2, 30)

# Freeze panes
worksheet.freeze_panes(2, 3)

# Set this new worksheet is visible by default
worksheet.activate()

# Properly close excel file
workbook.close()

# Save a picture
excel2img.export_img(excel_filename,'planning_log/planning_visualization_' + datetime.today().strftime("%Y-%m-%d") + '.png')


### Add Stage/Project information if deducible
# Items with a Milestone but missing information upwards
filter_items = {
  "and": [{
    "property": "Milestone",
    "relation": {
      "is_not_empty": True
    }
  }, {"or": [
        {
        "property": "Project",
        "relation": {
          "is_empty": True
            }
        },
        {
        "property": "Stage",
        "relation": {
          "is_empty": True
            }
        }
  ]}
  ]
}
items_tasks = get_all_items(notion, database_id=cfg.ID_DB_Tasks, filter=filter_items)
milestone_items = get_all_items(notion, database_id=cfg.ID_DB_Milestones)

# Loop through tasks and add info
for task in items_tasks:
    id = task['id']
    Milestone_id = task['properties']['Milestone']['relation'][0]['id']
    Milestone_object = [m for m in milestone_items if m['id']==Milestone_id][0]

    # Add missing Stage
    if len(task['properties']['Stage']['relation']) == 0:
        Stage_id = Milestone_object['properties']['Stage']['relation'][0]['id']
        notion.pages.update(id, properties={"Stage": {'relation': [{'id': Stage_id}]}})

    # Add missing Project
    # Note: this code will fail if the Milestone has no Project. This should not happen when properly following the setup page.
    if len(task['properties']['Stage']['relation']) == 0:
        Project_id = Milestone_object['properties']['Project']['relation'][0]['id']
        notion.pages.update(id, properties={"Project": {'relation': [{'id': Project_id}]}})

# Items with Stage info but no Project
filter_items = {
  "and": [{
    "property": "Stage",
    "relation": {
      "is_not_empty": True
    }
  },
        {
        "property": "Project",
        "relation": {
          "is_empty": True
            }
        }
  ]
}
items_tasks = get_all_items(notion, database_id=cfg.ID_DB_Tasks, filter=filter_items)
stage_items = get_all_items(notion, database_id=cfg.ID_DB_Stages)

for task in items_tasks:
    id = task['id']
    Stage_id = task['properties']['Stage']['relation'][0]['id']
    Stage_object = [m for m in stage_items if m['id']==Stage_id][0]
    Project_id = Stage_object['properties']['Project']['relation'][0]['id']
    notion.pages.update(id, properties={"Project": {'relation': [{'id': Project_id}]}})
