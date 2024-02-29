import config as cfg
from notion_client import Client
import json
from datetime import datetime, timedelta
from tqdm import tqdm
from funcs import get_all_items, empty_notes
import re
import numpy as np

### Define function
def this_week(d1):
    d2 = datetime.today()
    return d1.isocalendar()[1] == d2.isocalendar()[1] and d1.year == d2.year

### setup connections
notion = Client(auth=cfg.authenticator)

### Get currently repeating tasks
filter = {'and': [{'property': 'Repeats', 'select': { 'equals': 'Yes'}},
                  {'property':'Archive', 'checkbox':{'equals': False}},
                  {'property': 'Goal status', 'select': {'equals': 'No target'}}
                  ]}
items_tasks = get_all_items(notion, database_id=cfg.ID_DB_Tasks, filter=filter)

### Remove all existing repeating tasks
print("Removing all repeating tasks")
for item_task in tqdm(items_tasks):
    notes = notion.blocks.children.list(item_task['id'])
    if empty_notes(notes):
        notion.pages.update(item_task['id'], archived=True)

### Get remaining repeating tasks
filter = {'and': [{'property': 'Repeats', 'select': { 'equals': 'Yes'}},
                  {'property':'Archive', 'checkbox':{'equals': False}},
                  ]}
remaining_tasks = get_all_items(notion, database_id=cfg.ID_DB_Tasks, filter=filter)
remaining_tasks_info = []
for task in remaining_tasks:
    name = task['properties']['Name']['title'][0]['text']['content']
    date = task['properties']['Date planned']['date']['start']
    remaining_tasks_info.append((name,date))

### Read current entries and add one task per date
with open("repeating_tasks.json","r") as file:
    jsonData = json.load(file)

date_format = '%d-%m-%Y'
for json_task in jsonData:
    # Get info
    dates = []
    date_start = datetime.strptime(json_task['Start date'], date_format)
    date_end = datetime.strptime(json_task['End date'], date_format)
    days_between_repeats = json_task['Days between repeats']

    # Get all dates
    current_date = date_start
    while current_date <= date_end:
        dates.append(current_date)
        current_date += timedelta(days=days_between_repeats)

    # Only add tasks that are today or in the future
    dates = [date for date in dates if date.date() >= datetime.today().date()]

    print("Creating " + json_task['Name'] + ", " + str(len(dates)) + " times.")
    for date in tqdm(dates):

        # Set target if the task is this week
        targeted = 'No target'
        if this_week(date):
            targeted = 'Target'

        # Get info
        date_str = date.strftime('%Y-%m-%d')
        name = '🔁 ' + json_task['Name']

        # Only add task if it did not remain
        if (name, date_str) not in remaining_tasks_info:
            # Define task
            new_item = {
                'Goal status': {'select': { 'name': targeted}},
                'Name': {'title': [{'text': {'content': name, 'link': None}, 'plain_text':name}]},
                'Repeats': {'select': { 'name': 'Yes'}},
                'Date planned': {'date': {'start': date_str, 'end': None, 'time_zone': None}}
            }

            # Add task
            notion.pages.create(parent={"database_id": cfg.ID_DB_Tasks}, properties=new_item)