import config as cfg
from notion_client import Client
from tqdm import tqdm

# setup connections
notion = Client(auth=cfg.authenticator)
database_projects = notion.databases.query(database_id=cfg.ID_DB_Projects)
database_milestone = notion.databases.query(database_id=cfg.ID_DB_Milestones)
database_tasks = notion.databases.query(database_id=cfg.ID_DB_Tasks)

# Get project names
project_names = []
for item in database_projects['results']:
    project_names.append(item['properties']['Name']['title'][0]['text']['content'])


# For each milestone, add a dummy task
for item_milestone in tqdm(database_milestone['results']):
    id_milestone = item_milestone['id']
    id_stage = item_milestone['properties']['Stage']['relation'][0]['id']
    id_project = item_milestone['properties']['Project']['relation'][0]['id']

    new_item = {
        'Milestone': {'id': '%3Cjtf', 'type': 'relation', 'relation': [{'id': id_milestone}]},
        'Stage': {'relation': [{'id': id_stage}]},
        'Project': {'relation': [{'id': id_project}]},
        'Goal status': {'select': { 'name': 'No target'}},
        'Name': {'title': [{'text': {'content': 'Dummy Task', 'link': None}, 'plain_text': 'Dummy Task'}]}
    }
    notion.pages.create(parent={"database_id": cfg.ID_DB_Tasks}, properties=new_item)