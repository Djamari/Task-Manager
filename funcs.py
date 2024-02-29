import config as cfg
from datetime import date
import xlsxwriter.worksheet

def get_all_items(notion, database_id, filter=None):
    if filter is None:
        items_query = notion.databases.query(database_id=database_id)
    else:
        items_query = notion.databases.query(database_id=database_id, filter=filter)
    items_all = items_query['results']

    while items_query['has_more']:
        if filter is None:
            items_query = notion.databases.query(database_id=database_id, start_cursor=items_query['next_cursor'])
        else:
            items_query = notion.databases.query(database_id=database_id, start_cursor=items_query['next_cursor'], filter=filter)
        items_all.extend(items_query[ 'results'])

    return items_all


def create_image_message(img_dir_local, filename):
    path = 'TaskManager/' + img_dir_local + filename
    message = "[Image locally stored at: " + path + "]"
    text_item = {'object': 'block',
            'type': 'paragraph',
            'paragraph': {'rich_text': [{'type': 'text', 'text': {'content': message, 'link': None},
                                         'annotations': {'bold': True, 'italic': False, 'strikethrough': False, 'underline': False, 'code': False, 'color': 'default'},
                                         'plain_text': message, 'href': None}], 'color': 'brown_background'}}
    return text_item

def empty_notes(notes):
    if len(notes['results']) == 0:
        return True
    if len(notes['results']) == 1:
        if 'paragraph' in notes['results'][0].keys() and 'rich_text' in notes['results'][0]['paragraph'].keys():
            if len(notes['results'][0]['paragraph']['rich_text']) == 0:
                return True
    return False

def find_message_block(homepage):
    for block in homepage['results']:
        if 'paragraph' in block.keys():
            if len(block['paragraph']['rich_text']) > 0:
                if block['paragraph']['rich_text'][0]['text']['content'].startswith('Last cleanup'):
                    return block


