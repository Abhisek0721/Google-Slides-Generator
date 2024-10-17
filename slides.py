# -*- coding: utf-8 -*-
import copy
import os
import json
import re
import sys
import time
import textwrap
import traceback
from datetime import datetime
import locale

from googleapiclient.errors import HttpError
import googleapi
from urllib.parse import quote_plus
import config


def retry_function(func_to_test, max_retries=5, type_label="Task"):
    count = 0
    while True:
        count += 1
        if count > max_retries:
            print(f'[WARNING]⚠️ Too many retries for {type_label}')
            raise MaxRetriesReached(f"Max retries reached for {type_label}")

        try:
            result = func_to_test()
            return result
        except Exception as e:
            if '429' in repr(e):
                print(f'[WARNING]⚠️ Too many requests - {type_label} - Retry in 60sec:')
                time.sleep(60)
            else:
                e = "Error API : " + repr(e).split('<title>')[1].split('</title>')[0] if '<title>' in repr(e) else e
                max_retries = 3
                print(f'[WARNING]⚠️ Exception in {type_label} - Retry in 20sec: {repr(e)}')
                time.sleep(20)


def search_text_in_json(json_obj, search_text):
    if isinstance(json_obj, dict):
        for value in json_obj.values():
            if search_text_in_json(value, search_text):
                return True
    elif isinstance(json_obj, list):
        for item in json_obj:
            if search_text_in_json(item, search_text):
                return True
    elif isinstance(json_obj, str) and search_text in json_obj:
        return True

    return False

def add_paragraph(config, page_id, text):
    slides_paragraphs_height = int(config.get('SLIDES', 'slides_paragraphs_height'))
    slides_paragraphs_width = int(config.get('SLIDES', 'slides_paragraphs_width'))
    slides_paragraphs_translateX = int(config.get('SLIDES', 'slides_paragraphs_translateX'))
    slides_paragraphs_translateY = int(config.get('SLIDES', 'slides_paragraphs_translateY'))
    slides_paragraphs_fontsize = int(config.get('SLIDES', 'slides_paragraphs_fontsize'))

    try:
        element_id = 'MyParagraphBox_' + str(time.time()).replace('.', '')[3:]
        height = {'magnitude': slides_paragraphs_height, 'unit': 'PT'}
        width = {'magnitude': slides_paragraphs_width, 'unit': 'PT'}
        requests = [
            {
                'createShape': {
                    'objectId': element_id,
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': {
                        'pageObjectId': page_id,
                        'size': {'height': height, 'width': width},
                        'transform': {
                            'scaleX': 1, 'scaleY': 1,
                            'translateX': slides_paragraphs_translateX,
                            'translateY': slides_paragraphs_translateY,
                            'unit': 'PT'
                        }
                    }
                }
            },
            {
                'insertText': {
                    'objectId': element_id,
                    'insertionIndex': 0,
                    'text': text.replace("**", "")
                }
            },
            {
                'updateTextStyle': {
                    'objectId': element_id,
                    'style': {
                        'fontFamily': 'Montserrat',
                        'foregroundColor': {
                            'opaqueColor': {
                                'rgbColor': {'red': 0.0, 'green': 0.0, 'blue': 0.0}
                            }
                        },
                        'fontSize': {
                            'magnitude': slides_paragraphs_fontsize,
                            'unit': 'PT'
                        }
                    },
                    'fields': 'fontFamily,foregroundColor,fontSize',
                    'textRange': {'type': 'ALL'}
                }
            },
            {
                'updateParagraphStyle': {
                    'objectId': element_id,
                    'style': {'alignment': 'JUSTIFIED'},
                    'fields': 'alignment'
                }
            }
        ]

        # Remove if more than two *** in text
        text = re.sub(r'\*{3,}', '', text)

        # Handle bold text wrapped with '**'
        bold_parts = re.finditer(r'\*\*(.*?)\*\*', text)
        for match in bold_parts:
            start_index = match.start() - 2 * text[:match.start()].count("**")
            end_index = start_index + len(match.group(1))
            if len(match.group(1)) < 90:
                requests.append({
                    'updateTextStyle': {
                        'objectId': element_id,
                        'style': {
                            'bold': True
                        },
                        'fields': 'bold',
                        'textRange': {
                            'type': 'FIXED_RANGE',
                            'startIndex': start_index,
                            'endIndex': end_index
                        }
                    }
                })

        # Handle bold text wrapped with '✮'
        try:
            pattern = r'\u272E(.*?)\u272E'
            bold_parts = re.finditer(pattern, text)
            for match in bold_parts:
                start_index = match.start() - 2 * text[:match.start()].count("**")
                end_index = start_index + len(match.group(1)) + 3
                requests.append({
                    'updateTextStyle': {
                        'objectId': element_id,
                        'style': {
                            'bold': True
                        },
                        'fields': 'bold',
                        'textRange': {
                            'type': 'FIXED_RANGE',
                            'startIndex': start_index,
                            'endIndex': end_index
                        }
                    }
                })
        except Exception as e:
            print(repr(e))

        return requests
    except HttpError as error:
        print('[ERROR]['+']🖼️🔴 HttpError with add paragraph: ' + repr(error))

def duplicate_move_slide_id(service, presentation_id, slide_object_id, index):
    # print('[INFO] duplicate_slide('+presentation_id+', '+slide_object_id+')', 0)
    # Make the request to duplicate the slide.

    requests = [{'duplicateObject': {'objectId': slide_object_id}}]
    response = retry_function(lambda: service.presentations().batchUpdate(presentationId=presentation_id, body={
                    'requests': requests}).execute(), type_label='duplicate_slide_id')

    # The new slide's object ID is in the response.
    new_slide_id = response.get('replies')[0].get('duplicateObject').get('objectId')

    #print(f"Duplicated slide with object ID: {duplicate_slide_object_id}")
    move_slide_request = {
        'updateSlidesPosition': {
            'slideObjectIds': [new_slide_id],
            'insertionIndex': index
        }
    }

    requests = [move_slide_request]
    retry_function(lambda: service.presentations().batchUpdate(presentationId=presentation_id, body={
        'requests': requests}).execute(), type_label='move_slide_id')

    return new_slide_id

### MOST IMPORTANT FUNCTION FOR THE ASSESSMENT
def split_to_blocs(text, config):
    # Clean text
    text = text.strip('"')
    text = text.replace(' - ',', ')
    text = re.sub(r'\b(\w+)\s+-\s+(\w+)\b', r'\1-\2', text)
    text = re.sub(r'\[.*?\]', '', text)
    text = text.replace('{','').replace('}','').strip()
    text = re.sub(r'\n\s*\n[\s\n]*', '\n\n', text)
    # print('text=',repr(text))
    BLOC_MAX_LINES = int(config.get('SLIDES', 'slides_bloc_max_lines'))
    LINES_MAX_CHARS = int(config.get('SLIDES', 'slides_lines_max_chars'))
    TITLE_MAX_CHARS = int(config.get('SLIDES', 'slides_title_max_chars'))
    TITLE_PARAGRAPH_MIN_SPACE = int(config.get('SLIDES', 'slides_title_paragraph_min_space'))
    PARAGRAPHS_MIN_SPACE = int(config.get('SLIDES', 'slides_paragraphs_min_space'))

    def is_title(s):
        return len(s) < TITLE_MAX_CHARS

    # Split text with "\n\n" 
    elements_list = text.split("\n\n")

    # New list main_blocs_list
    main_blocs_list = []
    current_bloc = []
    prev_element = ''
    # For each element of text (titles and paragraphs)
    for element in elements_list:
        # print('\nMain=',main_blocs_list)
        # Case "title"
        # print("is_title(element)=",is_title(element))
        if is_title(element):
            # print(' Elt Title:'+repr(element))
            # current_bloc is empty => add element
            # current_bloc has elements, check if it is not too big for a new title
            if len(current_bloc) < BLOC_MAX_LINES - TITLE_PARAGRAPH_MIN_SPACE:
                if prev_element == 'paragraph':
                    current_bloc.append('\n\n'+element)
                else:
                    current_bloc.append(element)
            else:
                # If not enough space for title+paragraph, we add it to a new bloc
                # print('  New bloc because no space for title')
                main_blocs_list.append(' '.join(current_bloc))
                current_bloc = [element+'\n\n']
            prev_element = 'title'
        # Case "paragraphe"
        else:
            # print(' Elt Paragraph!',element)
            # We will split the paragraph with X chars
            lines = textwrap.wrap(element, LINES_MAX_CHARS)
            for j, line in enumerate(lines):
                # print('  Traitement ligne:', line)
                # If bloc is less than 18 elements
                if len(current_bloc) < BLOC_MAX_LINES:
                    # If previous element is a paragraph
                    if j == 0 and prev_element == 'paragraph':
                        if len(current_bloc) < BLOC_MAX_LINES - PARAGRAPHS_MIN_SPACE:
                            # print('  j=0 Add new paragraph to current bloc')
                            current_bloc.append('\n\n'+line)
                        else:
                            # print('  New bloc because no place for new paragraph')
                            main_blocs_list.append(' '.join(current_bloc))
                            current_bloc = [line]
                    else:# if prev elt is title, just add line
                        current_bloc.append(line)
                else:
                    # If bloc is 18 elements or more, add paragraph in new slide
                    # print('  New bloc because 18 elt already in current one')
                    # If a world was cut, do not add it to next slide
                    end_prev_line = ''
                    new_line = ' ' + line

                    if '.' in line:
                        end_prev_line = line.split('.')[0] + '.'
                        new_line = '.'.join(line.split('.')[1:])
                        main_blocs_list.append(' '.join(current_bloc) + ' ' + end_prev_line)
                        current_bloc = [new_line[1:]]
                    elif ',' in line:
                        end_prev_line = line.split(',')[0] + ','
                        new_line = '.'.join(line.split(',')[1:])
                        main_blocs_list.append(' '.join(current_bloc))
                        current_bloc = [end_prev_line + ' ' + new_line[1:]]

            # print('Current=',current_bloc)
            prev_element = 'paragraph'

    # Add last current_bloc to main list
    if current_bloc:
        main_blocs_list.append(' '.join(current_bloc))
    # print('############\n############\n############\n############\n############\n')

    # Clean blocs
    for i, bloc in enumerate(main_blocs_list):
        main_blocs_list[i] = bloc.strip()

    return main_blocs_list

def move_file(drive_service, file_id, folder_id):
    file = drive_service.files().get(fileId=file_id, fields='parents').execute()
    prev_parents = ','.join(file.get('parents'))
    drive_service.files().update(fileId=file_id,
                                 addParents=folder_id,
                                 removeParents=prev_parents,
                                 fields='id, parents').execute()

def get_results_list(client, results_enabled):
    slides_results_list = []
    for result in results_enabled:
        if 'Result ' + result in client and client['Result ' + result] != '' and client['Result ' + result] != 'NULL':
            slides_results_list.append(result)
    return slides_results_list

def replace_text(slides_service, presentation_id, slide_id, search, replace):
    requests = [{
        'replaceAllText': {
            'containsText': {
                'text': search,
                'matchCase': True
            },
            'replaceText': replace,
            'pageObjectIds': [slide_id]
        }
    }]

    return requests

def refresh_slides(slides_service, presentation_id):
    presentation = slides_service.presentations().get(presentationId=presentation_id).execute()
    return presentation.get('slides')


def prompt_splitter(lst):
    # This function take the list of prompts and split them by prefix to build the slides
    prefix_dict = {}
    for i in lst:
        prefix = i[:3]  # Get the first three characters after "Result" for slides building
        if prefix in prefix_dict:
            prefix_dict[prefix].append(i)
        else:
            prefix_dict[prefix] = [i]
    return list(prefix_dict.values())

def build_starter_slides(slides_titles_list, slides_subtitles_list, slides_introductions):
    ''' starter_slide = { main_title:
                            subtitle:
                            intro:
                        }'''
    starter_slides = {}
    for i, (titles, subtitle, intro) in enumerate(zip(slides_titles_list, slides_subtitles_list, slides_introductions)):
        starter_slides[titles.split(',')[0]] = {
            'subtitle': subtitle,
            'intro': intro
        }
    return starter_slides


def build_slides_data(config, client, slides_titles_list, results_names_lists_list, slides_results_list):
    '''output is slides_data = { main_title:
                                result:{
                                    title:
                                    content:
                          }'''
    slides_data = {}
    vers = config.get('MIS', 'version')
    for titles_list_str, results_names_list in zip(slides_titles_list, results_names_lists_list):
        titles_list = titles_list_str.strip(',').split(',')
        main_title = titles_list[0]
        titles_list = titles_list[1:]  # Skip the main title for the titles_list
        # print('titles_list=',main_title, titles_list)
        if 'couple' not in vers:
            results_names_list = results_names_list[1:]
        # print('results_names_list=',results_names_list)
        for result_name, title in zip(results_names_list, titles_list):
            # print('  Add',result_name, ' - ',title)
            if client.get('Result ' + result_name, "") not in ["", "NULL"] and result_name in slides_results_list:
                if main_title not in slides_data:
                    slides_data[main_title] = {}
                if result_name not in slides_data[main_title]:
                    slides_data[main_title][result_name] = {"title": None, "content": None}

                slides_data[main_title][result_name]['title'] = title
                slides_data[main_title][result_name]['content'] = client['Result ' + result_name]
    # print(json.dumps(slides_data, indent=4))
    return slides_data


def create_new_presentation(clients, client, slides_service, drive_service, config, vers):
    cli_id = str(client['row_id'])
    cli_uid = client['UID']
    template_presentation_id = config.get('SLIDES', 'slides_template_id')
    slides_folder_to_check_id = config.get('SLIDES', 'slides_folder_to_check_id')
    slides_folder_sent_id = config.get('SLIDES', 'slides_folder_sent_id')

    first_name, last_name = 'test', 'test'

    new_presentation_title = ''
    if 'dev' in vers:
        new_presentation_title += 'DEBUG '

    date_now = datetime.now().strftime('%d-%m-%Y-%H-%M')
    client_code = f"{last_name}${first_name}${client['Email']}"
    if 'freesuperpowers' in vers:
        new_presentation_title += f"FreeSPReport tocheck {client_code}${date_now}"
    elif 'holibotscript' in vers:
        new_presentation_title += f"HRReport tocheck {client_code}${date_now}"

    # Duplicate the template with new title (with client data)
    new_presentation = retry_function(lambda: drive_service.files().copy(fileId=template_presentation_id, body={"name": new_presentation_title}).execute(), type_label='duplicate_template')

    new_presentation_id = new_presentation.get('id')

    # Move new file to specified folder ("slides to check" folder)
    move_file(drive_service, new_presentation_id, slides_folder_to_check_id)

    slides = refresh_slides(slides_service, new_presentation_id)

    return new_presentation_id, slides

def results_count(slides_data):
    results_count = 0
    for main_title, results in slides_data.items():
        results_count += len(results)
    return results_count


def slides_filler(config, slides_service, new_presentation_id, slide_id, cursor, title, text_blocs_list):
    slide_ids = [slide_id]  # Start with the original slide id
    for j, text in enumerate(text_blocs_list):
        text = text.replace('\n\n\n', '\n\n')
        text = text.replace('\n \n', '\n\n')
        # if there is text remaining to add in slide
        if len(text) > 2:
            # if more than one slide and not last
            if len(text_blocs_list) > 1 and j != len(text_blocs_list) - 1:
                # print('  DEBUG - fill the slides ? Duplicate content template to index ' + str(cursor))
                # input('DEBUG - fill the slides ? Duplicate content template to index ' + str(cursor))
                new_slide_id = duplicate_move_slide_id(slides_service, new_presentation_id, slide_ids[-1], cursor)
                slide_ids.append(new_slide_id)  # Add the new id to the list
                global_updates_list = []
                # Update sub result title
                global_updates_list = replace_text(slides_service, new_presentation_id, slide_ids[-2], '{title}', title)
                # And add paragraph of text to slide
                global_updates_list.extend(add_paragraph(config, slide_ids[-2], text))
                retry_function(lambda: slides_service.presentations().batchUpdate(presentationId=new_presentation_id, body={
                        'requests': global_updates_list}).execute(), type_label='batchUpdate_slidefiller_1')
            else:
                # print('  DEBUG - fill the last slide, index ' + str(cursor))
                # input('DEBUG - fill the only/last slide ?')
                global_updates_list = []
                global_updates_list = replace_text(slides_service, new_presentation_id, slide_ids[-1], '{title}', title)
                global_updates_list.extend(add_paragraph(config, slide_ids[-1], text))
                retry_function(lambda: slides_service.presentations().batchUpdate(presentationId=new_presentation_id, body={
                        'requests': global_updates_list}).execute(), type_label='batchUpdate_slidefiller_2')
            cursor += 1
            # Prevent error 429 requests per minute per user
            time.sleep(3)
    return cursor

def build_slides_superpowers(config, client, titles_list, results_list, service, presentation):
    prompts_list = config.get('INTEL', 'intel_prompts_list').split(',')
    prompts_lists_list = prompt_splitter(prompts_list)

    slides_data = build_slides_data(config, client, titles_list, prompts_lists_list, results_list)
    # print(json.dumps(slides_data, indent=4))
    last_title = list(slides_data.keys())[-1]
    last_result = list(slides_data[last_title].keys())[-1]

    cursor_begin = int(config.get('SLIDES', 'slides_first_title_page'))
    index_content_template = cursor_begin
    cursor = cursor_begin + 1
    slides_done = ''

    first_key = list(slides_data.keys())[0]
    superpowers_dict = slides_data[first_key]

    # For each result, add in text in slides
    for i, (result, slide) in enumerate(superpowers_dict.items()):
        # print('  [DEBUG] result='+result+' - title=' + slide['title']+ ' - cur='+ str(cursor))
        slides = refresh_slides(service, presentation)
        slide_id_content_template = slides[index_content_template - 1]['objectId']
        title = slide['title'].strip()
        text_blocs_list = split_to_blocs(slide['content'], config)

        # Duplicate content + move to cursor
        # print('  DEBUG - Duplicate content template '+str(slide_id_content_template)+' + Move to index '+str(cursor))
        slide_id = duplicate_move_slide_id(service, presentation, slide_id_content_template, cursor)

        # print('  DEBUG - fill the slides...')
        # Duplicate and/or fill title+content slide(s)
        cursor = slides_filler(config, service, presentation, slide_id, cursor,
                               title, text_blocs_list)

        slides_done = 'OK'
        # break here to debug slides

    return slides_done

def build_slides_holistic(config, client, titles_list, results_list, starter_slide, service, presentation):
    prompts_list = config.get('INTEL', 'intel_prompts_list').split(',')
    prompts_lists_list = prompt_splitter(prompts_list)

    slides_data = build_slides_data(config, client, titles_list, prompts_lists_list, results_list)
    # print(json.dumps(slides_data, indent=4))
    last_title = list(slides_data.keys())[-1]
    last_result = list(slides_data[last_title].keys())[-1]

    cursor_begin = int(config.get('SLIDES', 'slides_first_title_page'))
    index_main_title_template = cursor_begin
    index_content_template = cursor_begin + 1
    cursor = cursor_begin + 2
    slides_done = ''
    # For each main title we fetch the result
    for main_title, results in slides_data.items():
        # print('[DEBUG]['+str(client['row_id'])+'] Creating slides of '+main_title)
        first_enabled_found = False
        # print('results='+repr(results))
        # For each result, add in text in slides
        for result, slide in results.items():
            # print('  [DEBUG] result='+result+' - title=' + slide['title']+ ' - cur='+ str(cursor))
            slides = refresh_slides(service, presentation)
            slide_id_main_title_template = slides[index_main_title_template - 1]['objectId']
            slide_id_content_template = slides[index_content_template - 1]['objectId']
            title = slide['title'].strip()
            text_blocs_list = split_to_blocs(slide['content'], config)

            # If first result of bloc OR only 1 result => duplicate 2 templates, move to cursor, fill
            if not first_enabled_found or len(results_list) == 1:
                first_enabled_found = True
                subtitle = starter_slide[main_title]['subtitle']
                intro = starter_slide[main_title]['intro'].replace('{first_name}', client['First name']).replace('&&&',
                                                                                                                 '\n\n')
                intro_blocs = split_to_blocs(intro, config)
                # print(' DEBUG - first bloc of ', main_title,' => ', title, 'cur='+str(cursor)+'&'+str(cursor+1))
                slide_id = duplicate_move_slide_id(service, presentation, slide_id_main_title_template,
                                                   cursor)
                global_updates_list = []
                global_updates_list = replace_text(service, presentation, slide_id, '{title}',
                                                   main_title.strip())
                global_updates_list.extend(replace_text(service, presentation, slide_id, '{subtitle}',
                                                   subtitle.strip()))
                retry_function(lambda: service.presentations().batchUpdate(
                    presentationId=presentation, body={
                        'requests': global_updates_list}).execute(),
                               type_label='batchUpdate_replace_title_subtitle')
                cursor += 1

            # Duplicate content + move to cursor
            # print('DEBUG - Duplicate content + Move to index '+str(cursor))
            slide_id = duplicate_move_slide_id(service, presentation, slide_id_content_template, cursor)

            # print('DEBUG - fill the slides...')
            # Duplicate and/or fill title+content slide(s)
            cursor = slides_filler(config, service, presentation, slide_id, cursor,
                                   title, text_blocs_list)

            slides_done = 'OK'
            # break here to debug slides

    return slides_done

def results_to_slides(clients, config, vers):
    print('[INFO]🖼️▶️ Slides maker')
    presentation_done = False

    drive_service = googleapi.get_drive_srv()
    slides_service = googleapi.get_slides_srv()

    # Put config titles in a list of X lists with the X category
    # For each title, built the dataset of categories
    slides_titles_list = config.slides_titles.split('\n')
    slides_subtitles_list = config.getslides_subtitles.split('\n')
    slides_introductions = config.slides_introductions.replace('\n\n','\n').split('\n')
    results_enabled_str = config.intel_results_enabled
    results_enabled = results_enabled_str.split(',')

    starter_slide = build_starter_slides(slides_titles_list, slides_subtitles_list, slides_introductions)
    # print(json.dumps(starter_slide, indent=4))

    if len(results_enabled) >= 1:
        # Format must be {Result name:Result content}
        for client in clients:
            slides_results_list = get_results_list(client, results_enabled)

            results_filled = True  # Assume all results are filled
            for title in results_enabled:
                if client['Result ' + title] == '':
                    results_filled = False
                    break

            if results_filled and len(slides_results_list) > 0 and client['Slides link'].strip() == '':
                presentation_done = True
                print('[INFO][' + str(client['row_id']) + '][' + client['UID'] + ']🖼️⏳ Creating slides... ')

                try:
                    # Create a new presentation
                    new_presentation_id, slides = create_new_presentation(clients, client, slides_service, drive_service, config, vers)

                    try:
                        if 'freesuperpowers' in vers or 'fullsuperpowers' in vers:
                            slides_done = build_slides_superpowers(config, client, slides_titles_list, slides_results_list,
                                                                   slides_service, new_presentation_id)
                        else:
                            slides_done = build_slides_holistic(config, client, slides_titles_list, slides_results_list,
                                                                starter_slide, slides_service, new_presentation_id)
                        if slides_done == 'OK':
                            print('[INFO][' + str(client['row_id']) + '][' + client['UID'] + ']🖼️✅ Slidesdeck done')
                    except Exception as e:
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        line_number = exc_tb.tb_lineno
                        print('[ERROR][' + str(client['row_id']) + '][' + client['UID'] + ']🖼️🔴 Deleting presentation because of error line '+str(line_number)+': '+repr(e))
                        drive_service.files().delete(fileId=new_presentation_id).execute()
                except Exception as e:
                    exc_type, exc_obj, exc_tb = sys.exc_info()
                    line_number = exc_tb.tb_lineno
                    print('[ERROR][' + str(client['row_id']) + '][' + client['UID'] + ']🖼️🔴 Error line '+str(line_number)+' creating slides: ' + repr(e))
        if not presentation_done:
            print('[INFO]🖼️ No slidesdeck to make')
    else:
        print('[INFO]🖼️ No result enabled')
    return clients
    

"""
HERE IS THE MAIN FUNCTION TO RUN TO BE ABLE TO PRODUCE REPORTS
"clients" variable can be loaded from the txt file I sent you clients_list.txt
"config" variable can be skipped and you can replace the values by what I sent in the email
"vers" variable can be equals to "freesuperpowersdev" or "holibotscriptdev"
"""


clients = []
vers = "freesuperpowersdev"

print('script start')
results_to_slides(clients, config, vers)
print('script end')
