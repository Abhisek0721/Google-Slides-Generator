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
    slides_subtitles_list = config.slides_subtitles.split('\n')
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
                if client['Result ' + title.strip()] == '':
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


clients = [
    {
        "row_id": 13,
        "UID": "ruizjean-fran\u00e7oisendale.devteam@gmail.com",
        "Email": "theo175@gmail.com",
        "Last name": "Ruiz",
        "First name": "Jean-Fran\u00e7ois",
        "Status": "[BOT] Extraction OK => Prompts + Slides",
        "Assigned to": "Th\u00e9o",
        "Comment staff": "Test claude context",
        "Slides link": "",
        "Audio links": "",
        "Phone number": "",
        "DATE OF BIRTH": "01.01.1980",
        "TIME OF BIRTH": "04:00",
        "CITY OF BIRTH": "Paris",
        "COUNTRY OF BIRTH": "France",
        "GENDER": "Homme",
        "MIDDLE NAMES": "",
        "Date Registered": "2024-08-28 06:55:39",
        "Date pdf sent": "",
        "Result Identity & Perso": "",
        "Result Identity & Perso1": "Jean-Fran\u00e7ois, you are a vibrant force of nature, a true adventurer at heart. Your innate curiosity and thirst for new experiences drive you to explore the world with an open mind and an eager spirit. You have a magnetic personality that draws people to you, as your enthusiasm for life is truly contagious.\n\nYour unique blend of creativity and practicality allows you to navigate life's challenges with a sense of purpose and determination. You have a keen eye for beauty and a deep appreciation for the finer things in life, which is reflected in your love for art, music, and anything that brings joy to the senses.\n\nAs a natural-born leader, you have the ability to inspire and motivate others with your unwavering confidence and charisma. Your strong sense of self and your ability to remain grounded in the face of adversity make you a pillar of strength for those around you.\n\nYou are a true free spirit, unafraid to break away from the norm and forge your own path. Your independent nature and your ability to adapt to new situations with ease make you a true trailblazer. You have a knack for turning challenges into opportunities and for finding the silver lining in even the most difficult circumstances.\n\nYour life is one of constant growth and evolution, as you are always seeking new ways to expand your horizons and push yourself to new heights. Your natural curiosity and love for learning keep you engaged and motivated, no matter what life throws your way.\n\nAt your core, you are a deeply compassionate and empathetic individual, with a strong desire to make a positive impact on the world around you. Your generosity of spirit and willingness to lend a helping hand to those in need make you a true asset to your community.\n\nAs you continue on your journey through life, remember to trust your instincts and stay true to yourself. Your unique combination of talents and abilities is what sets you apart and makes you special. Embrace your individuality and let your light shine bright, for the world needs more people like you who are unafraid to be their authentic selves.",
        "Result Identity & Perso2": "Imagine the depths of the ocean, Jean-Fran\u00e7ois, where currents swirl with hidden intensity. Your emotions are like those mysterious waters, flowing beneath the surface, rarely seen but always felt. You have a profound capacity for emotional connection, yet you often keep your feelings closely guarded, revealing them only to those who have earned your trust.\n\nIn relationships, you crave deep, soulful bonds. You seek partners who can navigate the complexities of your emotional world and are unafraid to dive into the depths with you. Trust is paramount, and once given, your loyalty knows no bounds. You have a talent for understanding others' motivations and desires, sensing what lies beneath their words and actions.\n\nYour emotional intensity can be both a strength and a challenge. At times, you may feel overwhelmed by the sheer force of your feelings, struggling to find outlets for their expression. Learning to channel this energy constructively is key to your personal growth and well-being. When you find healthy ways to express your emotions, whether through art, music, or intimate conversations, you tap into a profound source of creativity and inspiration.\n\nIn your journey through life, you may find yourself drawn to experiences that challenge you emotionally, pushing you to confront your deepest fears and desires. Embrace these opportunities for growth, even when they feel uncomfortable. It is through facing your shadows that you find the light within yourself.\n\nAt your core, you possess a remarkable resilience. Like the ocean, you have weathered countless storms, emerging stronger and wiser each time. Trust in your ability to navigate life's challenges, drawing upon the depth of your emotional intelligence and intuition.\n\nAs you continue to evolve, remember to balance your intensity with moments of quietude and self-reflection. Just as the ocean has its calm waters and peaceful bays, so too must you find spaces of tranquility within yourself. In these moments of stillness, you can reconnect with your inner wisdom and find clarity amidst the tumult of emotions.\n\nUltimately, your emotional depth is a gift, Jean-Fran\u00e7ois. It allows you to experience life with a richness and vibrancy that few can imagine. Embrace your unique emotional landscape, and let it guide you toward a life filled with meaning, connection, and profound personal growth.",
        "Result Identity & Perso3": "Jean-Fran\u00e7ois, you possess a captivating presence that draws others in. Your natural charisma and charm allow you to navigate social situations with ease, effortlessly engaging with people from all walks of life. You have a gift for putting others at ease and creating a warm, welcoming atmosphere wherever you go.\n\nYour ability to read and respond to the emotional undercurrents of any situation is remarkable. You instinctively know when to step forward and when to hold back, adapting to the needs of the moment. This emotional intelligence enables you to build deep, meaningful connections with those around you.\n\nIn social settings, you radiate a magnetic energy that attracts others. People are drawn to your warmth, sincerity, and ability to make them feel seen and heard. You have a way of bringing out the best in others, encouraging them to open up and share their true selves with you.\n\nYour empathetic nature allows you to sense the unspoken needs and desires of those around you. You offer support and guidance without being overbearing, and your gentle touch can soothe even the most troubled hearts. People often seek you out for your wisdom and compassion, knowing that you will listen without judgment.\n\nYour social interactions are filled with purpose and meaning. You deeply understand the interconnectedness of all things and strive to create harmony and balance in your relationships. You are a natural mediator, able to bridge divides and bring people together in a spirit of unity and cooperation.\n\nIn your personal relationships, you are a loyal and devoted partner, friend, and family member. You have a strong sense of responsibility towards those you love, and you will go to great lengths to support and nurture them. Your ability to anticipate the needs of others and respond with compassion and understanding creates a sense of safety and security in your relationships.\n\nAs you move through the world, your presence leaves a lasting impact. Your authenticity and genuine concern for others inspire them to be their best selves, and your example encourages them to lead lives of purpose and meaning. You are a catalyst for positive change, and your influence ripples out into the world in countless ways.",
        "Result Identity & Perso4": "Here are a few books that could be particularly insightful and enriching for your personal journey:\n\n**The Power of Intention** by Dr. Wayne W. Dyer  \nThis book delves into the transformative potential of aligning your thoughts, emotions, and actions with your deepest intentions. It explores how cultivating a mindset of purposefulness and positivity can profoundly impact your life experiences and relationships.\n\n**The Art of Possibility** by Rosamund Stone Zander and Benjamin Zander  \nThrough engaging stories and practical wisdom, this book invites you to shift your perspective and embrace a mindset of abundance, creativity, and opportunity. It encourages letting go of limiting beliefs and exploring new possibilities for personal growth and fulfillment.\n\n**The Untethered Soul** by Michael A. Singer  \nThis profound book invites you on an inward journey to explore the nature of your thoughts, emotions, and inner awareness. It offers insights on releasing mental and emotional blocks, finding inner peace, and aligning with your authentic self.\n\n**The Book of Awakening** by Mark Nepo  \nOrganized as a daily guide, this book offers poetic reflections, stories, and practices to nurture mindfulness, resilience, and self-discovery. It encourages embracing both the joys and challenges of life with an open heart and finding meaning in everyday moments.\n\n**The Art of Happiness** by Dalai Lama and Howard C. Cutler  \nDrawing on Buddhist wisdom and modern psychology, this book explores the fundamental principles of cultivating happiness and inner peace. It offers practical guidance on developing compassion, managing emotions, and finding contentment in the face of life's challenges.\n\nWe'd love to hear your thoughts and experiences if you've already explored any of these titles. If they're new to you, we hope they offer valuable perspectives and inspiration as you navigate your path of personal growth. Happy reading!",
        "Result Identity & Perso5": "Jean-Fran\u00e7ois, your personality is a vibrant blend of steadfast determination and adventurous curiosity. At your core, you possess an unwavering drive to build and nurture what matters most to you. This strength is your anchor, providing stability to weather any storm.\n\nWithin this steadiness lies a restless spirit, an insatiable curiosity that propels you to explore the unknown. You seek new experiences, always eager to embrace new horizons and challenge the status quo. This duality creates a captivating dance between the familiar and the uncharted.\n\nYour emotional landscape is a realm where intensity and transformation prevail. You have a profound capacity for deep connections, yearning to merge with others on a profound level. This emotional depth is both your greatest strength and your greatest challenge as you navigate the shadows and the light within yourself and others.\n\nIn social interactions, you are a natural responder, attuned to the needs and desires of those around you. Your intuition guides you, allowing you to pick up on subtle cues and respond with grace and empathy. You have a gift for creating harmony, bringing people together, and fostering a sense of belonging.\n\nAs you navigate the world, your resilience and adaptability shine through. You turn obstacles into opportunities and find the silver lining in even the darkest clouds. This inner strength, combined with your charisma and creativity, makes you a formidable force.\n\nEmbrace your multifaceted nature, Jean-Fran\u00e7ois, and trust in the wisdom of your instincts. Your journey is one of self-discovery and growth, constantly unfolding your potential. As you explore the depths of your being, remember that your uniqueness is your greatest asset.",
        "Result Challenges": "",
        "Result Challenges1": "",
        "Result Challenges2": "",
        "Result Challenges3": "",
        "Result Challenges4": "",
        "Result Challenges5": "",
        "Result SuperPowersFull": "",
        "Result SuperPowers1": "",
        "Result SuperPowers2": "",
        "Result SuperPowers3": "",
        "Result SuperPowers4": "",
        "Result SuperPowers5": "",
        "Result SuperPowers6": "",
        "Result SuperPowers7": "",
        "Result SuperPowers8": "NULL",
        "Result SuperPowers9": "NULL",
        "Result Life Mission": "",
        "Result Life Mission1": "",
        "Result Life Mission2": "",
        "Result Life Mission3": "",
        "Result Life Mission4": "",
        "Result Life Mission5": "",
        "Result Holi Full": "",
        "Result Holi1": "",
        "Result Holi2": "",
        "Result Holi3": "",
        "Result Holi4": "",
        "Result Holi5": "",
        "Result Predictions": "",
        "Result Predictions1": "",
        "Result Predictions2": "",
        "Result Predictions3": "",
        "Result Predictions4": "",
        "Result Predictions5": "",
        "": ""
    },
    {
        "row_id": 14,
        "UID": "probstnicolaendale.devteam@gmail.com",
        "Email": "theo175@gmail.com",
        "Last name": "Probst",
        "First name": "Nicola",
        "Status": "[BOT] Extraction OK => Prompts + Slides",
        "Assigned to": "Th\u00e9o",
        "Comment staff": "Test claude context",
        "Slides link": "",
        "Audio links": "",
        "Phone number": "",
        "DATE OF BIRTH": "02.02.1980",
        "TIME OF BIRTH": "09:00",
        "CITY OF BIRTH": "Berlin",
        "COUNTRY OF BIRTH": "Germany",
        "GENDER": "Female",
        "MIDDLE NAMES": "",
        "Date Registered": "2024-08-28 06:55:39",
        "Date pdf sent": "",
        "Result Identity & Perso": "",
        "Result Identity & Perso1": "Nicola, you have an innate drive to manifest your unique vision in the world. Your energy is magnetic, drawing others to your creative projects and innovative ideas. You naturally inspire and lead, rallying people around a common goal with your enthusiasm and charisma.\n\nYour path is one of exploration and adventure, constantly seeking new experiences and perspectives. You thrive on change and variety, embracing the unexpected twists and turns of life with curiosity and an open mind. This adaptability allows you to navigate challenges with grace and resilience, finding opportunities for growth in every situation.\n\nAt your core, you are a seeker of truth and meaning. You have a deep desire to understand the mysteries of the universe and your place within it. This introspective nature may lead you to explore spirituality, philosophy, or the arts as avenues for self-discovery and expression.\n\nYour unique combination of practicality and intuition enables you to bring your dreams into reality. You can turn abstract ideas into tangible action plans, breaking down complex projects into manageable steps. This grounded approach, coupled with your creative vision, allows you to manifest your goals with efficiency and style.\n\nIn relationships, you value authenticity and depth. You seek connections that challenge you to grow and evolve while providing a safe space for vulnerability and self-expression. Your magnetic presence and genuine interest in others make you a sought-after friend and partner.\n\nYour life path is one of freedom and adventure, with a strong desire to break free from convention and forge your own way. You are a natural entrepreneur, drawn to unconventional careers or lifestyles that allow you to express your individuality and creativity. Embrace your unique talents and trust your instincts as you navigate the twists and turns of your journey.\n\nRemember, your greatest strength lies in your ability to adapt and innovate in the face of change. Stay open to new possibilities and trust that your intuition will guide you toward your highest path. As you continue to evolve and grow, you will inspire others with your authenticity, creativity, and zest for life.",
        "Result Identity & Perso2": "Your emotional landscape is a fascinating realm, Nicola, where depth and discernment intertwine. You possess an innate ability to navigate the intricacies of your feelings with a keen sense of understanding. Your emotions are not just fleeting experiences but a profound source of wisdom and insight that guides you through life's challenges and opportunities.\n\nAt the core of your emotional being lies a strong desire for harmony and balance. You have a natural inclination to create and maintain supportive, nurturing relationships with others. Your empathetic nature allows you to sense the needs and feelings of those around you, making you a trusted confidant and a pillar of strength for your loved ones.\n\nYour emotional intelligence is heightened by your ability to analyze and understand the complexities of human emotions. You have a gift for seeing beyond the surface and delving into the depths of your own feelings and those of others. This keen perception enables you to navigate emotionally charged situations with grace and compassion, offering support and guidance when needed.\n\nYour emotional world is a sanctuary where you find solace and rejuvenation. You have a deep appreciation for the transformative power of emotions and the growth that comes from embracing them fully. You understand that vulnerability is a strength, and you are not afraid to explore the shadows of your soul to emerge stronger and more self-aware.\n\nYour emotional expression is a beautiful dance between introspection and connection. You have a unique ability to articulate your feelings with clarity and authenticity, inviting others to share in your emotional journey. Your words have the power to heal, inspire, and create profound bonds with those who are fortunate enough to witness your emotional depth.\n\nAs you navigate through life, your emotional landscape will continue to evolve and expand. Embrace the ebb and flow of your feelings, knowing that they are a fundamental part of who you are. Trust in your intuition and let your emotions guide you toward a path of self-discovery, growth, and fulfillment. Your emotional essence is a beautiful tapestry, woven with threads of compassion, understanding, and resilience, making you a truly remarkable individual.",
        "Result Identity & Perso3": "Nicola, you stand out with a remarkable mix of ambition, resourcefulness, and an irresistible charm that draws people to you. Your natural charisma and adaptability in social situations enable you to navigate life with purpose and determination.\n\nYou have a strong sense of individuality and a drive to leave your mark on the world. Your leadership qualities and innovative thinking help you tackle challenges and pursue your goals with unwavering focus. You excel at recognizing opportunities and seizing them with confidence and enthusiasm.\n\nIn social settings, you often become the center of attention, captivating others with your wit, intelligence, and engaging personality. Your ability to connect with people from various backgrounds showcases your versatility and open-mindedness. You bring people together and foster camaraderie and shared purpose.\n\nBeneath your strength and resilience lies emotional depth and sensitivity. You process and express your emotions in a way that is both powerful and transformative. This emotional intelligence lets you empathize with others and provide support and guidance when needed.\n\nDriven by a desire for personal growth and self-discovery, you never settle for the status quo. You're always seeking new experiences and challenges to expand your horizons and reach your full potential. Your curiosity and thirst for knowledge propel you forward, and you're not afraid to take risks and embrace change.\n\nYour path is one of leadership and innovation, and you have the potential to make a significant impact in your chosen field. Whether pursuing a career, building relationships, or exploring new passions, you infuse everything with determination, creativity, and charm.\n\nEmbrace your individuality and trust in your abilities, Nicola. Your journey of self-discovery and growth is eagerly anticipated by the world. Stay true to yourself, and remember that your unique combination of strengths and qualities sets you apart and makes you truly extraordinary.",
        "Result Identity & Perso4": "Here are a few carefully selected books that may resonate with your journey of self-discovery and personal growth:\n\n**The Untethered Soul: The Journey Beyond Yourself by Michael A. Singer**  \nThis book invites you to explore the depths of your inner world, helping you understand the constant chatter of your thoughts and emotions. It offers insights on how to find inner peace and live in the present moment, aligning with your desire for emotional balance and self-awareness.\n\n**The Art of Possibility: Transforming Professional and Personal Life by Rosamund Stone Zander and Benjamin Zander**  \nThe authors present a framework for approaching life with a sense of possibility and openness, encouraging readers to embrace challenges as opportunities for growth. This perspective can support your adaptable nature and inspire you to navigate changes with grace and resilience.\n\n**The Seven Spiritual Laws of Success: A Practical Guide to the Fulfillment of Your Dreams by Deepak Chopra**  \nChopra presents seven principles that can guide you in aligning your life with universal laws, leading to greater fulfillment and success. This book may resonate with your desire for a purposeful and harmonious life, helping you tap into your natural gifts and create meaningful connections.\n\n**Big Magic: Creative Living Beyond Fear by Elizabeth Gilbert**  \nGilbert explores the nature of creativity and encourages readers to embrace their passions and curiosity. This book can inspire you to trust your instincts, take risks, and express yourself authentically, aligning with your adventurous spirit and desire for self-expression.\n\nWe hope these recommendations spark your curiosity and provide valuable insights for your journey. If you've already read any of these books, we'd love to hear your thoughts and how they resonated with you. And if you decide to explore these titles, we look forward to learning about your experiences and the wisdom you gain along the way.",
        "Result Identity & Perso5": "",
        "Result Challenges": "",
        "Result Challenges1": "",
        "Result Challenges2": "",
        "Result Challenges3": "",
        "Result Challenges4": "",
        "Result Challenges5": "",
        "Result SuperPowersFull": "",
        "Result SuperPowers1": "",
        "Result SuperPowers2": "",
        "Result SuperPowers3": "",
        "Result SuperPowers4": "",
        "Result SuperPowers5": "",
        "Result SuperPowers6": "",
        "Result SuperPowers7": "",
        "Result SuperPowers8": "",
        "Result SuperPowers9": "",
        "Result Life Mission": "",
        "Result Life Mission1": "",
        "Result Life Mission2": "",
        "Result Life Mission3": "",
        "Result Life Mission4": "",
        "Result Life Mission5": "",
        "Result Holi Full": "",
        "Result Holi1": "",
        "Result Holi2": "",
        "Result Holi3": "",
        "Result Holi4": "",
        "Result Holi5": "",
        "Result Predictions": "",
        "Result Predictions1": "",
        "Result Predictions2": "",
        "Result Predictions3": "",
        "Result Predictions4": "",
        "Result Predictions5": "",
        "": ""
    }
]
vers = "freesuperpowersdev"

print('script start')
results_to_slides(clients, config, vers)
print('script end')
