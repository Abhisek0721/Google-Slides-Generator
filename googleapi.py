# -*- coding: utf-8 -*-
import contextlib
import glob
import json
import os
import sys
import time
from collections import defaultdict

import google.auth.transport.requests

from dateutil import parser
import gspread
from google.auth.exceptions import RefreshError

from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from utils import log, get_version


get_creds_cache = {}


# Function to load service account credentials from a JSON file
def get_cred():
    # path = '../creds/service-creds.json'
    creds_dir = os.path.join(os.getcwd(), 'creds')
    creds_path = os.path.join(creds_dir, 'service-creds.json')
    with open(creds_path, 'r') as f:
        dic = json.load(f)
    return dic

def build_delegated_creds(scopes):
    #email_service_delegation = ''
    service_account_key = get_cred()
    creds = service_account.Credentials.from_service_account_info(
        service_account_key,
        scopes=scopes
    )#.with_subject(email_service_delegation)

    return creds



def get_drive_srv():
    creds = build_delegated_creds(['https://www.googleapis.com/auth/drive'])
    # Refresh the credentials to obtain an access token
    request = google.auth.transport.requests.Request()
    creds.refresh(request)

    service = build('drive', 'v3', credentials=creds, cache_discovery=False)

    return service


def get_slides_srv():
    creds = build_delegated_creds(['https://www.googleapis.com/auth/presentations'])

    # Refresh the credentials to obtain an access token
    request = google.auth.transport.requests.Request()
    creds.refresh(request)

    service = build('slides', 'v1', credentials=creds, cache_discovery=False)

    return service

