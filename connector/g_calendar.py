import os
import socket
import ssl
import time
from datetime import datetime
from functools import wraps
from pathlib import Path

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from connector.event_mapping import EventMapping
from system.tools import convert_object_to_string
from system.tools import get_master_id
from system.tools import line_number
from system.tools import print_box
from system.tools import print_display
from system.tools import time_max
from system.tools import time_min
from system.tools import trim_id

_RETRYABLE_EXCEPTIONS = (ssl.SSLError,
                         ssl.SSLEOFError,
                         ConnectionResetError,
                         ConnectionAbortedError,
                         socket.timeout,
                         TimeoutError,
                         OSError)
_RETRY_STATUS_CODES = {429,
                       500,
                       502,
                       503,
                       504}
_MAX_RETRIES = 4
_RETRY_BASE_DELAY = 2.0


def _google_api_retry(func):
    @wraps(func)
    def wrapper(*args,
                **kwargs):
        delay = _RETRY_BASE_DELAY
        for attempt in range(1,
                             _MAX_RETRIES + 1):
            try:
                return func(*args,
                            **kwargs)
            except HttpError as http_error:
                if http_error.status_code in _RETRY_STATUS_CODES and attempt < _MAX_RETRIES:
                    print_display(f'{line_number()} Google API HTTP {http_error.status_code} on attempt {attempt}/{_MAX_RETRIES}, retrying in {delay:.0f}s...')
                    time.sleep(delay)
                    delay *= 2
                else:
                    raise
            except _RETRYABLE_EXCEPTIONS as net_error:
                if attempt < _MAX_RETRIES:
                    print_display(f'{line_number()} Transient network error on attempt {attempt}/{_MAX_RETRIES}: [{net_error}], retrying in {delay:.0f}s...')
                    time.sleep(delay)
                    delay *= 2
                else:
                    print_display(f'{line_number()} All {_MAX_RETRIES} retries exhausted: [{net_error}]')
                    raise

    return wrapper


class GoogleCalendarHelper:
    def __init__(self):
        credentials_dir = (Path(__file__).resolve().parent.parent / 'resources' / 'credentials').resolve()
        self.g_calendar_read = 'r'
        self.g_calendar_write = 'w'
        self.g_calendar_id = 'primary'
        self.g_calendar_token = str(credentials_dir / 'token.json')
        self.g_calendar_credentials = str(credentials_dir / 'credentials.json')
        self.g_calendar_scopes = ['https://www.googleapis.com/auth/calendar']
        self.g_calendar_service = self.get_google_service()

        # self.g_calendar_token = f'{credentials_dir}/token.json'

        # self.g_calendar_credentials = f'{credentials_dir}/credentials.json'

    def get_google_service(self):
        g_calendar_credentials = None
        if os.path.exists(self.g_calendar_token):
            g_calendar_credentials = Credentials.from_authorized_user_file(self.g_calendar_token,
                                                                           self.g_calendar_scopes)
        if not g_calendar_credentials or not g_calendar_credentials.valid:
            if g_calendar_credentials and g_calendar_credentials.expired and g_calendar_credentials.refresh_token:
                try:
                    g_calendar_credentials.refresh(Request())
                except RefreshError:
                    g_calendar_flow = InstalledAppFlow.from_client_secrets_file(self.g_calendar_credentials,
                                                                                self.g_calendar_scopes)
                    g_calendar_credentials = g_calendar_flow.run_local_server(port=0)
            else:
                g_calendar_flow = InstalledAppFlow.from_client_secrets_file(self.g_calendar_credentials,
                                                                            self.g_calendar_scopes)
                g_calendar_credentials = g_calendar_flow.run_local_server(port=0)
            with open(self.g_calendar_token,
                      self.g_calendar_write) as g_calendar_token_local:
                g_calendar_token_local.write(g_calendar_credentials.to_json())
        return build('calendar',
                     'v3',
                     credentials=g_calendar_credentials)

    @_google_api_retry
    def g_calendar_get_all_instances(self):
        return self.g_calendar_service.events().list(calendarId=self.g_calendar_id,
                                                     timeMin=time_min(),
                                                     timeMax=time_max(),
                                                     maxResults=2500,
                                                     singleEvents=False).execute()

    @_google_api_retry
    def g_calendar_get_all_sub_instances(self):
        return self.g_calendar_service.events().list(calendarId=self.g_calendar_id,
                                                     timeMin=time_min(),
                                                     timeMax=time_max(),
                                                     maxResults=2500,
                                                     singleEvents=True).execute()

    @_google_api_retry
    def g_calendar_get_single_instance(self,
                                       g_calendar_single_instance_id):
        return self.g_calendar_service.events().get(calendarId=self.g_calendar_id,
                                                    eventId=g_calendar_single_instance_id).execute()

    @_google_api_retry
    def g_calendar_get_all_single_instances_inside_recurrence(self,
                                                              g_calendar_single_instance_id):
        return self.g_calendar_service.events().instances(calendarId=self.g_calendar_id,
                                                          eventId=g_calendar_single_instance_id,
                                                          timeMin=time_min(),
                                                          timeMax=time_max(),
                                                          showDeleted=False).execute()

    @_google_api_retry
    def g_calendar_get_instance_by_ical_uid(self,
                                            g_calendar_ical_uid):
        result = self.g_calendar_service.events().list(calendarId=self.g_calendar_id,
                                                       iCalUID=g_calendar_ical_uid,
                                                       maxResults=1,
                                                       singleEvents=False).execute()
        items = result.get('items',
                           [])
        return items[0] if items else None

    @_google_api_retry
    def g_calendar_get_instance_by_summary_and_start(self,
                                                     g_calendar_summary,
                                                     g_calendar_start_date):
        if not isinstance(g_calendar_start_date,
                          datetime):
            g_calendar_start_date = datetime(g_calendar_start_date.year,
                                             g_calendar_start_date.month,
                                             g_calendar_start_date.day,
                                             g_calendar_start_date.hour,
                                             g_calendar_start_date.minute,
                                             g_calendar_start_date.second)
        time_min = g_calendar_start_date.strftime('%Y-%m-%dT00:00:00Z')
        time_max = g_calendar_start_date.strftime('%Y-%m-%dT23:59:59Z')
        result = self.g_calendar_service.events().list(calendarId=self.g_calendar_id,
                                                       timeMin=time_min,
                                                       timeMax=time_max,
                                                       maxResults=2500,
                                                       singleEvents=True).execute()
        for item in result.get('items',
                               []):
            if item.get('summary',
                        '').strip().lower() == g_calendar_summary.strip().lower():
                return item
        return None

    @_google_api_retry
    def insert_instance_g_calendar(self,
                                   g_calendar_instance_body):
        return self.g_calendar_service.events().insert(calendarId=self.g_calendar_id,
                                                       body=convert_object_to_string(g_calendar_instance_body)).execute()

    @_google_api_retry
    def update_instance_g_calendar(self,
                                   g_calendar_instance_id,
                                   g_calendar_instance_body):
        return self.g_calendar_service.events().update(calendarId=self.g_calendar_id,
                                                       eventId=g_calendar_instance_id,
                                                       body=convert_object_to_string(g_calendar_instance_body)).execute()

    @_google_api_retry
    def delete_instance_g_calendar(self,
                                   g_calendar_instance_id):
        result = 'Failed'
        existed = False
        try:
            result = self.g_calendar_service.events().get(calendarId=self.g_calendar_id,
                                                          eventId=g_calendar_instance_id).execute()
            existed = True
        except HttpError as http_error:
            print_display(f'{line_number()} [Google Calendar] GET DELETE ERROR: [{http_error.status_code} | {http_error.error_details}]')
        try:
            if existed:
                result = self.g_calendar_service.events().delete(calendarId=self.g_calendar_id,
                                                                 eventId=g_calendar_instance_id).execute()
        except HttpError as http_error:
            print_display(f'{line_number()} [Google Calendar] DELETE ERROR: [{http_error.status_code} | {http_error.error_details}]')
        return result


class GoogleCalendarConnector:
    def __init__(self,
                 event_mapping: EventMapping = None):
        self.event_mapping = event_mapping if event_mapping else EventMapping()
        self.g_calendar_service = GoogleCalendarHelper()
        self.g_calendar_events = None
        self.g_calendar_event_end_dates = None

    def get_all_instances_g_calendar(self):
        g_calendar_all_instances = self.g_calendar_service.g_calendar_get_all_instances()
        g_calendar_all_instances_items = g_calendar_all_instances.get('items',
                                                                      [])
        g_calendar_all_events = dict()
        g_calendar_instance_end_dates = dict()
        for g_calendar_single_item in g_calendar_all_instances_items:
            g_calendar_instance_id = g_calendar_single_item['id']
            g_calendar_all_events[g_calendar_instance_id] = g_calendar_single_item
            if 'recurrence' in g_calendar_single_item:
                for g_calendar_rule in g_calendar_single_item['recurrence']:
                    if 'UNTIL=' in g_calendar_rule:
                        g_calendar_rule_match = g_calendar_rule.split('UNTIL=')[1].split(';')[0].split('T')[0]
                        g_calendar_instance_end_dates[g_calendar_instance_id] = g_calendar_rule_match
                g_calendar_instance_list = self.g_calendar_service.g_calendar_get_all_single_instances_inside_recurrence(g_calendar_instance_id)
                for g_calendar_instance_list_item in g_calendar_instance_list.get('items',
                                                                                  []):
                    g_calendar_all_events[g_calendar_instance_list_item['id']] = g_calendar_instance_list_item
        self.g_calendar_events = g_calendar_all_events
        self.g_calendar_event_end_dates = g_calendar_instance_end_dates
        return self.g_calendar_events

    def get_all_sub_instances_g_calendar(self):
        g_calendar_all_instances = self.g_calendar_service.g_calendar_get_all_sub_instances()
        g_calendar_all_instances_items = g_calendar_all_instances.get('items',
                                                                      [])
        g_calendar_all_events = dict()
        g_calendar_instance_end_dates = dict()
        for g_calendar_single_item in g_calendar_all_instances_items:
            print_box(f'{line_number()} [{g_calendar_single_item}]')
            g_calendar_instance_id = g_calendar_single_item['id']
            g_calendar_all_events[g_calendar_instance_id] = g_calendar_single_item
        self.g_calendar_events = g_calendar_all_events
        self.g_calendar_event_end_dates = g_calendar_instance_end_dates
        return self.g_calendar_events

    def get_single_instance_g_calendar(self,
                                       single_instance_id):
        return self.g_calendar_service.g_calendar_get_single_instance(single_instance_id)

    def get_all_single_instances_inside_recurrence_g_calendar(self,
                                                              single_instance_id):
        return self.g_calendar_service.g_calendar_get_all_single_instances_inside_recurrence(single_instance_id)

    def g_calendar_insert_instance(self,
                                   g_calendar_instance_body):
        def check_item(current_item,
                       event_mapping: EventMapping):
            if current_item:
                pair1 = event_mapping.get_recurrent_pair(get_master_id(current_item))
                print_display(f'{line_number()} [Google Calendar] 04) LOOKING [{trim_id(current_item)}] = [{trim_id(pair1[1])}]')
                if pair1:
                    return 1
                pair2 = event_mapping.get_instance_pair(get_master_id(current_item))
                print_display(f'{line_number()} [Google Calendar] 05) LOOKING [{trim_id(current_item)}] = [{trim_id(pair2[1])}]')
                if pair2:
                    return 2
            print_display(f'{line_number()} [Google Calendar] NOT FOUND')
            return 0

        insert_result = None
        try:
            print_display(f'{line_number()} [Google Calendar] INSERT <<==')
            insert_result = self.g_calendar_service.insert_instance_g_calendar(convert_object_to_string(g_calendar_instance_body))
        except HttpError as http_error:
            if http_error.status_code == 409:
                g_calendar_instance_id = g_calendar_instance_body['iCalUID']
                g_calendar_instance_id_trim = trim_id(g_calendar_instance_id)
                g_calendar_summary = g_calendar_instance_body['summary']
                print_display(f'{line_number()} [Google Calendar] INSERT RESULT ERROR: [The/requested/identifier [{g_calendar_instance_id_trim}] [{g_calendar_summary}] already/exists.]')
                g_calendar_summary = g_calendar_instance_body['summary']
                g_calendar_date = g_calendar_instance_body['start']['dateTime']

                item0 = self.get_instance_by_summary_and_start_g_calendar(g_calendar_summary,
                                                                          g_calendar_date)['id']
                print_display(f'{line_number()} [Google Calendar] 01) LOOKING SUBJECT/DATE: [{g_calendar_summary}|{type(g_calendar_summary)}][{g_calendar_date}|{type(g_calendar_date)}] = [{trim_id(item0)}]')
                if check_item(item0,
                              self.event_mapping) == 1:
                    print_display(f'{line_number()} [Google Calendar] 01) LOOKING SUBJECT/DATE: FOUND')
                else:
                    item1 = self.get_instance_by_ical_uid_g_calendar(g_calendar_instance_id)['id']
                    print_display(f'{line_number()} [Google Calendar] 02) LOOKING iCalUID: [{trim_id(g_calendar_instance_id)}] = [{trim_id(item1)}]')
                    if check_item(item1,
                                  self.event_mapping) == 2:
                        print_display(f'{line_number()} [Google Calendar] 02) LOOKING iCalUID: FOUND')
                    else:
                        print_display(f'{line_number()} [Google Calendar] 01) INSERT RESULT ERROR: [{http_error.status_code} | {http_error.error_details}]')
            else:
                print_display(f'{line_number()} [Google Calendar] 02) INSERT RESULT ERROR: [{http_error.status_code} | {http_error.error_details}]')
        return insert_result

    def get_instance_by_ical_uid_g_calendar(self,
                                            g_calendar_ical_uid):
        return self.g_calendar_service.g_calendar_get_instance_by_ical_uid(g_calendar_ical_uid)

    def g_calendar_update_instance(self,
                                   g_calendar_instance_id,
                                   g_calendar_instance_body):
        print_display(f'{line_number()} [Google Calendar] UPDATE <<==')
        return self.g_calendar_service.update_instance_g_calendar(g_calendar_instance_id,
                                                                  convert_object_to_string(g_calendar_instance_body))

    def g_calendar_delete_instance(self,
                                   g_calendar_instance_id):
        print_display(f'{line_number()} [Google Calendar] DELETE [{trim_id(g_calendar_instance_id)}]')
        return self.g_calendar_service.delete_instance_g_calendar(g_calendar_instance_id)

    def get_instance_by_summary_and_start_g_calendar(self,
                                                     g_calendar_summary,
                                                     g_calendar_start_date):
        return self.g_calendar_service.g_calendar_get_instance_by_summary_and_start(g_calendar_summary,
                                                                                    g_calendar_start_date)
