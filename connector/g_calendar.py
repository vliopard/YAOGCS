import os
import ssl
import socket
import time
from functools import wraps
from pathlib import Path

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from utils.handling_time import time_now
from utils.handling_time import time_max
from utils.utils import line_number
from utils.utils import object_serializer
from utils.utils import print_display

_RETRYABLE_EXCEPTIONS = (
    ssl.SSLError,
    ssl.SSLEOFError,
    ConnectionResetError,
    ConnectionAbortedError,
    socket.timeout,
    TimeoutError,
    OSError,
)
_RETRY_STATUS_CODES = {429, 500, 502, 503, 504}
_MAX_RETRIES = 4
_RETRY_BASE_DELAY = 2.0


def _google_api_retry(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        delay = _RETRY_BASE_DELAY
        for attempt in range(1, _MAX_RETRIES + 1):
            try:
                return func(*args, **kwargs)
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
        # credentials_dir = '../resources/credentials'
        base_dir = Path(__file__).resolve().parent.parent
        credentials_dir = (base_dir / 'resources' / 'credentials').resolve()
        self.g_calendar_read = 'r'
        self.g_calendar_write = 'w'
        self.g_calendar_id = 'primary'
        # self.g_calendar_token = f'{credentials_dir}/token.json'
        # self.g_calendar_credentials = f'{credentials_dir}/credentials.json'
        self.g_calendar_token = str(credentials_dir / 'token.json')
        self.g_calendar_credentials = str(credentials_dir / 'credentials.json')
        self.g_calendar_scopes = ['https://www.googleapis.com/auth/calendar']
        self.g_calendar_service = self.get_google_service()

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
    def get_g_calendar_events_list(self):
        return self.g_calendar_service.events().list(calendarId=self.g_calendar_id,
                                                     timeMin=time_now(),
                                                     timeMax=time_max(),
                                                     maxResults=2500,
                                                     singleEvents=False).execute()

    @_google_api_retry
    def get_g_calendar_events_instances(self,
                                        g_calendar_event_id):
        return self.g_calendar_service.events().instances(calendarId=self.g_calendar_id,
                                                          eventId=g_calendar_event_id,
                                                          timeMin=time_now(),
                                                          timeMax=time_max(),
                                                          showDeleted=True).execute()

    @_google_api_retry
    def insert_g_calendar_event(self,
                                event_body):
        return self.g_calendar_service.events().insert(calendarId=self.g_calendar_id,
                                                       body=object_serializer(event_body)).execute()

    @_google_api_retry
    def update_g_calendar_event(self,
                                google_id,
                                event_body):
        return self.g_calendar_service.events().update(calendarId=self.g_calendar_id,
                                                       eventId=google_id,
                                                       body=object_serializer(event_body)).execute()

    @_google_api_retry
    def delete_g_calendar_event(self,
                                g_calendar_instance):
        try:
            return self.g_calendar_service.events().delete(calendarId=self.g_calendar_id,
                                                           eventId=g_calendar_instance).execute()
        except HttpError as http_error:
            print_display(f'{line_number()} DELETE ERROR: [{http_error}]')
            return 'Failed'

    @_google_api_retry
    def get_g_calendar_event_instance(self,
                                      g_calendar_instance_id):
        return self.g_calendar_service.events().get(calendarId=self.g_calendar_id,
                                                    eventId=g_calendar_instance_id).execute()


class GoogleCalendarConnector:
    def __init__(self):
        self.g_calendar_service = GoogleCalendarHelper()
        self.g_calendar_events = None
        self.g_calendar_cancelled = None
        self.g_calendar_event_end_dates = None

    def get_g_calendar_events(self):
        events_result = self.g_calendar_service.get_g_calendar_events_list()
        g_calendar_events_list = events_result.get('items',
                                                   [])
        local_g_calendar_events = dict()
        local_g_calendar_cancelled = dict()
        local_g_calendar_event_end_dates = dict()
        for g_calendar_single_event in g_calendar_events_list:
            g_calendar_event_id = g_calendar_single_event['id']
            local_g_calendar_events[g_calendar_event_id] = g_calendar_single_event
            if 'recurrence' in g_calendar_single_event:
                for g_calendar_rule in g_calendar_single_event['recurrence']:
                    if 'UNTIL=' in g_calendar_rule:
                        until_match = g_calendar_rule.split('UNTIL=')[1].split(';')[0].split('T')[0]
                        local_g_calendar_event_end_dates[g_calendar_event_id] = until_match
                g_calendar_instance_list = self.g_calendar_service.get_g_calendar_events_instances(g_calendar_event_id)
                for g_calendar_instance_element in g_calendar_instance_list.get('items',
                                                                                []):
                    if g_calendar_instance_element.get('status') == 'cancelled':
                        original_start = g_calendar_instance_element.get('originalStartTime',
                                                                         {}).get('dateTime',
                                                                                 '')
                        if original_start:
                            if g_calendar_event_id not in local_g_calendar_cancelled:
                                local_g_calendar_cancelled[g_calendar_event_id] = []
                            local_g_calendar_cancelled[g_calendar_event_id].append(original_start[-10:])
                    else:
                        local_g_calendar_events[g_calendar_instance_element['id']] = g_calendar_instance_element
        self.g_calendar_events = local_g_calendar_events
        self.g_calendar_cancelled = local_g_calendar_cancelled
        self.g_calendar_event_end_dates = local_g_calendar_event_end_dates
        return self.g_calendar_events

    def g_calendar_instance(self,
                            instance_id):
        return self.g_calendar_service.get_g_calendar_event_instance(instance_id)

    def g_calendar_instances(self,
                             instance_id):
        return self.g_calendar_service.get_g_calendar_events_instances(instance_id)

    def g_calendar_insert(self,
                          event_body):
        insert_result = None
        try:
            print_display(f'{line_number()} INSERT <<==')
            insert_result = self.g_calendar_service.insert_g_calendar_event(object_serializer(event_body))
        except HttpError as http_error:
            if http_error.status_code == 409:
                cuid = event_body['iCalUID'][-20:]
                smm = event_body['summary']
                print_display(f'{line_number()} G CALENDAR INSERT RESULT ERROR: [The requested identifier <{cuid}> <{smm}> already exists.]')
            else:
                print_display(f'{line_number()} G CALENDAR INSERT RESULT ERROR: [{http_error}]')
        return insert_result

    def g_calendar_update(self,
                          google_id,
                          event_body):
        print_display(f'{line_number()} UPDATE <<==')
        return self.g_calendar_service.update_g_calendar_event(google_id,
                                                               object_serializer(event_body))

    def g_calendar_delete(self,
                          instance):
        print_display(f'{line_number()} DELETE [{instance}]')
        return self.g_calendar_service.delete_g_calendar_event(instance)
