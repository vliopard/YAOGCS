import ctypes
import json
from datetime import datetime
from time import sleep
from time import time

from connector.calendar_event import CalendarEvent
from connector.g_calendar import GoogleCalendarConnector
from connector.ms_outlook import MicrosoftOutlookConnector
from system.routines import sync_outlook_to_google
from utils.utils import line_number
from utils.utils import object_serializer
from utils.utils import print_display


class SystemObserver:
    def __init__(self):
        self.enabled = True
        self.first_sleep = True
        self.sleep_timeout = 250
        self.continuous = 0x80000000
        self.system_required = 0x00000001
        self.display_required = 0x00000002
        self.timeout = 30 * 60

    def system_observer(self):
        if self.first_sleep:
            self.first_sleep = False
        else:
            print_display(f'{line_number()} Sleeping [{self.sleep_timeout}]...')
            sleep(self.sleep_timeout)

        if self.enabled:
            ctypes.windll.kernel32.SetThreadExecutionState(self.continuous | self.system_required | self.display_required)

    def time_out(self):
        return self.timeout

    def system_original_state(self):
        if self.enabled:
            print_display(f'{line_number()} Continuous system state...')
            ctypes.windll.kernel32.SetThreadExecutionState(self.continuous)


def import_export():
    connection_ms_outlook = MicrosoftOutlookConnector()
    connection_g_calendar = GoogleCalendarConnector()
    ms_outlook_events = connection_ms_outlook.get_ms_outlook_events()
    g_calendar_events = connection_g_calendar.get_g_calendar_events()
    type_action = 'X'
    if type_action == 'oo':
        # ---------------------------
        # 1. Import Outlook, export Outlook
        # ---------------------------
        print('\n=== Import Outlook → Export Outlook ===')
        for event_id, event_data in ms_outlook_events.items():
            ue = CalendarEvent()
            ue.import_ms_outlook(event_data)
            result = ue.export_ms_outlook()
            print(f'Outlook Event ID: {event_id}')
            print(json.dumps(object_serializer(result),
                             indent=4))
    if type_action == 'og':
        # ---------------------------
        # 2. Import Outlook, export GCal
        # ---------------------------
        print('\n=== Import Outlook → Export GCal ===')
        for event_id, event_data in ms_outlook_events.items():
            ue = CalendarEvent()
            ue.import_ms_outlook(event_data)
            result = ue.export_g_calendar()
            print(f'Outlook Event ID: {event_id}')
            print(json.dumps(object_serializer(result),
                             indent=4))
    if type_action == 'gg':
        # ---------------------------
        # 3. Import GCal, export GCal
        # ---------------------------
        print('\n=== Import GCal → Export GCal ===')
        for event_id, event_data in g_calendar_events.items():
            ue = CalendarEvent()
            ue.import_g_calendar(event_data)
            result = ue.export_g_calendar()
            print(f'GCal Event ID: {event_id}')
            print(json.dumps(object_serializer(result),
                             indent=4))
    if type_action == 'go':
        # ---------------------------
        # 4. Import GCal, export Outlook
        # ---------------------------
        print('\n=== Import GCal → Export Outlook ===')
        for event_id, event_data in g_calendar_events.items():
            ue = CalendarEvent()
            ue.import_g_calendar(event_data)
            result = ue.export_ms_outlook()
            print(f'GCal Event ID: {event_id}')
            print(json.dumps(object_serializer(result),
                             indent=4))


def main_observer(enabled=True):
    if enabled:
        print_display(f'{line_number()}')
        system_observer = SystemObserver()
        try:
            last_sync = 0
            while True:
                system_observer.system_observer()
                now = time()
                antes = datetime.fromtimestamp(now + system_observer.time_out()).strftime('%Y.%m.%d %p %I:%M:%S')
                nls = now - last_sync
                print_display(f'{line_number()} [{now}]-[{last_sync}]>=[{nls}][{system_observer.time_out()}]')
                if nls >= system_observer.time_out():
                    current_time = datetime.now().strftime('%Y.%m.%d %p %I:%M:%S')
                    print_display(f'[{current_time}] Syncing Outlook to Google...')
                    connection_ms_outlook = MicrosoftOutlookConnector()
                    connection_g_calendar = GoogleCalendarConnector()
                    sync_outlook_to_google(connection_ms_outlook,
                                           connection_g_calendar)
                    last_sync = now
                print_display(f'[{antes}] NEXT Syncing Outlook to Google...')
        except KeyboardInterrupt:
            system_observer.system_original_state()
        print('Bye...')


if __name__ == '__main__':
    main_observer(enabled=True)
