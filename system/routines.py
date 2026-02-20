import re

from connector.calendar_event import CalendarEvent
from connector.event_mapping import EventMapping
from utils.handling_time import extract_date_full
from utils.utils import line_number
from utils.utils import com_object_to_dictionary
from utils.utils import print_display
from utils.utils import sort_json_list
from utils.utils import PauseToken


def copy_ms_outlook_single_event_to_g_calendar(local_ms_outlook_connection,
                                               local_g_calendar_connection,
                                               event_mapping,
                                               pause_token: PauseToken):
    print_display(f'{line_number()}')
    ms_outlook_events = local_ms_outlook_connection.get_ms_outlook_events(pause_token)
    for ms_outlook_current_id, ms_outlook_current_event in ms_outlook_events.items():
        pause_token.check()
        if not ms_outlook_current_event.get('IsRecurring',
                                            False):
            single_pair = event_mapping.get_single_event_pair(ms_outlook_current_id)
            if not single_pair:
                calendar_event = CalendarEvent()
                calendar_event.import_ms_outlook(ms_outlook_current_event)
                g_calendar_exported_event = calendar_event.export_g_calendar()
                print_display(f'{line_number()} INSERTING EVENT: [{g_calendar_exported_event}]')
                g_calendar_inserted_appointment = local_g_calendar_connection.g_calendar_insert(g_calendar_exported_event)
                if g_calendar_inserted_appointment:
                    g_calendar_master_id = g_calendar_inserted_appointment.get('id')
                    print_display(f'{line_number()} ADDING EVENT: [{ms_outlook_current_id[-10:]}] -> [{g_calendar_master_id[-10:]}]')
                    event_mapping.add_single_event(ms_outlook_current_id,
                                                   g_calendar_master_id)


def copy_g_calendar_single_event_to_ms_outlook(local_g_calendar_connection,
                                               local_ms_outlook_connection,
                                               event_mapping,
                                               pause_token: PauseToken):
    print_display(f'{line_number()}')
    g_calendar_all_events = local_g_calendar_connection.get_g_calendar_events()
    for g_calendar_event_id, g_calendar_event_item in g_calendar_all_events.items():
        pause_token.check()
        recurrence_one = 'recurrence' in g_calendar_event_item
        recurrence_two = 'recurringEventId' in g_calendar_event_item
        if not recurrence_one and not recurrence_two:
            single_pair = event_mapping.get_single_event_pair(g_calendar_event_id)
            if not single_pair:
                calendar_event = CalendarEvent()
                calendar_event.import_g_calendar(g_calendar_event_item)
                ms_outlook_exported_event = calendar_event.export_ms_outlook()
                print_display(f'{line_number()} INSERTING EVENT: [{ms_outlook_exported_event}]')
                ms_outlook_inserted_appointment = local_ms_outlook_connection.ms_outlook_insert(ms_outlook_exported_event)
                if ms_outlook_inserted_appointment:
                    ms_outlook_event_id = ms_outlook_inserted_appointment.EntryID
                    event_mapping.add_single_event(ms_outlook_event_id,
                                                   g_calendar_event_id)


def copy_g_calendar_recurrent_event_to_ms_outlook(local_g_calendar_connection,
                                                  local_ms_outlook_connection,
                                                  event_mapping,
                                                  pause_token: PauseToken):
    print_display(f'{line_number()} START')
    g_calendar_events_local = local_g_calendar_connection.g_calendar_events
    for g_calendar_event_id, g_calendar_event_data in g_calendar_events_local.items():
        pause_token.check()
        if 'recurrence' in g_calendar_event_data:
            master_pair = event_mapping.get_recurrent_master_pair(g_calendar_event_id)
            if not master_pair:
                calendar_event = CalendarEvent()
                calendar_event.import_g_calendar(g_calendar_event_data)
                ms_outlook_exported_event = calendar_event.export_ms_outlook()
                ms_outlook_inserted_appointment = local_ms_outlook_connection.ms_outlook_insert(ms_outlook_exported_event)
                if ms_outlook_inserted_appointment:
                    ms_outlook_entry_id = ms_outlook_inserted_appointment.EntryID
                    ms_outlook_master_id = ms_outlook_inserted_appointment.GlobalAppointmentID
                    print_display(f'{line_number()} ADDING RECURRENCE MASTER: [{g_calendar_event_id[-10:]}] -> [{ms_outlook_master_id[-10:]}]')
                    event_mapping.add_recurrent_master(ms_outlook_master_id,
                                                       g_calendar_event_id)
                    g_calendar_instances = local_g_calendar_connection.g_calendar_instances(g_calendar_event_id).get('items',
                                                                                                                     [])
                    ms_outlook_instances = local_ms_outlook_connection.get_recurrence_instances(ms_outlook_entry_id)
                    local_ms_outlook_connection.set_recurrence_id(ms_outlook_master_id,
                                                                  g_calendar_event_id)
                    for g_calendar_instance, ms_outlook_instance in zip(sort_json_list(g_calendar_instances,
                                                                                       'start.dateTime'),
                                                                        ms_outlook_instances):
                        gci = g_calendar_instance['id'][-10:]
                        moi = ms_outlook_instance['EntryID'][-10:]
                        print_display(f'{line_number()} ADDING RECURRENCE INSTANCE: [{ms_outlook_master_id[-10:]}] GCal[{gci}] <-> Outlook[{moi}]')
                        ms_outlook_start = str(ms_outlook_instance['StartUTC']).replace(':',
                                                                                        '').replace(' ',
                                                                                                    '').replace('+',
                                                                                                                '').replace('-',
                                                                                                                            '')
                        ms_outlook_end = str(ms_outlook_instance['EndUTC']).replace(':',
                                                                                    '').replace(' ',
                                                                                                '').replace('+',
                                                                                                            '').replace('-',
                                                                                                                        '')
                        eid = ms_outlook_instance['EntryID']
                        event_mapping.add_recurrent_instance(ms_outlook_master_id,
                                                             f'{eid}{ms_outlook_start}{ms_outlook_end}',
                                                             g_calendar_instance['id'])


def get_master_id(text: str) -> str:
    return re.sub(r'_\d{8}T\d{6}Z$',
                  '',
                  text)


def copy_ms_outlook_recurrent_event_to_g_calendar(local_ms_outlook_connection,
                                                  local_g_calendar_connection,
                                                  event_mapping,
                                                  pause_token: PauseToken):
    print_display(f'{line_number()}')
    ms_outlook_events = local_ms_outlook_connection.get_ms_outlook_events(pause_token)
    for ms_outlook_current_id, ms_outlook_current_event in ms_outlook_events.items():
        pause_token.check()
        if ms_outlook_current_event.get('IsRecurring',
                                        True):
            master_pair = event_mapping.get_recurrent_master_pair(ms_outlook_current_id)
            if not master_pair:
                calendar_event = CalendarEvent()
                calendar_event.import_ms_outlook(ms_outlook_current_event)
                g_calendar_exported_event = calendar_event.export_g_calendar()
                g_calendar_inserted_appointment = local_g_calendar_connection.g_calendar_insert(g_calendar_exported_event)
                if g_calendar_inserted_appointment:
                    g_calendar_master_id = g_calendar_inserted_appointment.get('id')
                    g_calendar_master_idr = get_master_id(g_calendar_master_id)
                    print_display(f'{line_number()} ADDING RECURRENCE MASTER: [{ms_outlook_current_id[-10:]}] -> [{g_calendar_master_id[-10:]}]')
                    event_mapping.add_recurrent_master(ms_outlook_current_id,
                                                       g_calendar_master_id)
                    ms_outlook_instances = local_ms_outlook_connection.get_recurrence_instances(ms_outlook_current_id)
                    g_calendar_instances = local_g_calendar_connection.g_calendar_instances(g_calendar_master_id).get('items',
                                                                                                                      [])
                    local_ms_outlook_connection.set_recurrence_id(ms_outlook_current_id,
                                                                  g_calendar_master_idr)
                    for ms_outlook_instance, g_calendar_instance in zip(ms_outlook_instances,
                                                                        sort_json_list(g_calendar_instances,
                                                                                       'start.dateTime')):
                        eid = ms_outlook_instance['EntryID'][-10:]
                        gci = g_calendar_instance['id'][-10:]
                        print_display(f'{line_number()} ADDING RECURRENCE INSTANCE: [{ms_outlook_current_id[-10:]}] Outlook[{eid}] <-> GCal[{gci}]')
                        ms_outlook_start = str(ms_outlook_instance['StartUTC']).replace(':',
                                                                                        '').replace(' ',
                                                                                                    '').replace('+',
                                                                                                                '').replace('-',
                                                                                                                            '')
                        ms_outlook_end = str(ms_outlook_instance['EndUTC']).replace(':',
                                                                                    '').replace(' ',
                                                                                                '').replace('+',
                                                                                                            '').replace('-',
                                                                                                                        '')
                        eid = ms_outlook_instance['EntryID']
                        event_mapping.add_recurrent_instance(ms_outlook_current_id,
                                                             f'{eid}{ms_outlook_start}{ms_outlook_end}',
                                                             g_calendar_instance['id'])


def replicate_deletion_from_ms_outlook_to_g_calendar_single_event(local_g_calendar_connection,
                                                                  local_ms_outlook_connection,
                                                                  event_mapping,
                                                                  pause_token: PauseToken):
    print_display(f'{line_number()} Checking for deleted single events in Outlook...')
    current_ms_outlook_events = local_ms_outlook_connection.get_ms_outlook_events(pause_token)
    current_ms_outlook_ids = set(current_ms_outlook_events.keys())
    all_mappings = event_mapping.get_all_mappings()
    mapped_ms_outlook_ids = set(all_mappings['single_events'].keys())
    deleted_ms_outlook_ids = mapped_ms_outlook_ids - current_ms_outlook_ids
    for ms_outlook_id in deleted_ms_outlook_ids:
        pause_token.check()
        pair = event_mapping.get_single_event_pair(ms_outlook_id)
        if not pair or pair[1] is None:
            continue
        google_event_id = pair[1]
        google_event = local_g_calendar_connection.g_calendar_instances(google_event_id)
        if not google_event:
            print_display(f'{line_number()} Google event [{google_event_id[-10:]}] already deleted, cleaning mapping')
            event_mapping.remove_single_event(ms_outlook_id)
            continue
        if 'recurrence' not in google_event and 'recurringEventId' not in google_event:
            print_display(f'{line_number()} Deleting Google single event [{google_event_id[-10:]}] (Outlook source deleted)')
            try:
                print_display(f'{line_number()} g_calendar_delete [{google_event_id[-10:]}] {type(google_event_id)}')
                answer = local_g_calendar_connection.g_calendar_delete(google_event_id)
                print_display(f'{line_number()} answer [{answer}] {type(answer)}')
                print_display(f'{line_number()} remove_single_event [{ms_outlook_id[-10:]}] {type(ms_outlook_id)}')
                event_mapping.remove_single_event(ms_outlook_id)
            except Exception as exception:
                print_display(f'{line_number()} Error deleting Google event: {exception}')


def replicate_deletion_from_g_calendar_to_ms_outlook_single_event(local_g_calendar_connection,
                                                                  local_ms_outlook_connection,
                                                                  event_mapping,
                                                                  pause_token: PauseToken):
    print_display(f'{line_number()} Checking for deleted single events in Google Calendar...')
    current_g_calendar_events = local_g_calendar_connection.get_g_calendar_events()
    current_g_calendar_ids = set(current_g_calendar_events.keys())
    all_mappings = event_mapping.get_all_mappings()
    mapped_g_calendar_ids = set(all_mappings['single_events'].values())
    mapped_g_calendar_ids.discard(None)
    deleted_g_calendar_ids = mapped_g_calendar_ids - current_g_calendar_ids
    for g_calendar_id in deleted_g_calendar_ids:
        pause_token.check()
        pair = event_mapping.get_single_event_pair(g_calendar_id)
        if not pair or pair[0] is None:
            continue
        ms_outlook_event_id = pair[0]
        ms_outlook_event = local_ms_outlook_connection.get_ms_outlook_item(ms_outlook_event_id)
        if not ms_outlook_event:
            print_display(f'{line_number()} Outlook event [{ms_outlook_event_id[-10:]}] already deleted, cleaning mapping')
            event_mapping.remove_single_event(g_calendar_id)
            continue
        if not ms_outlook_event.get('IsRecurring',
                                    False):
            print_display(f'{line_number()} Deleting Outlook single event [{ms_outlook_event_id[-10:]}] (Google Calendar source deleted)')
            try:
                print_display(f'{line_number()} ms_outlook_delete_event [{ms_outlook_event_id[-10:]}] {type(ms_outlook_event_id)}')
                answer = local_ms_outlook_connection.ms_outlook_delete_event(ms_outlook_event_id)
                print_display(f'{line_number()} answer [{answer}] {type(answer)}')
                print_display(f'{line_number()} remove_single_event [{g_calendar_id[-10:]}] {type(g_calendar_id)}')
                event_mapping.remove_single_event(g_calendar_id)
            except Exception as exception:
                print_display(f'{line_number()} Error deleting Outlook event: {exception}')


def replicate_deletion_of_single_event_from_ms_outlook_to_g_calendar_recurrent_event(local_g_calendar_connection,
                                                                                     local_ms_outlook_connection,
                                                                                     event_mapping,
                                                                                     pause_token: PauseToken):
    print_display(f'{line_number()} Checking for deleted recurrent events in Microsoft Outlook...')
    master_pair = event_mapping.get_all_mappings()
    for ms_outlook_id in master_pair['recurrent_events']:
        pause_token.check()
        for i in master_pair['recurrent_events'][ms_outlook_id]['instances']:
            g_calendar_id = master_pair['recurrent_events'][ms_outlook_id]['instances'][i]
            d_id = extract_date_full(g_calendar_id)
            print_display(f'{line_number()} Detected deleted Google instance [{g_calendar_id[-10:]}] with date ID [{d_id}]')
            try:
                g_calendar_idr = get_master_id(g_calendar_id)
                result = local_ms_outlook_connection.get_occurrence_by_g_calendar_master_and_start(g_calendar_idr,
                                                                                                   d_id)
                if not result:
                    print_display(f'{line_number()} Deleting Google Calendar instance [{g_calendar_id[-10:]}] (Microsoft Outlook instance [{ms_outlook_id[-10:]}] was deleted)')
                    g_calendar_delete = local_g_calendar_connection.g_calendar_delete(g_calendar_id)
                    print('=' * 150)
                    print_display(f'{line_number()} [{g_calendar_delete}]')
                    print('=' * 150)
                    if g_calendar_delete != 'Failed':
                        print_display(f'{line_number()} Successfully deleted Outlook instance')
                        event_mapping.remove_recurrent_instance(g_calendar_id)
                        print_display(f'{line_number()} Marked instance as deleted in mapping')
            except ValueError as value_error:
                print_display(f'{line_number()} Error deleting Outlook instance: {value_error}')


def replicate_deletion_of_single_event_from_g_calendar_to_ms_outlook_recurrent_event(local_g_calendar_connection,
                                                                                     local_ms_outlook_connection,
                                                                                     event_mapping,
                                                                                     pause_token: PauseToken):
    print_display(f'{line_number()} Checking for deleted recurrent events in Google Calendar...')
    master_pair = event_mapping.get_all_mappings()
    for ms_outlook_id in master_pair['recurrent_events']:
        pause_token.check()
        for i in master_pair['recurrent_events'][ms_outlook_id]['instances']:
            g_calendar_id = master_pair['recurrent_events'][ms_outlook_id]['instances'][i]
            instance = local_g_calendar_connection.g_calendar_instance(g_calendar_id)
            if instance['status'] == 'cancelled':
                d_id = extract_date_full(g_calendar_id)
                print_display(f'{line_number()} Detected deleted Google instance [{g_calendar_id[-10:]}] with date ID [{d_id}]')
                try:
                    print_display(f'{line_number()} Deleting Outlook instance [{ms_outlook_id[-10:]}] (Google Calendar instance [{g_calendar_id[-10:]}] was deleted)')
                    g_calendar_idr = get_master_id(g_calendar_id)
                    ms_outlook_delete = local_ms_outlook_connection.delete_occurrence_by_g_calendar_master_and_start(g_calendar_idr,
                                                                                                                     d_id)
                    if ms_outlook_delete:
                        print_display(f'{line_number()} Successfully deleted Outlook instance')
                        event_mapping.remove_recurrent_instance(g_calendar_id)
                        print_display(f'{line_number()} Marked instance as deleted in mapping')
                except ValueError as value_error:
                    print_display(f'{line_number()} Error deleting Outlook instance: {value_error}')


def replicate_deletion_from_ms_outlook_to_g_calendar_recurrent_event(g_calendar_local_connection,
                                                                     ms_outlook_local_connection,
                                                                     event_mapping,
                                                                     pause_token: PauseToken):
    print_display(f'{line_number()} Checking for deleted recurrent event in Outlook...')
    all_mappings = event_mapping.get_all_mappings()
    recurrent_events = all_mappings['recurrent_events']
    for ms_outlook_master_id, master_data in recurrent_events.items():
        pause_token.check()
        g_calendar_master_id = master_data['g_calendar_master_id']
        g_calendar_master_idr = get_master_id(g_calendar_master_id)
        g_calendar_instance_exists = g_calendar_local_connection.g_calendar_instance(g_calendar_master_id)
        ms_outlook_instance_exists = ms_outlook_local_connection.get_master_by_g_calendar_id(g_calendar_master_idr)
        ms_outlook_master_idn = com_object_to_dictionary(ms_outlook_instance_exists)
        if not ms_outlook_master_idn and g_calendar_instance_exists['status'] != 'cancelled':
            g_calendar_local_connection.g_calendar_delete(g_calendar_master_id)
            print_display(f'{line_number()} Successfully deleted Google instance')
            event_mapping.g_calendar_remove_recurrence(g_calendar_master_idr)
            print_display(f'{line_number()} Marked instance as deleted in mapping')


def replicate_deletion_from_g_calendar_to_ms_outlook_recurrent_event(g_calendar_local_connection,
                                                                     ms_outlook_local_connection,
                                                                     event_mapping,
                                                                     pause_token: PauseToken):
    print_display(f'{line_number()} Checking for deleted recurrent event in Outlook...')
    all_mappings = event_mapping.get_all_mappings()
    recurrent_events = all_mappings['recurrent_events']
    for ms_outlook_master_id, master_data in recurrent_events.items():
        pause_token.check()
        g_calendar_master_id = master_data['g_calendar_master_id']
        g_calendar_master_idr = get_master_id(g_calendar_master_id)
        g_calendar_instance_exists = g_calendar_local_connection.g_calendar_instance(g_calendar_master_id)
        ms_outlook_instance_exists = ms_outlook_local_connection.get_master_by_g_calendar_id(g_calendar_master_idr)
        ms_outlook_master_idn = com_object_to_dictionary(ms_outlook_instance_exists)
        ms_outlook_instance_exists = False
        ms_outlook_master_index = None
        if 'EntryID' in ms_outlook_master_idn:
            ms_outlook_master_index = ms_outlook_master_idn['EntryID']
            ms_outlook_instance_exists = True
        print_display(f'{line_number()} Checking master event Outlook[{ms_outlook_instance_exists}] <-> GCal[{g_calendar_master_idr}]')
        print_display(f'{line_number()} Checking master event Outlook[{ms_outlook_instance_exists}] <-> GCal[{g_calendar_instance_exists}]')
        if g_calendar_instance_exists['status'] == 'cancelled' and ms_outlook_instance_exists:
            ms_outlook_local_connection.ms_outlook_delete_event(ms_outlook_master_index)
            print_display(f'{line_number()} Successfully deleted Google instance')
            event_mapping.remove_recurrence(ms_outlook_master_index)
            print_display(f'{line_number()} Marked instance as deleted in mapping')


def sync_outlook_to_google(local_ms_outlook_connection,
                           local_g_calendar_connection,
                           pause_token: PauseToken):
    print_display(f'{line_number()}')
    event_mapping = EventMapping()
    replicate_deletion_from_ms_outlook_to_g_calendar_single_event(local_g_calendar_connection,
                                                                  local_ms_outlook_connection,
                                                                  event_mapping,
                                                                  pause_token)
    replicate_deletion_from_g_calendar_to_ms_outlook_single_event(local_g_calendar_connection,
                                                                  local_ms_outlook_connection,
                                                                  event_mapping,
                                                                  pause_token)
    replicate_deletion_of_single_event_from_g_calendar_to_ms_outlook_recurrent_event(local_g_calendar_connection,
                                                                                     local_ms_outlook_connection,
                                                                                     event_mapping,
                                                                                     pause_token)
    replicate_deletion_of_single_event_from_ms_outlook_to_g_calendar_recurrent_event(local_g_calendar_connection,
                                                                                     local_ms_outlook_connection,
                                                                                     event_mapping,
                                                                                     pause_token)
    replicate_deletion_from_ms_outlook_to_g_calendar_recurrent_event(local_g_calendar_connection,
                                                                     local_ms_outlook_connection,
                                                                     event_mapping,
                                                                     pause_token)
    replicate_deletion_from_g_calendar_to_ms_outlook_recurrent_event(local_g_calendar_connection,
                                                                     local_ms_outlook_connection,
                                                                     event_mapping,
                                                                     pause_token)
    copy_ms_outlook_single_event_to_g_calendar(local_ms_outlook_connection,
                                               local_g_calendar_connection,
                                               event_mapping,
                                               pause_token)
    copy_g_calendar_single_event_to_ms_outlook(local_g_calendar_connection,
                                               local_ms_outlook_connection,
                                               event_mapping,
                                               pause_token)
    copy_ms_outlook_recurrent_event_to_g_calendar(local_ms_outlook_connection,
                                                  local_g_calendar_connection,
                                                  event_mapping,
                                                  pause_token)
    copy_g_calendar_recurrent_event_to_ms_outlook(local_g_calendar_connection,
                                                  local_ms_outlook_connection,
                                                  event_mapping,
                                                  pause_token)
