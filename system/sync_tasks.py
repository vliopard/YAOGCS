from connector.calendar_event import CalendarEvent
from connector.event_mapping import EventMapping
from connector.g_calendar import GoogleCalendarConnector
from connector.ms_outlook import MicrosoftOutlookConnector
from utils.handling_time import extract_date_full
from utils.utils import com_object_to_dictionary
from utils.utils import get_master_id
from utils.utils import line_number
from utils.utils import print_display
from utils.utils import print_overline
from utils.utils import print_underline
from utils.utils import sort_json_list


def replicate_deletion_from_ms_outlook_to_g_calendar_single_event(event_mapping):
    local_ms_outlook_connection = MicrosoftOutlookConnector()
    local_g_calendar_connection = GoogleCalendarConnector()

    print_display(f'{line_number()} Checking for deleted single events in [Microsoft Outlook]...')
    current_ms_outlook_events = local_ms_outlook_connection.get_ms_outlook_events()
    ms_outlook_mapped_single_events = set(event_mapping.get_all_mappings()['single_events'].keys()) - set(current_ms_outlook_events.keys())
    for ms_outlook_id in ms_outlook_mapped_single_events:
        event_pair = event_mapping.get_single_event_pair(ms_outlook_id)
        if event_pair:
            google_event_id = event_pair[1]
            google_event = local_g_calendar_connection.get_g_calendar_item(google_event_id)
            if not google_event:
                print_display(f'{line_number()} [Google Calendar] event [{google_event_id[-10:]}] already deleted, cleaning mapping...')
                event_mapping.remove_single_event(ms_outlook_id)
                continue
            if 'recurrence' not in google_event and 'recurringEventId' not in google_event:
                print_display(f'{line_number()} Deleting [Google Calendar] single event [{google_event_id[-10:]}] ([Microsoft Outlook] source deleted)')
                try:
                    print_display(f'{line_number()} g_calendar_delete_event [{google_event_id[-10:]}]')
                    local_g_calendar_connection.g_calendar_delete_event(google_event_id)
                    print_display(f'{line_number()} remove_single_event [{ms_outlook_id[-10:]}]')
                    event_mapping.remove_single_event(ms_outlook_id)
                except Exception as exception:
                    print_display(f'{line_number()} Error deleting [Google Calendar] event: [{exception}]')


def replicate_deletion_from_g_calendar_to_ms_outlook_single_event(event_mapping):
    local_ms_outlook_connection = MicrosoftOutlookConnector()
    local_g_calendar_connection = GoogleCalendarConnector()

    print_display(f'{line_number()} Checking for deleted single events in [Google Calendar]...')
    current_g_calendar_events = local_g_calendar_connection.get_g_calendar_events()
    all_mappings = event_mapping.get_all_mappings()
    mapped_g_calendar_ids = set(all_mappings['single_events'].values())
    mapped_g_calendar_ids.discard(None)
    g_calendar_mapped_single_events = mapped_g_calendar_ids - set(current_g_calendar_events.keys())
    for g_calendar_id in g_calendar_mapped_single_events:
        event_pair = event_mapping.get_single_event_pair(g_calendar_id)
        if event_pair:
            ms_outlook_event_id = event_pair[0]
            ms_outlook_event = local_ms_outlook_connection.get_ms_outlook_item(ms_outlook_event_id)
            if not ms_outlook_event:
                print_display(f'{line_number()} [Microsoft Outlook] event [{ms_outlook_event_id[-10:]}] already deleted, cleaning mapping')
                event_mapping.remove_single_event(g_calendar_id)
                continue
            if not ms_outlook_event.get('IsRecurring',
                                        False):
                print_display(f'{line_number()} Deleting [Microsoft Outlook] single event [{ms_outlook_event_id[-10:]}] ([Google Calendar] source deleted)')
                try:
                    print_display(f'{line_number()} ms_outlook_delete_event [{ms_outlook_event_id[-10:]}]')
                    local_ms_outlook_connection.ms_outlook_delete_event(ms_outlook_event_id)
                    print_display(f'{line_number()} remove_single_event [{g_calendar_id[-10:]}]')
                    event_mapping.remove_single_event(g_calendar_id)
                except Exception as exception:
                    print_display(f'{line_number()} Error deleting [Microsoft Outlook] event: [{exception}]')


def replicate_deletion_of_single_event_from_g_calendar_to_ms_outlook_recurrent_event(event_mapping):
    local_ms_outlook_connection = MicrosoftOutlookConnector()
    local_g_calendar_connection = GoogleCalendarConnector()

    print_display(f'{line_number()} Checking for deleted recurrent events in [Google Calendar]...')
    master_pair = event_mapping.get_all_mappings()
    for ms_outlook_id in master_pair['recurrent_events']:
        for instance_event in master_pair['recurrent_events'][ms_outlook_id]['instances']:
            g_calendar_id = master_pair['recurrent_events'][ms_outlook_id]['instances'][instance_event]
            g_calendar_instance = local_g_calendar_connection.get_g_calendar_item(g_calendar_id)
            if g_calendar_instance['status'] == 'cancelled':
                g_calendar_date_item = extract_date_full(g_calendar_id)
                print_display(f'{line_number()} Detected deleted [Google Calendar] instance [{g_calendar_id[-10:]}] with date ID [{g_calendar_date_item}]')
                try:
                    print_display(f'{line_number()} Deleting [Microsoft Outlook] instance [{ms_outlook_id[-10:]}] ([Google Calendar] instance [{g_calendar_id[-10:]}] was deleted)')
                    g_calendar_id_master = get_master_id(g_calendar_id)
                    ms_outlook_delete = local_ms_outlook_connection.delete_occurrence_by_g_calendar_master_and_start(g_calendar_id_master,
                                                                                                                     g_calendar_date_item)
                    if ms_outlook_delete:
                        print_display(f'{line_number()} Successfully deleted [Microsoft Outlook] instance')
                        event_mapping.remove_recurrent_instance(g_calendar_id)
                        print_display(f'{line_number()} Deleted instance in mapping...')
                except ValueError as value_error:
                    print_display(f'{line_number()} Error deleting [Microsoft Outlook] instance: [{value_error}]')


def replicate_deletion_of_single_event_from_ms_outlook_to_g_calendar_recurrent_event(event_mapping):
    local_ms_outlook_connection = MicrosoftOutlookConnector()
    local_g_calendar_connection = GoogleCalendarConnector()

    print_display(f'{line_number()} Checking for deleted recurrent events in [Microsoft Outlook]...')
    master_pair = event_mapping.get_all_mappings()
    for ms_outlook_id in master_pair['recurrent_events']:
        for instance_event in master_pair['recurrent_events'][ms_outlook_id]['instances']:
            g_calendar_id = master_pair['recurrent_events'][ms_outlook_id]['instances'][instance_event]
            g_calendar_date_item = extract_date_full(g_calendar_id)
            print_display(f'{line_number()} Detected deleted [Google Calendar] instance [{g_calendar_id[-10:]}] with date ID [{g_calendar_date_item}]')
            try:
                g_calendar_id_master = get_master_id(g_calendar_id)
                result = local_ms_outlook_connection.get_occurrence_by_g_calendar_master_and_start(g_calendar_id_master,
                                                                                                   g_calendar_date_item)
                if not result:
                    print_display(f'{line_number()} Deleting [Google Calendar] instance [{g_calendar_id[-10:]}] ([Microsoft Outlook] instance [{ms_outlook_id[-10:]}] was deleted)')
                    g_calendar_delete = local_g_calendar_connection.g_calendar_delete_event(g_calendar_id)
                    if g_calendar_delete != 'Failed':
                        print_display(f'{line_number()} Successfully deleted [Google Calendar] instance')
                        event_mapping.remove_recurrent_instance(g_calendar_id)
                        print_display(f'{line_number()} Deleted instance in mapping...')
            except ValueError as value_error:
                print_display(f'{line_number()} Error deleting [Microsoft Outlook] instance: {value_error}')


def replicate_deletion_from_ms_outlook_to_g_calendar_recurrent_event(event_mapping):
    ms_outlook_local_connection = MicrosoftOutlookConnector()
    g_calendar_local_connection = GoogleCalendarConnector()

    print_display(f'{line_number()} Checking for deleted recurrent event in [Microsoft Outlook]...')
    recurrent_events = event_mapping.get_all_mappings()['recurrent_events']
    for ms_outlook_master_id, master_data in recurrent_events.items():
        g_calendar_id = master_data['g_calendar_master_id']
        g_calendar_master_id = get_master_id(g_calendar_id)
        g_calendar_instance_exists = g_calendar_local_connection.get_g_calendar_item(g_calendar_id)
        ms_outlook_instance_exists = ms_outlook_local_connection.get_master_by_g_calendar_id(g_calendar_master_id)
        ms_outlook_master_id_item = com_object_to_dictionary(ms_outlook_instance_exists)
        if not ms_outlook_master_id_item and g_calendar_instance_exists['status'] != 'cancelled':
            g_calendar_local_connection.g_calendar_delete_event(g_calendar_id)
            print_display(f'{line_number()} Successfully deleted [Google Calendar] instance')
            event_mapping.g_calendar_remove_recurrence(g_calendar_master_id)
            print_display(f'{line_number()} Deleted instance in mapping')


def replicate_deletion_from_g_calendar_to_ms_outlook_recurrent_event(event_mapping):
    ms_outlook_local_connection = MicrosoftOutlookConnector()
    g_calendar_local_connection = GoogleCalendarConnector()

    print_display(f'{line_number()} Checking for deleted recurrent event in [Google Calendar]...')
    recurrent_events = event_mapping.get_all_mappings()['recurrent_events']
    for ms_outlook_master_id, master_data in recurrent_events.items():
        g_calendar_id = master_data['g_calendar_master_id']
        g_calendar_master_id = get_master_id(g_calendar_id)
        g_calendar_instance_exists = g_calendar_local_connection.get_g_calendar_item(g_calendar_id)
        ms_outlook_instance_exists = ms_outlook_local_connection.get_master_by_g_calendar_id(g_calendar_master_id)
        ms_outlook_master_id_item = com_object_to_dictionary(ms_outlook_instance_exists)
        if 'EntryID' in ms_outlook_master_id_item:
            ms_outlook_master_index = ms_outlook_master_id_item['EntryID']
            g_calendar_instance_id = g_calendar_instance_exists['id'][-10:]
            print_display(f'{line_number()} Checking master event [Microsoft Outlook] [{ms_outlook_master_index}] <=> [Google Calendar] [{g_calendar_master_id[-10:]}]')
            print_display(f'{line_number()} Checking master event [Microsoft Outlook] [{ms_outlook_master_index}] <=> [Google Calendar] [{g_calendar_instance_id}]')
            if g_calendar_instance_exists['status'] == 'cancelled':
                ms_outlook_local_connection.ms_outlook_delete_event(ms_outlook_master_index)
                print_display(f'{line_number()} Successfully deleted [Google Calendar] instance')
                event_mapping.remove_recurrence(ms_outlook_master_index)
                print_display(f'{line_number()} Deleted instance in mapping...')


def copy_ms_outlook_single_event_to_g_calendar(event_mapping):
    local_ms_outlook_connection = MicrosoftOutlookConnector()
    local_g_calendar_connection = GoogleCalendarConnector()

    print_display(f'{line_number()} Checking for new single events in [Microsoft Outlook]...')
    ms_outlook_events = local_ms_outlook_connection.get_ms_outlook_events()
    for ms_outlook_current_id, ms_outlook_current_event in ms_outlook_events.items():
        if not ms_outlook_current_event.get('IsRecurring',
                                            False):
            single_pair = event_mapping.get_single_event_pair(ms_outlook_current_id)
            if not single_pair:
                calendar_event = CalendarEvent()
                calendar_event.import_ms_outlook(ms_outlook_current_event)
                g_calendar_exported_event = calendar_event.export_g_calendar()
                print_display(f'{line_number()} INSERTING EVENT: [{ms_outlook_current_id[-10:]}]')
                g_calendar_inserted_appointment = local_g_calendar_connection.g_calendar_insert(g_calendar_exported_event)
                if g_calendar_inserted_appointment:
                    g_calendar_master_id = g_calendar_inserted_appointment.get('id')
                    print_display(f'{line_number()} ADDING EVENT: [{ms_outlook_current_id[-10:]}] -> [{g_calendar_master_id[-10:]}]')
                    event_mapping.add_single_event(ms_outlook_current_id,
                                                   g_calendar_master_id)


def copy_g_calendar_single_event_to_ms_outlook(event_mapping):
    local_ms_outlook_connection = MicrosoftOutlookConnector()
    local_g_calendar_connection = GoogleCalendarConnector()

    print_display(f'{line_number()} Checking for new single events in [Google Calendar]...')
    g_calendar_all_events = local_g_calendar_connection.get_g_calendar_events()
    for g_calendar_event_id, g_calendar_event_item in g_calendar_all_events.items():
        recurrence_one = 'recurrence' in g_calendar_event_item
        recurrence_two = 'recurringEventId' in g_calendar_event_item
        if not recurrence_one and not recurrence_two:
            single_pair = event_mapping.get_single_event_pair(g_calendar_event_id)
            if not single_pair:
                calendar_event = CalendarEvent()
                calendar_event.import_g_calendar(g_calendar_event_item)
                ms_outlook_exported_event = calendar_event.export_ms_outlook()
                print_display(f'{line_number()} INSERTING EVENT: [{g_calendar_event_id[-10:]}]')
                ms_outlook_inserted_appointment = local_ms_outlook_connection.ms_outlook_insert(ms_outlook_exported_event)
                if ms_outlook_inserted_appointment:
                    ms_outlook_event_id = ms_outlook_inserted_appointment.EntryID
                    print_display(f'{line_number()} ADDING EVENT: [{g_calendar_event_id[-10:]}] -> [{ms_outlook_event_id[-10:]}]')
                    event_mapping.add_single_event(ms_outlook_event_id,
                                                   g_calendar_event_id)


def copy_ms_outlook_recurrent_event_to_g_calendar(event_mapping):
    local_ms_outlook_connection = MicrosoftOutlookConnector()
    local_g_calendar_connection = GoogleCalendarConnector()

    print_display(f'{line_number()} Checking for new recurrent events in [Microsoft Outlook]...')
    ms_outlook_events = local_ms_outlook_connection.get_ms_outlook_events()
    for ms_outlook_current_id, ms_outlook_current_event in ms_outlook_events.items():
        if ms_outlook_current_event.get('IsRecurring',
                                        True):
            master_pair = event_mapping.get_recurrent_master_pair(ms_outlook_current_id)
            if not master_pair:
                calendar_event = CalendarEvent()
                calendar_event.import_ms_outlook(ms_outlook_current_event)
                g_calendar_exported_event = calendar_event.export_g_calendar()
                g_calendar_inserted_appointment = local_g_calendar_connection.g_calendar_insert(g_calendar_exported_event)
                if g_calendar_inserted_appointment:
                    g_calendar_id = g_calendar_inserted_appointment.get('id')
                    g_calendar_master_id = get_master_id(g_calendar_id)
                    print_display(f'{line_number()} ADDING RECURRENCE MASTER: [{ms_outlook_current_id[-10:]}] => [{g_calendar_id[-10:]}]')
                    event_mapping.add_recurrent_master(ms_outlook_current_id,
                                                       g_calendar_id)
                    ms_outlook_instances = local_ms_outlook_connection.get_recurrence_instances(ms_outlook_current_id)
                    g_calendar_instances = local_g_calendar_connection.g_calendar_instances(g_calendar_id).get('items',
                                                                                                               [])
                    local_ms_outlook_connection.set_recurrence_id(ms_outlook_current_id,
                                                                  g_calendar_master_id)
                    for ms_outlook_instance, g_calendar_instance in zip(ms_outlook_instances,
                                                                        sort_json_list(g_calendar_instances,
                                                                                       'start.dateTime')):
                        ms_outlook_instance_string = ms_outlook_instance['EntryID'][-10:]
                        g_calendar_instance_string = g_calendar_instance['id'][-10:]
                        print_display(f'{line_number()} ADDING RECURRENCE INSTANCE: [{ms_outlook_current_id[-10:]}] [Microsoft Outlook] [{ms_outlook_instance_string}] <=> [Google Calendar] [{g_calendar_instance_string}]')
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
                        ms_outlook_instance_string = ms_outlook_instance['EntryID']
                        event_mapping.add_recurrent_instance(ms_outlook_current_id,
                                                             f'{ms_outlook_instance_string}{ms_outlook_start}{ms_outlook_end}',
                                                             g_calendar_instance['id'])


def copy_g_calendar_recurrent_event_to_ms_outlook(event_mapping):
    local_ms_outlook_connection = MicrosoftOutlookConnector()
    local_g_calendar_connection = GoogleCalendarConnector()

    print_display(f'{line_number()} Checking for new recurrent events in [Google Calendar]...')
    g_calendar_events_local = local_g_calendar_connection.get_g_calendar_events()
    g_calendar_total_items = g_calendar_events_local.items()
    g_calendar_total_items_progress = 0
    g_calendar_total_items_count = len(g_calendar_total_items)
    for g_calendar_event_id, g_calendar_event_data in g_calendar_total_items:
        g_calendar_total_items_progress += 1
        if 'recurrence' in g_calendar_event_data:
            master_pair = event_mapping.get_recurrent_master_pair(g_calendar_event_id)
            if not master_pair:
                calendar_event = CalendarEvent()
                calendar_event.import_g_calendar(g_calendar_event_data)
                ms_outlook_exported_event = calendar_event.export_ms_outlook()
                ms_outlook_inserted_appointment = local_ms_outlook_connection.ms_outlook_insert(ms_outlook_exported_event)
                if ms_outlook_inserted_appointment:
                    ms_outlook_entry_id = ms_outlook_inserted_appointment.EntryID
                    print_display(f'{line_number()} 01-({g_calendar_total_items_progress}/{g_calendar_total_items_count}) ADDING RECURRENCE MASTER: [{g_calendar_event_id[-10:]}] => [{ms_outlook_entry_id[-10:]}]')
                    event_mapping.add_recurrent_master(ms_outlook_entry_id,
                                                       g_calendar_event_id)
                    g_calendar_instances = local_g_calendar_connection.g_calendar_instances(g_calendar_event_id).get('items',
                                                                                                                     [])
                    ms_outlook_instances = local_ms_outlook_connection.get_recurrence_instances(ms_outlook_entry_id)
                    local_ms_outlook_connection.set_recurrence_id(ms_outlook_entry_id,
                                                                  g_calendar_event_id)
                    ms_outlook_total_items_progress = 0
                    ms_outlook_total_items_count = len(ms_outlook_instances)
                    for g_calendar_instance, ms_outlook_instance in zip(sort_json_list(g_calendar_instances,
                                                                                       'start.dateTime'),
                                                                        ms_outlook_instances):
                        ms_outlook_total_items_progress += 1
                        g_calendar_instance_string = g_calendar_instance['id'][-10:]
                        ms_outlook_instance_string = ms_outlook_instance['EntryID'][-10:]
                        print_display(f'{line_number()} 02-({ms_outlook_total_items_progress}/{ms_outlook_total_items_count}) ADDING RECURRENCE INSTANCE: [{ms_outlook_entry_id[-10:]}] [Google Calendar] [{g_calendar_instance_string}] <=> [Microsoft Outlook] [{ms_outlook_instance_string}]')
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
                        ms_outlook_entry_id_string = ms_outlook_instance['EntryID']
                        event_mapping.add_recurrent_instance(ms_outlook_entry_id,
                                                             f'{ms_outlook_entry_id_string}{ms_outlook_start}{ms_outlook_end}',
                                                             g_calendar_instance['id'])


def sync_task(event_mapping):
    ms_outlook_to_g_calendar = 'Microsoft Outlook to Google Calendar'
    g_calendar_to_ms_outlook = 'Google Calendar to Microsoft Outlook'
    ways = [ms_outlook_to_g_calendar,

            # g_calendar_to_ms_outlook
            ]
    print_underline()
    if ms_outlook_to_g_calendar in ways and g_calendar_to_ms_outlook in ways:
        print_display(f'{line_number()} Starting synchronization task: [Microsoft Outlook] <=> [Google Calendar]')
    elif ms_outlook_to_g_calendar in ways:
        print_display(f'{line_number()} Starting synchronization task: [Microsoft Outlook] => [Google Calendar]')
    elif g_calendar_to_ms_outlook in ways:
        print_display(f'{line_number()} Starting synchronization task: [Google Calendar] => [Microsoft Outlook]')
    print_overline()
    if ms_outlook_to_g_calendar in ways:
        # Microsoft Outlook to Google Calendar
        replicate_deletion_from_ms_outlook_to_g_calendar_single_event(event_mapping)
        replicate_deletion_of_single_event_from_ms_outlook_to_g_calendar_recurrent_event(event_mapping)
        replicate_deletion_from_ms_outlook_to_g_calendar_recurrent_event(event_mapping)
        copy_ms_outlook_single_event_to_g_calendar(event_mapping)
        copy_ms_outlook_recurrent_event_to_g_calendar(event_mapping)
    if g_calendar_to_ms_outlook in ways:
        # Google Calendar to Microsoft Outlook
        replicate_deletion_from_g_calendar_to_ms_outlook_single_event(event_mapping)
        replicate_deletion_of_single_event_from_g_calendar_to_ms_outlook_recurrent_event(event_mapping)
        replicate_deletion_from_g_calendar_to_ms_outlook_recurrent_event(event_mapping)
        copy_g_calendar_single_event_to_ms_outlook(event_mapping)
        copy_g_calendar_recurrent_event_to_ms_outlook(event_mapping)

    # TODO: IF EVENT INFORMATION CHANGES, SYNC DATA (HOW TO KNOW WHO CHANGED?) HASH1<=>HASH1 / HASH1<=>HASH_C / HASH_C<=>HASH1


if __name__ == '__main__':
    event_mapping = EventMapping()
    event_mapping.reset()
    sync_task(event_mapping)
