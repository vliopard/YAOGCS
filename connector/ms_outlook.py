import gc
from datetime import datetime
from datetime import timedelta
from datetime import timezone
from time import sleep

import pywintypes
import win32com.client

import system.constants as constants
from system.tools import convert_com_object_to_dictionary
from system.tools import line_number
from system.tools import print_display
from system.tools import print_overline
from system.tools import print_underline
from system.tools import release_com_object_memory
from system.tools import strip_symbols
from system.tools import trim_id


class MicrosoftOutlookHelper:
    def __init__(self):
        self.ms_outlook_client = win32com.client.Dispatch('Outlook.Application')
        self.ms_outlook_namespace = self.ms_outlook_client.GetNamespace('MAPI')
        self.ms_outlook_calendar = self.ms_outlook_namespace.GetDefaultFolder(9)

    def ms_outlook_create_item(self):
        return self.ms_outlook_client.CreateItem(1)

    def ms_outlook_get_all_instances(self):
        return self.ms_outlook_calendar.Items

    def ms_outlook_get_item(self,
                            ms_outlook_instance_id):
        return self.ms_outlook_namespace.GetItemFromID(ms_outlook_instance_id.split('_')[0],
                                                       self.ms_outlook_calendar.StoreID)


class MicrosoftOutlookConnector:
    def __init__(self):
        self.ms_outlook_data = MicrosoftOutlookHelper()

    def get_restriction(self,
                        ms_outlook_all_instances):
        time_now = datetime.now()
        time_begin = time_now - timedelta(days=constants.DAY_PAST)
        time_end = time_now + timedelta(days=constants.DAY_NEXT)
        ms_outlook_all_instances.IncludeRecurrences = True
        ms_outlook_all_instances.Sort('[Start]')
        restriction_string = "([Start] >= '{}' OR [End] >= '{}') AND [End] <= '{}'"
        restriction = restriction_string.format(time_begin.strftime('%m/%d/%Y %H:%M %p'),
                                                time_begin.strftime('%m/%d/%Y %H:%M %p'),
                                                time_end.strftime('%m/%d/%Y %H:%M %p'))
        return ms_outlook_all_instances.Restrict(restriction)

    def get_item_ms_outlook(self,
                            ms_outlook_instance_id):
        return convert_com_object_to_dictionary(self.ms_outlook_data.ms_outlook_get_item(ms_outlook_instance_id))

    def get_occurrence_ms_outlook(self,
                                  ms_outlook_instance_id,
                                  ms_outlook_start_date):
        try:
            ms_outlook_appointment = self.ms_outlook_data.ms_outlook_get_item(ms_outlook_instance_id)
            if not ms_outlook_appointment:
                return None
            if not ms_outlook_appointment.IsRecurring:
                return None
            ms_outlook_recurrence = ms_outlook_appointment.GetRecurrencePattern()
            ms_outlook_occurrence = ms_outlook_recurrence.GetOccurrence(datetime.strptime(ms_outlook_start_date,
                                                                                          '%Y-%m-%d'))
            return convert_com_object_to_dictionary(ms_outlook_occurrence)
        except pywintypes.com_error as com_error:
            print_display(f'{line_number()} [Microsoft Outlook] COM ERROR: {com_error}')
            return None

    def get_instance_data_ms_outlook(self,
                                     ms_outlook_instance,
                                     ms_outlook_properties):
        ms_outlook_instance_data = dict()
        ms_outlook_selected_instances_length = len(ms_outlook_properties)
        for ms_outlook_index, ms_outlook_property in enumerate(ms_outlook_properties):
            if ms_outlook_property == 'GetInspector':
                continue
            try:
                ms_outlook_attributes = getattr(ms_outlook_instance,
                                                ms_outlook_property)
                if callable(ms_outlook_attributes):
                    continue
                '''
                if item_property == 'Recipients':
                    attendees = []
                    for recipient in item.Recipients:
                        attendees.append({
                                'email': recipient.Address})
                        _release(recipient)
                    local_event_data[item_property] = attendees
                '''
                if ms_outlook_property in ('Start',
                                           'End'):
                    if hasattr(ms_outlook_attributes,
                               'Format'):
                        ms_outlook_instance_data[ms_outlook_property] = ms_outlook_attributes.Format('%Y-%m-%dT%H:%M:%S')
                    else:
                        ms_outlook_instance_data[ms_outlook_property] = ms_outlook_attributes
                elif ms_outlook_property == 'Body':
                    ms_outlook_instance_data[ms_outlook_property] = ms_outlook_attributes[:200] if ms_outlook_attributes else ''
                elif ms_outlook_property == 'IsRecurring' and ms_outlook_attributes:
                    ms_outlook_instance_data[ms_outlook_property] = ms_outlook_attributes
                    ms_outlook_recurrence_pattern = ms_outlook_instance.GetRecurrencePattern()
                    ms_outlook_instance_data['recurrence_type'] = ms_outlook_recurrence_pattern.RecurrenceType
                    ms_outlook_instance_data['recurrence_interval'] = ms_outlook_recurrence_pattern.Interval
                    ms_outlook_instance_data['recurrence_end'] = ms_outlook_recurrence_pattern.PatternEndDate.Format('%Y-%m-%d')
                    release_com_object_memory(ms_outlook_recurrence_pattern)
                else:
                    ms_outlook_instance_data[ms_outlook_property] = ms_outlook_attributes
            except pywintypes.com_error as com_error_type:
                print_display(f'{line_number()} [Microsoft Outlook] COM ERROR for property [{ms_outlook_index}/{ms_outlook_selected_instances_length}] [{ms_outlook_property}] of item [{ms_outlook_instance.Subject}]: [{com_error_type}]')
                continue
        return ms_outlook_instance_data

    def get_all_instances_ms_outlook(self):
        ms_outlook_all_instances = self.ms_outlook_data.ms_outlook_get_all_instances()
        ms_outlook_selected_instances = self.get_restriction(ms_outlook_all_instances)
        ms_outlook_selected_instances_length = ms_outlook_selected_instances.Count
        ms_outlook_instances = dict()
        for ms_outlook_index, ms_outlook_instance in enumerate(ms_outlook_selected_instances):
            try:
                ms_outlook_properties = [ms_outlook_attributes for ms_outlook_attributes in dir(ms_outlook_instance) if not ms_outlook_attributes.startswith('_')]
            except Exception as exception:
                print_display(f'{line_number()} [Microsoft Outlook] SKIPPING ITEM: [{ms_outlook_index}/{ms_outlook_selected_instances_length}] — dir() - FAILED: [{exception}]')
                release_com_object_memory(ms_outlook_instance)
                continue
            ms_outlook_instance_data = self.get_instance_data_ms_outlook(ms_outlook_instance,
                                                                         ms_outlook_properties)
            if 'EntryID' not in ms_outlook_instance_data:
                print_display(f'{line_number()} [Microsoft Outlook] SKIPPING ITEM: [{ms_outlook_index}/{ms_outlook_selected_instances_length}] — missing EntryID')
                release_com_object_memory(ms_outlook_instance)
                continue

            ms_outlook_entry_id = ms_outlook_instance_data['EntryID'] + '_' + strip_symbols(ms_outlook_instance_data['StartUTC'])
            ms_outlook_instances[ms_outlook_entry_id] = ms_outlook_instance_data
            release_com_object_memory(ms_outlook_instance)
            if ms_outlook_index % 50 == 0:
                gc.collect()
        gc.collect()
        return ms_outlook_instances

    def get_master_by_g_calendar_id(self,
                                    g_calendar_master_id: str):
        ms_outlook_helper = MicrosoftOutlookHelper()
        ms_outlook_all_instances = ms_outlook_helper.ms_outlook_get_all_instances()
        ms_outlook_all_instances.IncludeRecurrences = False
        time_now = datetime.now()
        time_begin = (time_now - timedelta(days=constants.DAY_PAST)).strftime('%m/%d/%Y %H:%M %p')
        time_end = (time_now + timedelta(days=constants.DAY_NEXT)).strftime('%m/%d/%Y %H:%M %p')
        restriction_string = f"[Start] >= '{time_begin}' AND [Start] <= '{time_end}'"
        ms_outlook_selected_instances = ms_outlook_all_instances.Restrict(restriction_string)
        for ms_outlook_instance in ms_outlook_selected_instances:
            if not ms_outlook_instance.IsRecurring:
                continue
            try:
                print_display(f'{line_number()} [Microsoft Outlook] Checking item [{ms_outlook_instance.Subject}] (IsRecurring: [{ms_outlook_instance.IsRecurring}])')
                ms_outlook_property = ms_outlook_instance.UserProperties.Find('GCalendarMasterID')
                ms_outlook_property_value = ms_outlook_property.Value if ms_outlook_property else 'NOT SET'
                print_display(f'{line_number()} [Microsoft Outlook] GCalendarMasterID for item [{ms_outlook_instance.Subject}]: [{trim_id(ms_outlook_property_value)}]')
                print_display(f'{line_number()} [Microsoft Outlook] GCalendarMasterID for item [{ms_outlook_instance.Subject}]: [{trim_id(g_calendar_master_id)}]')
                if ms_outlook_property and ms_outlook_property.Value == g_calendar_master_id:
                    print_display(f'{line_number()} [Microsoft Outlook] Found master [{ms_outlook_instance.Subject}] for GCalendarMasterID [{trim_id(g_calendar_master_id)}]')
                    return ms_outlook_instance
            except Exception as exception:
                print_display(f'{line_number()} [Microsoft Outlook] Error checking GCalendarMasterID for item [{ms_outlook_instance.Subject}]: [{exception}]')
                continue
        print_display(f'{line_number()} [Microsoft Outlook] Master not found for GCalendarMasterID [{trim_id(g_calendar_master_id)}]')
        return None

    def get_occurrence_by_g_calendar_master_and_start(self,
                                                      g_calendar_master_id: str,
                                                      g_calendar_start_date: str):
        try:
            g_calendar_start_date_utc = datetime.strptime(g_calendar_start_date,
                                                          '%Y-%m-%d-%H-%M-%S').replace(tzinfo=timezone.utc)
            print_display(f'{line_number()} UTC [{g_calendar_start_date_utc}]')
        except ValueError as value_error:
            raise ValueError(f'[Microsoft Outlook] Invalid start_utc format: [{value_error}]')
        g_calendar_master_id_item = self.get_master_by_g_calendar_id(g_calendar_master_id)
        if not g_calendar_master_id_item:
            print_display(f'{line_number()} [Microsoft Outlook] MASTER ID not found: [{trim_id(g_calendar_master_id)}]')
            return None
        if not g_calendar_master_id_item.IsRecurring:
            print_display(f'{line_number()} [Microsoft Outlook] Item is not recurring')
            raise ValueError('[Microsoft Outlook] Item is not recurring')
        g_calendar_date_offset = g_calendar_master_id_item.StartUTC.replace(tzinfo=None) - g_calendar_master_id_item.Start.replace(tzinfo=None)
        g_calendar_local_date = g_calendar_start_date_utc - g_calendar_date_offset
        print_display(f'{line_number()} [Microsoft Outlook] master [{g_calendar_local_date}] [{g_calendar_master_id_item.Subject}]')
        g_calendar_recurrence = g_calendar_master_id_item.GetRecurrencePattern()
        print_display(f'{line_number()} [Microsoft Outlook] recurrence [{g_calendar_recurrence}]')
        try:
            g_calendar_occurrence = g_calendar_recurrence.GetOccurrence(g_calendar_local_date)
            return g_calendar_occurrence
        except Exception as value_error:
            print_display(f'{line_number()} [Microsoft Outlook] Occurrence not found: [{value_error}]')
        return None

    def set_recurrence_id(self,
                          ms_outlook_master_id,
                          g_calendar_master_id):
        try:
            print_display(f'{line_number()} 01) Setting GCalendarMasterID [{trim_id(g_calendar_master_id)}] for [Microsoft Outlook] master [{trim_id(ms_outlook_master_id)}]')
            ms_outlook_appointment = self.ms_outlook_data.ms_outlook_get_item(ms_outlook_master_id)
            print_display(f'{line_number()} 02) Appointment retrieved for master ID [{trim_id(ms_outlook_master_id)}]: [{ms_outlook_appointment.Subject}]')
            if ms_outlook_appointment:
                print_display(f'{line_number()} 03) Setting GCalendarMasterID for appointment [{ms_outlook_appointment.Subject}]')
                ms_outlook_appointment.UserProperties.Add('GCalendarMasterID',
                                                          1)
                print_display(f'{line_number()} 04) GCalendarMasterID property added for appointment [{ms_outlook_appointment.Subject}]')
                ms_outlook_appointment.UserProperties['GCalendarMasterID'].Value = g_calendar_master_id
                print_display(f'{line_number()} 05) GCalendarMasterID set successfully for [Microsoft Outlook] master [{trim_id(ms_outlook_master_id)}]')
                ms_outlook_appointment.Save()
        except Exception as exception:
            print_display(f'{line_number()} 06) Error setting GCalendarMasterID [{trim_id(ms_outlook_master_id)}][{trim_id(g_calendar_master_id)}] for [Microsoft Outlook] item: [{exception}]')

    def get_recurrence_instances(self,
                                 ms_outlook_instance_id):
        # TODO: 18 TO 180
        ms_outlook_recurrence = self.ms_outlook_data.ms_outlook_get_item(ms_outlook_instance_id)
        if not ms_outlook_recurrence.IsRecurring:
            raise ValueError('[Microsoft Outlook] Provided ID is not a recurring appointment')
        ms_outlook_recurrence_pattern = ms_outlook_recurrence.GetRecurrencePattern()
        ms_outlook_recurrence_list = list()
        ms_outlook_recurrence_start = ms_outlook_recurrence.Start
        ms_outlook_recurrence_end = ms_outlook_recurrence.End
        ms_outlook_recurrence_current = ms_outlook_recurrence_start
        while ms_outlook_recurrence_current.date() <= ms_outlook_recurrence_pattern.PatternEndDate.date():
            try:
                ms_outlook_recurrence_item = ms_outlook_recurrence_pattern.GetOccurrence(ms_outlook_recurrence_current)
                ms_outlook_recurrence_list.append(convert_com_object_to_dictionary(ms_outlook_recurrence_item))
            except pywintypes.com_error:
                pass
            ms_outlook_recurrence_current = ms_outlook_recurrence_current + (ms_outlook_recurrence_end - ms_outlook_recurrence_start)
        release_com_object_memory(ms_outlook_recurrence_pattern)
        release_com_object_memory(ms_outlook_recurrence)
        gc.collect()
        return ms_outlook_recurrence_list

    def insert_instance_ms_outlook(self,
                                   ms_outlook_instance_body):
        try:
            print_display(f'{line_number()} [Microsoft Outlook] INSERT <<==')
            ms_outlook_appointment = self.ms_outlook_data.ms_outlook_create_item()
            ms_outlook_appointment_subject = ms_outlook_instance_body.get('Subject',
                                                                          '')
            if ms_outlook_appointment_subject is None:
                ms_outlook_appointment_subject = ''
            ms_outlook_appointment.Subject = str(ms_outlook_appointment_subject)
            ms_outlook_appointment_body = ms_outlook_instance_body.get('Body',
                                                                       '')
            if ms_outlook_appointment_body is None:
                ms_outlook_appointment_body = ''
            ms_outlook_appointment.Body = str(ms_outlook_appointment_body)

            '''
            organizer = event_body.get('Organizer',
                                       '')
            if organizer is None:
                organizer = ''
            appointment.Organizer = str(organizer)
            '''

            ms_outlook_appointment_location = ms_outlook_instance_body.get('Location',
                                                                           '')
            if ms_outlook_appointment_location is None:
                ms_outlook_appointment_location = ''
            ms_outlook_appointment.Location = str(ms_outlook_appointment_location)
            ms_outlook_start_date_utc = ms_outlook_instance_body.get('StartUTC',
                                                                     '')
            ms_outlook_end_date_utc = ms_outlook_instance_body.get('EndUTC',
                                                                   '')
            if ms_outlook_start_date_utc:
                if isinstance(ms_outlook_start_date_utc,
                              str):
                    if 'T' in ms_outlook_start_date_utc:
                        ms_outlook_start_date = datetime.fromisoformat(ms_outlook_start_date_utc.replace('Z',
                                                                                                         '+00:00'))
                    else:
                        ms_outlook_start_date = datetime.fromisoformat(ms_outlook_start_date_utc)
                else:
                    ms_outlook_start_date = ms_outlook_start_date_utc
                ms_outlook_appointment.Start = ms_outlook_start_date
            if ms_outlook_end_date_utc:
                if isinstance(ms_outlook_end_date_utc,
                              str):
                    if 'T' in ms_outlook_end_date_utc:
                        recurrence_end_date = datetime.fromisoformat(ms_outlook_end_date_utc.replace('Z',
                                                                                                     '+00:00'))
                    else:
                        recurrence_end_date = datetime.fromisoformat(ms_outlook_end_date_utc)
                else:
                    recurrence_end_date = ms_outlook_end_date_utc
                ms_outlook_appointment.End = recurrence_end_date
            reminder_minutes = ms_outlook_instance_body.get('ReminderMinutesBeforeStart',
                                                            15)
            if reminder_minutes is not None:
                ms_outlook_appointment.ReminderSet = True
                ms_outlook_appointment.ReminderMinutesBeforeStart = int(reminder_minutes)
            sensitivity = ms_outlook_instance_body.get('Sensitivity',
                                                       0)
            if sensitivity is not None:
                ms_outlook_appointment.Sensitivity = int(sensitivity)
            busy_status = ms_outlook_instance_body.get('BusyStatus',
                                                       2)
            if busy_status is not None:
                ms_outlook_appointment.BusyStatus = int(busy_status)
            '''
            required_attendees = event_body.get('RequiredAttendees',
                                                '')
            optional_attendees = event_body.get('OptionalAttendees',
                                                '')
            if required_attendees:
                for email in str(required_attendees).split(';'):
                    email = email.strip()
                    if email:
                        recipient = appointment.Recipients.Add(email)
                        recipient.Type = 1  # 1 = Required
            if optional_attendees:
                for email in str(optional_attendees).split(';'):
                    email = email.strip()
                    if email:
                        recipient = appointment.Recipients.Add(email)
                        recipient.Type = 2  # 2 = Optional
            if appointment.Recipients.Count > 0:
                appointment.Recipients.ResolveAll()
            '''
            if ms_outlook_instance_body.get('IsRecurring'):
                recurrence = ms_outlook_appointment.GetRecurrencePattern()
                recurrence.RecurrenceType = int(ms_outlook_instance_body.get('recurrence_type',
                                                                             0))
                recurrence.Interval = int(ms_outlook_instance_body.get('recurrence_interval',
                                                                       1))
                recurrence.PatternStartDate = ms_outlook_appointment.Start
                recurrence_end = ms_outlook_instance_body.get('recurrence_end')
                if recurrence_end:
                    recurrence_end_date = datetime.strptime(recurrence_end,
                                                            '%Y-%m-%d')
                    recurrence_end_date = recurrence_end_date.replace(hour=ms_outlook_appointment.Start.hour,
                                                                      minute=ms_outlook_appointment.Start.minute,
                                                                      second=ms_outlook_appointment.Start.second)
                    print_display(f'{line_number()} [Microsoft Outlook] Setting recurrence end date: [{recurrence_end_date}] [{ms_outlook_appointment.Subject}]')
                    try:
                        recurrence.PatternEndDate = recurrence_end_date
                        print_display(f'{line_number()} [Microsoft Outlook] Recurrence end date set successfully: [{recurrence.PatternEndDate}]')
                    except OSError as os_error:
                        print_underline()
                        print_display(f'{line_number()} [Microsoft Outlook] OSError when setting recurrence end date: [{os_error}]')
                        print_overline()
            ms_outlook_appointment.Save()
            print_display(f'{line_number()} [Microsoft Outlook] INSERT SUCCESS: Event [{ms_outlook_appointment.Subject}] created with ID: [{trim_id(ms_outlook_appointment.EntryID)}]')
            sleep(1)
            return ms_outlook_appointment
        except Exception as exception:
            print_display(f'{line_number()} [Microsoft Outlook] INSERT ERROR: [{exception}]')
            import traceback
            traceback.print_exc()
            return None

    def update_instance_ms_outlook(self,
                                   ms_outlook_instance_id,
                                   ms_outlook_instance_body):
        try:
            print_display(f'{line_number()} [Microsoft Outlook] UPDATE <<==')
            ms_outlook_appointment = self.ms_outlook_data.ms_outlook_get_item(ms_outlook_instance_id)
            if 'Subject' in ms_outlook_instance_body:
                ms_outlook_subject = ms_outlook_instance_body['Subject']
                if ms_outlook_subject is None:
                    ms_outlook_subject = ''
                ms_outlook_appointment.Subject = str(ms_outlook_subject)
            if 'Body' in ms_outlook_instance_body:
                ms_outlook_body = ms_outlook_instance_body['Body']
                if ms_outlook_body is None:
                    ms_outlook_body = ''
                ms_outlook_appointment.Body = str(ms_outlook_body)
            if 'Location' in ms_outlook_instance_body:
                ms_outlook_location = ms_outlook_instance_body['Location']
                if ms_outlook_location is None:
                    ms_outlook_location = ''
                ms_outlook_appointment.Location = str(ms_outlook_location)

            '''
            if 'Organizer' in event_body:
                organizer = event_body['Organizer']
                if organizer is None:
                    organizer = ''
                appointment.Organizer = str(organizer)
            '''

            if 'StartUTC' in ms_outlook_instance_body:
                start_str = ms_outlook_instance_body['StartUTC']
                if isinstance(start_str,
                              str):
                    if 'T' in start_str:
                        start_dt = datetime.fromisoformat(start_str.replace('Z',
                                                                            '+00:00'))
                    else:
                        start_dt = datetime.fromisoformat(start_str)
                else:
                    start_dt = start_str
                ms_outlook_appointment.Start = start_dt
            if 'EndUTC' in ms_outlook_instance_body:
                end_str = ms_outlook_instance_body['EndUTC']
                if isinstance(end_str,
                              str):
                    if 'T' in end_str:
                        end_dt = datetime.fromisoformat(end_str.replace('Z',
                                                                        '+00:00'))
                    else:
                        end_dt = datetime.fromisoformat(end_str)
                else:
                    end_dt = end_str
                ms_outlook_appointment.End = end_dt
            if 'ReminderMinutesBeforeStart' in ms_outlook_instance_body:
                reminder_minutes = ms_outlook_instance_body['ReminderMinutesBeforeStart']
                if reminder_minutes is not None:
                    ms_outlook_appointment.ReminderSet = True
                    ms_outlook_appointment.ReminderMinutesBeforeStart = int(reminder_minutes)
            if 'Sensitivity' in ms_outlook_instance_body:
                sensitivity = ms_outlook_instance_body['Sensitivity']
                if sensitivity is not None:
                    ms_outlook_appointment.Sensitivity = int(sensitivity)
            if 'BusyStatus' in ms_outlook_instance_body:
                busy_status = ms_outlook_instance_body['BusyStatus']
                if busy_status is not None:
                    ms_outlook_appointment.BusyStatus = int(busy_status)
            '''
            if 'RequiredAttendees' in event_body or 'OptionalAttendees' in event_body:
                while appointment.Recipients.Count > 0:
                    appointment.Recipients.Remove(1)
                if 'RequiredAttendees' in event_body:
                    required_attendees = event_body['RequiredAttendees']
                    if required_attendees:
                        for email in str(required_attendees).split(';'):
                            email = email.strip()
                            if email:
                                recipient = appointment.Recipients.Add(email)
                                recipient.Type = 1  # 1 = Required
                
                if 'OptionalAttendees' in event_body:
                    optional_attendees = event_body['OptionalAttendees']
                    if optional_attendees:
                        for email in str(optional_attendees).split(';'):
                            email = email.strip()
                            if email:
                                recipient = appointment.Recipients.Add(email)
                                recipient.Type = 2  # 2 = Optional
                if appointment.Recipients.Count > 0:
                    appointment.Recipients.ResolveAll()
            '''
            ms_outlook_appointment.Save()
            print_display(f'{line_number()} [Microsoft Outlook] UPDATE SUCCESS: Event [{ms_outlook_appointment.Subject}] updated')
            return ms_outlook_appointment
        except Exception as exception:
            print_display(f'{line_number()} [Microsoft Outlook] UPDATE ERROR: {exception}')
            import traceback
            traceback.print_exc()
            return None

    def delete_instance_ms_outlook(self,
                                   ms_outlook_instance_id):
        try:
            print_display(f'{line_number()} DELETE <<==')
            ms_outlook_instance = self.ms_outlook_data.ms_outlook_get_item(ms_outlook_instance_id)
            ms_outlook_instance_subject = ms_outlook_instance.Subject
            ms_outlook_instance.Delete()
            print_display(f'{line_number()} [Microsoft Outlook] DELETE SUCCESS: Event [{ms_outlook_instance_subject}] deleted')
            return True
        except Exception as exception:
            print_display(f'{line_number()} [Microsoft Outlook] DELETE ERROR: [{exception}]')
            return False

    def delete_occurrence_ms_outlook(self,
                                     ms_outlook_instance_id,
                                     ms_outlook_instance_body):
        try:
            print_display(f'{line_number()} [Microsoft Outlook] DELETE INSIDE RECURRENCE <<==')
            appointment = self.ms_outlook_data.ms_outlook_get_item(ms_outlook_instance_id)
            if not appointment:
                return None
            if not appointment.IsRecurring:
                return None
            recurrence = appointment.GetRecurrencePattern()
            occurrence = recurrence.GetOccurrence(datetime.strptime(ms_outlook_instance_body,
                                                                    '%Y-%m-%d'))
            occurrence.Delete()
            return True
        except Exception as exception:
            print_display(f'{line_number()} [Microsoft Outlook] DELETE INSIDE RECURRENCE ERROR: {exception}')
            return False

    def delete_occurrence_by_g_calendar_master_and_start(self,
                                                         g_calendar_master_id: str,
                                                         start_utc: str):
        try:
            utc_dt = datetime.strptime(start_utc,
                                       '%Y-%m-%d-%H-%M-%S').replace(tzinfo=timezone.utc)
            print_display(f'{line_number()} UTC {utc_dt}')
        except ValueError as value_error:
            raise ValueError(f'[Microsoft Outlook] Invalid start_utc format: {value_error}')
        master = self.get_master_by_g_calendar_id(g_calendar_master_id)
        if not master:
            print_display(f'{line_number()} [Microsoft Outlook] Recurring master not found')
            raise ValueError('[Microsoft Outlook] Recurring master not found')
        if not master.IsRecurring:
            print_display(f'{line_number()} [Microsoft Outlook] Item is not recurring')
            raise ValueError('[Microsoft Outlook] Item is not recurring')
        offset = master.StartUTC.replace(tzinfo=None) - master.Start.replace(tzinfo=None)
        local_dt = utc_dt - offset
        print_display(f'{line_number()} [Microsoft Outlook] master [{local_dt}] [{master.Subject}]')
        recurrence = master.GetRecurrencePattern()
        print_display(f'{line_number()} [Microsoft Outlook] recurrence {recurrence}')
        try:
            occurrence = recurrence.GetOccurrence(local_dt)
            occurrence.Delete()
            return True
        except Exception as value_error:
            print_display(f'{line_number()} [Microsoft Outlook] Occurrence not found: {value_error}')
            raise ValueError(f'[Microsoft Outlook] Occurrence not found: {value_error}')

    def delete_occurrence_by_g_calendar_master_and_start_utc(self,
                                                             g_calendar_master_id: str,
                                                             start_utc: datetime):
        local_dt = start_utc.astimezone()
        master = self.get_master_by_g_calendar_id(g_calendar_master_id)
        if not master:
            raise ValueError('[Microsoft Outlook] Recurring master not found')
        if not master.IsRecurring:
            raise ValueError('[Microsoft Outlook] Item is not recurring')
        recurrence = master.GetRecurrencePattern()
        try:
            occurrence = recurrence.GetOccurrence(local_dt)
            occurrence.Delete()
            return True
        except Exception as exception:
            raise ValueError(f'[Microsoft Outlook] Occurrence not found: [{exception}]')

    def get_all_recurring_masters_ms_outlook(self):
        '''
        Fetches all recurring masters with no time restriction.
        Needed for recurrences that started before the 18-day window
        but still have active instances within it (e.g. Lunch since 2019).
        '''
        ms_outlook_all_instances = self.ms_outlook_data.ms_outlook_get_all_instances()
        ms_outlook_all_instances.IncludeRecurrences = False
        ms_outlook_all_instances.Sort('[Start]')
        ms_outlook_instances = dict()
        for ms_outlook_instance in ms_outlook_all_instances:
            try:
                if not getattr(ms_outlook_instance,
                               'IsRecurring',
                               False):
                    release_com_object_memory(ms_outlook_instance)
                    continue
                ms_outlook_properties = [attr for attr in dir(ms_outlook_instance) if not attr.startswith('_')]
                ms_outlook_instance_data = self.get_instance_data_ms_outlook(ms_outlook_instance,
                                                                             ms_outlook_properties)
                if 'EntryID' not in ms_outlook_instance_data:
                    release_com_object_memory(ms_outlook_instance)
                    continue
                ms_outlook_instances[ms_outlook_instance_data['EntryID']] = ms_outlook_instance_data
            except Exception as exception:
                print_display(f'{line_number()} [Microsoft Outlook] SKIPPING MASTER: [{exception}]')
            finally:
                release_com_object_memory(ms_outlook_instance)
        return ms_outlook_instances

    def get_recurring_masters_in_window_ms_outlook(self):
        time_now = datetime.now()
        time_begin = time_now - timedelta(days=constants.DAY_PAST)
        time_end = time_now + timedelta(days=constants.DAY_NEXT)
        ms_outlook_all_instances = self.ms_outlook_data.ms_outlook_get_all_instances()
        ms_outlook_all_instances.IncludeRecurrences = True
        ms_outlook_all_instances.Sort('[Start]')
        restriction_string = "([Start] >= '{}') AND [Start] <= '{}'"
        restriction = restriction_string.format(time_begin.strftime('%m/%d/%Y %H:%M %p'),
                                                time_end.strftime('%m/%d/%Y %H:%M %p'))
        ms_outlook_selected_instances = ms_outlook_all_instances.Restrict(restriction)
        ms_outlook_masters = dict()
        for ms_outlook_instance in ms_outlook_selected_instances:
            try:
                if not getattr(ms_outlook_instance,
                               'IsRecurring',
                               False):
                    release_com_object_memory(ms_outlook_instance)
                    continue
                ms_outlook_recurrence_pattern = ms_outlook_instance.GetRecurrencePattern()
                ms_outlook_master = ms_outlook_recurrence_pattern.Appointment
                if ms_outlook_master is None:
                    release_com_object_memory(ms_outlook_instance)
                    continue
                master_entry_id = ms_outlook_master.EntryID
                if master_entry_id in ms_outlook_masters:
                    release_com_object_memory(ms_outlook_recurrence_pattern)
                    release_com_object_memory(ms_outlook_instance)
                    continue
                ms_outlook_properties = [attr for attr in dir(ms_outlook_master) if not attr.startswith('_')]
                ms_outlook_instance_data = self.get_instance_data_ms_outlook(ms_outlook_master,
                                                                             ms_outlook_properties)
                if 'EntryID' in ms_outlook_instance_data:
                    ms_outlook_masters[master_entry_id] = ms_outlook_instance_data
            except Exception as exception:
                print_display(f'{line_number()} [Microsoft Outlook] SKIPPING MASTER: [{exception}]')
            finally:
                release_com_object_memory(ms_outlook_instance)
        return ms_outlook_masters
