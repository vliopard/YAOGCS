import gc
from datetime import datetime
from datetime import timedelta
from datetime import timezone
from time import sleep

import pywintypes
import win32com.client
from tqdm import tqdm

from utils.utils import _release
from utils.utils import com_object_to_dictionary
from utils.utils import line_number
from utils.utils import print_display
from utils.utils import print_overline
from utils.utils import print_underline


class MicrosoftOutlookHelper:
    def __init__(self):
        self.ms_outlook_client = win32com.client.Dispatch('Outlook.Application')
        self.ms_outlook_namespace = self.ms_outlook_client.GetNamespace('MAPI')
        self.ms_outlook_calendar = self.ms_outlook_namespace.GetDefaultFolder(9)

    def ms_outlook_items(self):
        return self.ms_outlook_calendar.Items

    def ms_outlook_create(self):
        return self.ms_outlook_client.CreateItem(1)

    def ms_outlook_get_item(self,
                            entry_id):
        return self.ms_outlook_namespace.GetItemFromID(entry_id,
                                                       self.ms_outlook_calendar.StoreID)


class MicrosoftOutlookConnector:
    def __init__(self):
        self.ms_outlook_deleted_exceptions_map = None
        self.ms_outlook_data = MicrosoftOutlookHelper()

    def get_ms_outlook_item(self,
                            entry_id):
        return com_object_to_dictionary(self.ms_outlook_data.ms_outlook_get_item(entry_id))

    def get_ms_outlook_events(self):
        time_now = datetime.now()
        time_begin = time_now - timedelta(days=18)
        time_end = time_now + timedelta(days=180)
        items = self.ms_outlook_data.ms_outlook_items()
        items.IncludeRecurrences = False
        items.Sort('[Start]')
        restriction_string = "([Start] >= '{}' OR [End] >= '{}') AND [End] <= '{}'"
        restriction = restriction_string.format(time_begin.strftime('%m/%d/%Y %H:%M %p'),
                                                time_begin.strftime('%m/%d/%Y %H:%M %p'),
                                                time_end.strftime('%m/%d/%Y %H:%M %p'))
        selected_items = items.Restrict(restriction)
        ms_outlook_events = dict()
        for item_index, item in enumerate(tqdm(selected_items,
                                               desc='Processing Outlook items',
                                               total=len(selected_items))):
            local_event_data = dict()
            try:
                item_properties = [item_attributes for item_attributes in dir(item) if not item_attributes.startswith('_')]
            except Exception as exception:
                print_display(f'{line_number()} SKIPPING ITEM: [{item_index}] — dir() - FAILED: [{exception}]')
                _release(item)
                continue
            for item_property in item_properties:
                if item_property == 'GetInspector':
                    continue
                try:
                    item_value = getattr(item,
                                         item_property)
                    if callable(item_value):
                        continue
                    if item_property == 'Recipients':
                        attendees = []
                        for recipient in item.Recipients:
                            attendees.append({
                                    'email': recipient.Address})
                            _release(recipient)
                        local_event_data[item_property] = attendees
                    elif item_property in ('Start',
                                           'End'):
                        if hasattr(item_value,
                                   'Format'):
                            local_event_data[item_property] = item_value.Format('%Y-%m-%dT%H:%M:%S')
                        else:
                            local_event_data[item_property] = item_value
                    elif item_property == 'Body':
                        local_event_data[item_property] = item_value[:200] if item_value else ''
                    elif item_property == 'IsRecurring' and item_value:
                        local_event_data[item_property] = item_value
                        recurrence_pattern = item.GetRecurrencePattern()
                        local_event_data['recurrence_type'] = recurrence_pattern.RecurrenceType
                        local_event_data['recurrence_interval'] = recurrence_pattern.Interval
                        local_event_data['recurrence_end'] = recurrence_pattern.PatternEndDate.Format('%Y-%m-%d')
                        _release(recurrence_pattern)
                    else:
                        local_event_data[item_property] = item_value
                except pywintypes.com_error as com_error_type:
                    print_display(f'{line_number()} COM ERROR for property [{item_property}] of item [{item.Subject}]: [{com_error_type}]')
                    continue
            if 'EntryID' not in local_event_data:
                print_display(f'{line_number()} SKIPPING ITEM: [{item_index}] — missing EntryID')
                _release(item)
                continue
            ms_outlook_events[local_event_data['EntryID']] = local_event_data
            _release(item)
            if item_index % 50 == 0:
                gc.collect()
        gc.collect()
        return ms_outlook_events

    def ms_outlook_insert(self,
                          event_body):
        try:
            print_display(f'{line_number()} INSERT <<==')
            appointment = self.ms_outlook_data.ms_outlook_create()
            subject = event_body.get('Subject',
                                     '')
            if subject is None:
                subject = ''
            appointment.Subject = str(subject)
            body = event_body.get('Body',
                                  '')
            if body is None:
                body = ''
            appointment.Body = str(body)
            organizer = event_body.get('Organizer',
                                       '')
            if organizer is None:
                organizer = ''
            appointment.Organizer = str(organizer)
            location = event_body.get('Location',
                                      '')
            if location is None:
                location = ''
            appointment.Location = str(location)
            start_str = event_body.get('StartUTC',
                                       '')
            end_str = event_body.get('EndUTC',
                                     '')
            if start_str:
                if isinstance(start_str,
                              str):
                    if 'T' in start_str:
                        start_dt = datetime.fromisoformat(start_str.replace('Z',
                                                                            '+00:00'))
                    else:
                        start_dt = datetime.fromisoformat(start_str)
                else:
                    start_dt = start_str
                appointment.Start = start_dt
            if end_str:
                if isinstance(end_str,
                              str):
                    if 'T' in end_str:
                        recurrence_end_date = datetime.fromisoformat(end_str.replace('Z',
                                                                                     '+00:00'))
                    else:
                        recurrence_end_date = datetime.fromisoformat(end_str)
                else:
                    recurrence_end_date = end_str
                appointment.End = recurrence_end_date
            reminder_minutes = event_body.get('ReminderMinutesBeforeStart',
                                              15)
            if reminder_minutes is not None:
                appointment.ReminderSet = True
                appointment.ReminderMinutesBeforeStart = int(reminder_minutes)
            sensitivity = event_body.get('Sensitivity',
                                         0)
            if sensitivity is not None:
                appointment.Sensitivity = int(sensitivity)
            busy_status = event_body.get('BusyStatus',
                                         2)
            if busy_status is not None:
                appointment.BusyStatus = int(busy_status)
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
            if event_body.get('IsRecurring'):
                recurrence = appointment.GetRecurrencePattern()
                recurrence.RecurrenceType = int(event_body.get('recurrence_type',
                                                               0))
                recurrence.Interval = int(event_body.get('recurrence_interval',
                                                         1))
                recurrence.PatternStartDate = appointment.Start
                recurrence_end = event_body.get('recurrence_end')
                if recurrence_end:
                    recurrence_end_date = datetime.strptime(recurrence_end,
                                                            '%Y-%m-%d')
                    recurrence_end_date = recurrence_end_date.replace(hour=appointment.Start.hour,
                                                                      minute=appointment.Start.minute,
                                                                      second=appointment.Start.second)
                    print_display(f'{line_number()} Setting recurrence end date: [{recurrence_end_date}] [{appointment.Subject}]')
                    try:
                        recurrence.PatternEndDate = recurrence_end_date
                        print_display(f'{line_number()} Recurrence end date set successfully: [{recurrence.PatternEndDate}]')
                    except OSError as os_error:
                        print_underline()
                        print_display(f'{line_number()} OSError when setting recurrence end date: [{os_error}]')
                        print_overline()
            appointment.Save()
            print_display(f'{line_number()} [Microsoft Outlook] INSERT SUCCESS: Event [{appointment.Subject}] created with ID: [{appointment.EntryID[-10:]}]')
            sleep(1)
            return appointment
        except Exception as exception:
            print_display(f'{line_number()} [Microsoft Outlook] INSERT ERROR: [{exception}]')
            import traceback
            traceback.print_exc()
            return None

    def ms_outlook_update(self,
                          entry_id,
                          event_body):
        try:
            print_display(f'{line_number()} UPDATE <<==')
            appointment = self.ms_outlook_data.ms_outlook_get_item(entry_id)
            if 'Subject' in event_body:
                subject = event_body['Subject']
                if subject is None:
                    subject = ''
                appointment.Subject = str(subject)
            if 'Body' in event_body:
                body = event_body['Body']
                if body is None:
                    body = ''
                appointment.Body = str(body)
            if 'Location' in event_body:
                location = event_body['Location']
                if location is None:
                    location = ''
                appointment.Location = str(location)
            if 'Organizer' in event_body:
                organizer = event_body['Organizer']
                if organizer is None:
                    organizer = ''
                appointment.Organizer = str(organizer)
            if 'StartUTC' in event_body:
                start_str = event_body['StartUTC']
                if isinstance(start_str,
                              str):
                    if 'T' in start_str:
                        start_dt = datetime.fromisoformat(start_str.replace('Z',
                                                                            '+00:00'))
                    else:
                        start_dt = datetime.fromisoformat(start_str)
                else:
                    start_dt = start_str
                appointment.Start = start_dt
            if 'EndUTC' in event_body:
                end_str = event_body['EndUTC']
                if isinstance(end_str,
                              str):
                    if 'T' in end_str:
                        end_dt = datetime.fromisoformat(end_str.replace('Z',
                                                                        '+00:00'))
                    else:
                        end_dt = datetime.fromisoformat(end_str)
                else:
                    end_dt = end_str
                appointment.End = end_dt
            if 'ReminderMinutesBeforeStart' in event_body:
                reminder_minutes = event_body['ReminderMinutesBeforeStart']
                if reminder_minutes is not None:
                    appointment.ReminderSet = True
                    appointment.ReminderMinutesBeforeStart = int(reminder_minutes)
            if 'Sensitivity' in event_body:
                sensitivity = event_body['Sensitivity']
                if sensitivity is not None:
                    appointment.Sensitivity = int(sensitivity)
            if 'BusyStatus' in event_body:
                busy_status = event_body['BusyStatus']
                if busy_status is not None:
                    appointment.BusyStatus = int(busy_status)
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
            appointment.Save()
            print_display(f'{line_number()} [Microsoft Outlook] UPDATE SUCCESS: Event [{appointment.Subject}] updated')
            return appointment
        except Exception as exception:
            print_display(f'{line_number()} [Microsoft Outlook] UPDATE ERROR: {exception}')
            import traceback
            traceback.print_exc()
            return None

    def ms_outlook_delete_event(self,
                                entry_id):
        try:
            print_display(f'{line_number()} DELETE <<==')
            ms_outlook_event = self.ms_outlook_data.ms_outlook_get_item(entry_id)
            ms_outlook_event_subject = ms_outlook_event.Subject
            ms_outlook_event.Delete()
            print_display(f'{line_number()} [Microsoft Outlook] DELETE SUCCESS: Event [{ms_outlook_event_subject}] deleted')
            return True
        except Exception as exception:
            print_display(f'{line_number()} [Microsoft Outlook] DELETE ERROR: [{exception}]')
            return False

    def ms_outlook_get_instance(self,
                                entry_id,
                                start_date):
        try:
            appointment = self.ms_outlook_data.ms_outlook_get_item(entry_id)
            if not appointment:
                return None
            if not appointment.IsRecurring:
                return None
            recurrence = appointment.GetRecurrencePattern()
            occurrence = recurrence.GetOccurrence(datetime.strptime(start_date,
                                                                    '%Y-%m-%d'))
            return com_object_to_dictionary(occurrence)
        except pywintypes.com_error as com_error:
            print_display(f'{line_number()} COM ERROR: {com_error}')
            return None

    def find_matching_recurrent_master(self,
                                       subject: str,
                                       start_dt_local: datetime):
        ms_outlook_helper = MicrosoftOutlookHelper()
        items = ms_outlook_helper.ms_outlook_items()
        items.IncludeRecurrences = True
        items.Sort('[Start]')
        start_window = (start_dt_local - timedelta(days=30)).strftime('%m/%d/%Y %H:%M %p')
        end_window = (start_dt_local + timedelta(days=30)).strftime('%m/%d/%Y %H:%M %p')
        restriction = f"[Start] >= '{start_window}' AND [Start] <= '{end_window}'"
        restricted = items.Restrict(restriction)
        for item in restricted:
            if item.IsRecurring and item.Subject == subject:
                return item
        return None

    def get_master_by_g_calendar_id(self,
                                    g_calendar_master_id: str):
        ms_outlook_helper = MicrosoftOutlookHelper()
        items = ms_outlook_helper.ms_outlook_items()
        items.IncludeRecurrences = False
        time_now = datetime.now()
        time_begin = (time_now - timedelta(days=18)).strftime('%m/%d/%Y %H:%M %p')
        time_end = (time_now + timedelta(days=180)).strftime('%m/%d/%Y %H:%M %p')
        restriction = f"[Start] >= '{time_begin}' AND [Start] <= '{time_end}'"
        restricted_items = items.Restrict(restriction)
        for item in restricted_items:
            if not item.IsRecurring:
                continue
            try:
                print_display(f'{line_number()} Checking item [{item.Subject}] (IsRecurring: [{item.IsRecurring}])')
                prop = item.UserProperties.Find('GCalendarMasterID')
                v_prop = prop.Value if prop else 'NOT SET'
                print_display(f'{line_number()} GCalendarMasterID for item [{item.Subject}]: [{v_prop[-10:]}]')
                print_display(f'{line_number()} GCalendarMasterID for item [{item.Subject}]: [{g_calendar_master_id[-10:]}]')
                if prop and prop.Value == g_calendar_master_id:
                    print_display(f'{line_number()} Found master [{item.Subject}] for GCalendarMasterID [{g_calendar_master_id[-10:]}]')
                    return item
            except Exception as exception:
                print_display(f'{line_number()} Error checking GCalendarMasterID for item [{item.Subject}]: [{exception}]')
                continue
        print_display(f'{line_number()} Master not found for GCalendarMasterID [{g_calendar_master_id[-10:]}]')
        return None

    def get_occurrence_by_g_calendar_master_and_start(self,
                                                      g_calendar_master_id: str,
                                                      start_utc: str):
        try:
            utc_dt = datetime.strptime(start_utc,
                                       '%Y-%m-%d-%H-%M-%S').replace(tzinfo=timezone.utc)
            print_display(f'{line_number()} UTC [{utc_dt}]')
        except ValueError as value_error:
            raise ValueError(f'Invalid start_utc format: [{value_error}]')
        master = self.get_master_by_g_calendar_id(g_calendar_master_id)
        if not master:
            print_display(f'{line_number()} MASTER ID not found: [{g_calendar_master_id[-10:]}]')
            return None
        if not master.IsRecurring:
            print_display(f'{line_number()} Item is not recurring')
            raise ValueError('Item is not recurring')
        offset = master.StartUTC.replace(tzinfo=None) - master.Start.replace(tzinfo=None)
        local_dt = utc_dt - offset
        print_display(f'{line_number()} master [{local_dt}] [{master.Subject}]')
        recurrence = master.GetRecurrencePattern()
        print_display(f'{line_number()} recurrence [{recurrence}]')
        try:
            occurrence = recurrence.GetOccurrence(local_dt)
            return occurrence
        except Exception as value_error:
            print_display(f'{line_number()} Occurrence not found: [{value_error}]')
        return None

    def delete_occurrence_by_g_calendar_master_and_start(self,
                                                         g_calendar_master_id: str,
                                                         start_utc: str):
        try:
            utc_dt = datetime.strptime(start_utc,
                                       '%Y-%m-%d-%H-%M-%S').replace(tzinfo=timezone.utc)
            print_display(f'{line_number()} UTC {utc_dt}')
        except ValueError as value_error:
            raise ValueError(f'Invalid start_utc format: {value_error}')
        master = self.get_master_by_g_calendar_id(g_calendar_master_id)
        if not master:
            print_display(f'{line_number()} Recurring master not found')
            raise ValueError('Recurring master not found')
        if not master.IsRecurring:
            print_display(f'{line_number()} Item is not recurring')
            raise ValueError('Item is not recurring')
        offset = master.StartUTC.replace(tzinfo=None) - master.Start.replace(tzinfo=None)
        local_dt = utc_dt - offset
        print_display(f'{line_number()} master [{local_dt}] [{master.Subject}]')
        recurrence = master.GetRecurrencePattern()
        print_display(f'{line_number()} recurrence {recurrence}')
        try:
            occurrence = recurrence.GetOccurrence(local_dt)
            occurrence.Delete()
            return True
        except Exception as value_error:
            print_display(f'{line_number()} Occurrence not found: {value_error}')
            raise ValueError(f'Occurrence not found: {value_error}')

    def delete_occurrence_by_g_calendar_master_and_start_utc(self,
                                                             g_calendar_master_id: str,
                                                             start_utc: datetime):
        local_dt = start_utc.astimezone()
        master = self.get_master_by_g_calendar_id(g_calendar_master_id)
        if not master:
            raise ValueError('Recurring master not found')
        if not master.IsRecurring:
            raise ValueError('Item is not recurring')
        recurrence = master.GetRecurrencePattern()
        try:
            occurrence = recurrence.GetOccurrence(local_dt)
            occurrence.Delete()
            return True
        except Exception as exception:
            raise ValueError(f'Occurrence not found: [{exception}]')

    def ms_outlook_delete_inside_recurrence(self,
                                            entry_id,
                                            start_date):
        try:
            print_display(f'{line_number()} DELETE INSIDE RECURRENCE<<==')
            appointment = self.ms_outlook_data.ms_outlook_get_item(entry_id)
            if not appointment:
                return None
            if not appointment.IsRecurring:
                return None
            recurrence = appointment.GetRecurrencePattern()
            occurrence = recurrence.GetOccurrence(datetime.strptime(start_date,
                                                                    '%Y-%m-%d'))
            occurrence.Delete()
            return True
        except Exception as exception:
            print_display(f'{line_number()} [Microsoft Outlook] DELETE INSIDE RECURRENCE ERROR: {exception}')
            return False

    def set_recurrence_id(self,
                          ms_outlook_master_id,
                          g_calendar_master_id):
        try:
            print_display(f'{line_number()} 01) Setting GCalendarMasterID [{g_calendar_master_id[-10:]}] for [Microsoft Outlook] master [{ms_outlook_master_id[-10:]}]')
            appointment = self.ms_outlook_data.ms_outlook_get_item(ms_outlook_master_id)
            print_display(f'{line_number()} 02) Appointment retrieved for master ID [{ms_outlook_master_id[-10:]}]: [{appointment.Subject}]')
            if appointment:
                print_display(f'{line_number()} 03) Setting GCalendarMasterID for appointment [{appointment.Subject}]')
                appointment.UserProperties.Add('GCalendarMasterID',
                                               1)
                print_display(f'{line_number()} 04) GCalendarMasterID property added for appointment [{appointment.Subject}]')
                appointment.UserProperties['GCalendarMasterID'].Value = g_calendar_master_id
                print_display(f'{line_number()} 05) GCalendarMasterID set successfully for [Microsoft Outlook] master [{ms_outlook_master_id[-10:]}]')
                appointment.Save()
        except Exception as exception:
            print_display(f'{line_number()} 06) Error setting GCalendarMasterID [{ms_outlook_master_id[-10:]}][{g_calendar_master_id[-10:]}] for [Microsoft Outlook] item: [{exception}]')

    def get_recurrence_instances(self,
                                 master_id):
        master = self.ms_outlook_data.ms_outlook_get_item(master_id)
        if not master.IsRecurring:
            raise ValueError('Provided ID is not a recurring appointment')
        pattern = master.GetRecurrencePattern()
        instances = list()
        start = master.Start
        end = master.End
        current = start
        while current.date() <= pattern.PatternEndDate.date():
            try:
                occ = pattern.GetOccurrence(current)
                instances.append(com_object_to_dictionary(occ))
            except pywintypes.com_error:
                pass
            current = current + (end - start)
        _release(pattern)
        _release(master)
        gc.collect()
        return instances
