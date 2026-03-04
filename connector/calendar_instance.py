import json
from datetime import datetime

import tzlocal
from dateutil import parser

from connector.g_calendar import GoogleCalendarConnector
from connector.ms_outlook import MicrosoftOutlookConnector
from system.tools import convert_to_local
from system.tools import remove_timezone_info
from system.tools import line_number
from system.tools import compare_rule
from system.tools import print_display
from system.tools import print_overline
from system.tools import print_underline


# FIX: Outlook DayOfWeekMask bit values => iCalendar BYDAY abbreviations
_OL_DAY_MASK_TO_BYDAY = {
        1 : 'SU',
        2 : 'MO',
        4 : 'TU',
        8 : 'WE',
        16: 'TH',
        32: 'FR',
        64: 'SA',
}

# FIX: iCalendar BYDAY abbreviations => Outlook DayOfWeekMask bit values
_BYDAY_TO_OL_DAY_MASK = {v: k for k, v in _OL_DAY_MASK_TO_BYDAY.items()}

# FIX: Outlook Instance values (1-5) => iCalendar BYSETPOS values.
# Outlook uses 5 to mean "last"; iCalendar uses -1.
_OL_INSTANCE_TO_BYSETPOS = {
        1: 1,
        2: 2,
        3: 3,
        4: 4,
        5: -1,
}

# FIX: iCalendar BYSETPOS values => Outlook Instance values
_BYSETPOS_TO_OL_INSTANCE = {v: k for k, v in _OL_INSTANCE_TO_BYSETPOS.items()}


def _day_mask_to_byday(mask: int) -> str:
    """Convert an Outlook DayOfWeekMask integer to a BYDAY string (e.g. 'MO,TU,WE,TH,FR')."""
    return ','.join(day for bit, day in _OL_DAY_MASK_TO_BYDAY.items() if mask & bit)


def _byday_to_day_mask(byday: str) -> int:
    """Convert a BYDAY string (e.g. 'MO,TU,WE,TH,FR') to an Outlook DayOfWeekMask integer."""
    total = 0
    for day in byday.split(','):
        total |= _BYDAY_TO_OL_DAY_MASK.get(day.strip().upper(), 0)
    return total
# FIX END

class CalendarInstance:
    def __init__(self):
        self.shared_uid = None
        self.shared_subject = None
        self.shared_description = None
        self.shared_location = None
        self.shared_start_date = None
        self.shared_end_date = None

        # self.shared_organizer = None

        # self.shared_attendees = []

        self.shared_recurrence = None
        self.shared_reminder_minutes = 15
        self.shared_visibility = 'public'
        self.shared_status = 'confirmed'
        self.ms_outlook_only = dict()
        self.g_calendar_only = dict()

    def __eq__(self,
               other):
        if not isinstance(other,
                          CalendarInstance):
            return False
        eq_subject = self.shared_subject == other.shared_subject
        print_display(f'{line_number()} [{self.shared_subject}]')
        print_display(f'{line_number()} [{other.shared_subject}]')
        eq_description = self.shared_description == other.shared_description
        if self.shared_description:
            print_display(f'{line_number()} [{self.shared_description.strip()}]')
        else:
            print_display(f'{line_number()} [-empty-]')
        if other.shared_description:
            print_display(f'{line_number()} [{other.shared_description.strip()}]')
        else:
            print_display(f'{line_number()} [-empty-]')
        eq_location = self.shared_location == other.shared_location
        print_display(f'{line_number()} [{self.shared_location}]')
        print_display(f'{line_number()} [{other.shared_location}]')
        eq_start_date = self.shared_start_date == other.shared_start_date
        print_display(f'{line_number()} [{self.shared_start_date}]')
        print_display(f'{line_number()} [{other.shared_start_date}]')
        eq_end_date = self.shared_end_date == other.shared_end_date
        print_display(f'{line_number()} [{self.shared_end_date}]')
        print_display(f'{line_number()} [{other.shared_end_date}]')

        '''     
        eq_organizer = self.shared_organizer == other.shared_organizer
        print_display(f'{line_number()} [{self.shared_organizer}]')
        print_display(f'{line_number()} [{other.shared_organizer}]')

        eq_attendees = self._normalized_attendees() == other._normalized_attendees()
        print_display(f'{line_number()} [{self.shared_attendees}]')
        print_display(f'{line_number()} [{other.shared_attendees}]')
        '''

        eq_recurrence = compare_rule(self.shared_recurrence,
                                     other.shared_recurrence)
        print_display(f'{line_number()} [{self.shared_recurrence}] [{eq_recurrence}]')
        print_display(f'{line_number()} [{other.shared_recurrence}] [{eq_recurrence}]')
        eq_reminder = self.shared_reminder_minutes == other.shared_reminder_minutes
        print_display(f'{line_number()} [{self.shared_reminder_minutes}]')
        print_display(f'{line_number()} [{other.shared_reminder_minutes}]')
        eq_visibility = self.shared_visibility == other.shared_visibility
        print_display(f'{line_number()} [{self.shared_visibility}]')
        print_display(f'{line_number()} [{other.shared_visibility}]')
        eq_status = self.shared_status == other.shared_status
        print_display(f'{line_number()} [{self.shared_status}]')
        print_display(f'{line_number()} [{other.shared_status}]')
        eq_result = eq_subject and eq_description and eq_location and eq_start_date and eq_end_date and eq_recurrence and eq_reminder and eq_visibility and eq_status
        print_underline()
        print_display(f'{line_number()} SUB[{eq_subject}] DES[{eq_description}] LOC[{eq_location}] STD[{eq_start_date}] END[{eq_end_date}] REC[{eq_recurrence}] REM[{eq_reminder}] VIS[{eq_visibility}] STA[{eq_status}] | [RES[{eq_result}]]')
        print_overline()
        return eq_result

    '''
    def _normalized_attendees(self):
        return sorted((attendees['email'],
                       attendees['optional']) for attendees in self.shared_attendees)
    '''

    def import_g_calendar(self,
                          g_calendar_event: dict):
        self.shared_uid = g_calendar_event.get('iCalUID')
        self.shared_subject = g_calendar_event.get('summary')
        self.shared_description = g_calendar_event.get('description')
        self.shared_location = g_calendar_event.get('location')
        self.shared_start_date = convert_to_local(g_calendar_event.get('start',
                                                                       {}).get('dateTime'))
        self.shared_end_date = convert_to_local(g_calendar_event.get('end',
                                                                     {}).get('dateTime'))
        '''
        self.shared_organizer = g_calendar_event.get('organizer',
                                                     {}).get('email')        
        self.shared_attendees = []       
        if 'attendees' in g_calendar_event:
            for att in g_calendar_event['attendees']:
                self.shared_attendees.append({
                        'email'   : att.get('email'),
                        'optional': att.get('optional',
                                            False)})
        '''
        if 'recurrence' in g_calendar_event:
            self.shared_recurrence = g_calendar_event['recurrence'][0]
        self.shared_reminder_minutes = (g_calendar_event.get('reminders',
                                                             {}).get('overrides',
                                                                     [{}])[0].get('minutes',
                                                                                  15))
        self.shared_visibility = g_calendar_event.get('visibility',
                                                      'public')
        self.shared_status = g_calendar_event.get('status',
                                                  'confirmed')
        self.g_calendar_only = {calendar_key: calendar_value for calendar_key, calendar_value in g_calendar_event.items() if calendar_key not in ['iCalUID',
                                                                                                                                                  'summary',
                                                                                                                                                  'description',
                                                                                                                                                  'location',
                                                                                                                                                  'start',
                                                                                                                                                  'end',

                                                                                                                                                  # 'organizer',
                                                                                                                                                  # 'attendees',

                                                                                                                                                  'recurrence',
                                                                                                                                                  'reminders',
                                                                                                                                                  'visibility',
                                                                                                                                                  'status']}

    def export_g_calendar(self) -> dict:
        local_timezone = tzlocal.get_localzone_name()
        g_calendar_event = {
                'iCalUID'    : self.shared_uid,
                'summary'    : self.shared_subject,
                'description': self.shared_description,
                'location'   : self.shared_location,
                'start'      : {
                        'dateTime': self.shared_start_date,
                        'timeZone': local_timezone},
                'end'        : {
                        'dateTime': self.shared_end_date,
                        'timeZone': local_timezone},
        '''        
                'organizer'  : {
                        'email': self.shared_organizer},
        '''        
                'reminders'  : {
                        'useDefault': False,
                        'overrides' : [{
                                'method' : 'popup',
                                'minutes': self.shared_reminder_minutes}]},
                'visibility' : self.shared_visibility,
                'status'     : self.shared_status}
        '''
        if self.shared_attendees:
            g_calendar_event['attendees'] = self.shared_attendees
        '''
        if self.shared_recurrence:
            g_calendar_event['recurrence'] = [self.shared_recurrence]
        g_calendar_event.update(self.g_calendar_only)
        return g_calendar_event

    def import_ms_outlook(self,
                          ms_outlook_event: dict):
        self.shared_uid = ms_outlook_event.get('GlobalAppointmentID')
        self.shared_subject = ms_outlook_event.get('Subject')
        self.shared_description = ms_outlook_event.get('Body')
        self.shared_location = ms_outlook_event.get('Location')
        self.shared_start_date = convert_to_local(ms_outlook_event.get('StartUTC'))
        self.shared_end_date = convert_to_local(ms_outlook_event.get('EndUTC'))

        '''
        self.shared_organizer = ms_outlook_event.get('Organizer')

        self.shared_attendees = []
        if ms_outlook_event.get('RequiredAttendees'):
            for required_attendee_email in ms_outlook_event['RequiredAttendees'].split(';'):
                self.shared_attendees.append({
                        'email'   : required_attendee_email.strip(),
                        'optional': False})
        if ms_outlook_event.get('OptionalAttendees'):
            for optional_attendee_email in ms_outlook_event['OptionalAttendees'].split(';'):
                self.shared_attendees.append({
                        'email'   : optional_attendee_email.strip(),
                        'optional': True})
        '''
        if ms_outlook_event.get('IsRecurring'):
            # FIX: added type 6 (olRecursYearNth) which was missing entirely
            # and would have fallen through to the 'DAILY' default
            frequency_map = {
                    0: 'DAILY',
                    1: 'WEEKLY',
                    2: 'MONTHLY',
                    3: 'MONTHLY',
                    5: 'YEARLY',
                    6: 'YEARLY'}
            recurrence_type = ms_outlook_event.get('recurrence_type', 0)
            recurrence_frequency = frequency_map.get(recurrence_type, 'DAILY')
            recurrence_interval = ms_outlook_event.get('recurrence_interval',
                                                       1)
            recurrence_until = ms_outlook_event.get('recurrence_end')
            recurrence_until_string = None
            if recurrence_until:
                try:
                    recurrence_until_date = datetime.strptime(recurrence_until,
                                                              '%Y-%m-%d')
                    recurrence_until_string = recurrence_until_date.strftime('%Y%m%dT235959Z')
                except ValueError:
                    recurrence_until_string = recurrence_until.replace('-',
                                                                       '') + 'T235959Z'
            recurrence_rule = f'RRULE:FREQ={recurrence_frequency};INTERVAL={recurrence_interval}'
            # FIX: for weekly recurrences, append BYDAY so Google Calendar
            # knows exactly which days of the week the event repeats on.
            # Without this, a Mon-Fri recurrence would be treated as
            # repeating only on the start date's weekday.            
            day_mask = ms_outlook_event.get('recurrence_day_of_week_mask', 0)
            recurrence_instance = ms_outlook_event.get('recurrence_instance', 0)
            recurrence_month_of_year = ms_outlook_event.get('recurrence_month_of_year', 0)

            if recurrence_frequency == 'WEEKLY':
                # FIX: weekly — append BYDAY so Google Calendar knows which
                # days the event repeats on (e.g. Mon-Fri).
                # Without this a Mon-Fri pattern appears as a single-day weekly.
                if day_mask:
                    byday = _day_mask_to_byday(day_mask)
                    if byday:
                        recurrence_rule += f';BYDAY={byday}'
            # FIX END
            elif recurrence_frequency == 'MONTHLY' and recurrence_type == 3:
                # FIX: monthly-nth (olRecursMonthNth) — "Nth weekday of month"
                # e.g. "3rd Thursday" needs BYDAY=TH and BYSETPOS=3.
                # Plain monthly (type 2) needs neither; the date is in start.
                if day_mask:
                    byday = _day_mask_to_byday(day_mask)
                    bysetpos = _OL_INSTANCE_TO_BYSETPOS.get(recurrence_instance, recurrence_instance)
                    if byday:
                        recurrence_rule += f';BYDAY={byday};BYSETPOS={bysetpos}'

            elif recurrence_frequency == 'YEARLY':
                # FIX: both yearly types need BYMONTH so Google Calendar
                # knows which month of the year the event falls in.
                if recurrence_month_of_year:
                    recurrence_rule += f';BYMONTH={recurrence_month_of_year}'
                if recurrence_type == 6:
                    # FIX: yearly-nth (olRecursYearNth) also needs BYDAY and
                    # BYSETPOS — e.g. "last Friday of November"
                    if day_mask:
                        byday = _day_mask_to_byday(day_mask)
                        bysetpos = _OL_INSTANCE_TO_BYSETPOS.get(recurrence_instance, recurrence_instance)
                        if byday:
                            recurrence_rule += f';BYDAY={byday};BYSETPOS={bysetpos}'

            if recurrence_until_string:
                recurrence_rule += f';UNTIL={recurrence_until_string}'
            self.shared_recurrence = recurrence_rule
        self.shared_reminder_minutes = ms_outlook_event.get('ReminderMinutesBeforeStart',
                                                            15)
        self.shared_visibility = 'public' if ms_outlook_event.get('Sensitivity',
                                                                  0) == 0 else 'private'
        self.shared_status = 'confirmed' if ms_outlook_event.get('BusyStatus',
                                                                 2) == 2 else 'tentative'
        self.ms_outlook_only = {item_key: item_value for item_key, item_value in ms_outlook_event.items() if item_key not in ['GlobalAppointmentID',
                                                                                                                              'Subject',
                                                                                                                              'Body',
                                                                                                                              'Location',
                                                                                                                              'StartUTC',
                                                                                                                              'EndUTC',

                                                                                                                              # 'Organizer',
                                                                                                                              # 'RequiredAttendees',
                                                                                                                              # 'OptionalAttendees',

                                                                                                                              'IsRecurring',
                                                                                                                              'recurrence_type',
                                                                                                                              'recurrence_interval',
                                                                                                                              'recurrence_day_of_week_mask',
                                                                                                                              'recurrence_instance',
                                                                                                                              'recurrence_month_of_year',
                                                                                                                              'recurrence_end',
                                                                                                                              'ReminderMinutesBeforeStart',
                                                                                                                              'Sensitivity',
                                                                                                                              'BusyStatus']}

    def export_ms_outlook(self) -> dict:
        ms_outlook_export_event = {
                'GlobalAppointmentID'       : self.shared_uid,
                'Subject'                   : self.shared_subject,
                'Body'                      : self.shared_description,
                'Location'                  : self.shared_location,
                'StartUTC'                  : remove_timezone_info(self.shared_start_date),
                'EndUTC'                    : remove_timezone_info(self.shared_end_date),

                # 'Organizer'                 : self.shared_organizer,

                'ReminderMinutesBeforeStart': self.shared_reminder_minutes,
                'Sensitivity'               : 0 if self.shared_visibility == 'public' else 2,
                'BusyStatus'                : 2 if self.shared_status == 'confirmed' else 1,
                'IsRecurring'               : bool(self.shared_recurrence)}
        '''
        required_attendees = [required_attendee_email['email'] for required_attendee_email in self.shared_attendees if not required_attendee_email.get('optional')]
        optional_attendees = [optional_attendee_email['email'] for optional_attendee_email in self.shared_attendees if optional_attendee_email.get('optional')]
        if required_attendees:
            ms_outlook_export_event['RequiredAttendees'] = ';'.join(required_attendees)
        if optional_attendees:
            ms_outlook_export_event['OptionalAttendees'] = ';'.join(optional_attendees)
        '''
        if self.shared_recurrence:
            recurrence_rule = self.shared_recurrence

            # FIX: determine recurrence_type more precisely now that we
            # distinguish type 3 (monthly-nth) and type 6 (yearly-nth)
            has_byday = 'BYDAY=' in recurrence_rule
            has_bysetpos = 'BYSETPOS=' in recurrence_rule
            has_bymonth = 'BYMONTH=' in recurrence_rule

            if 'FREQ=DAILY' in recurrence_rule:
                ms_outlook_export_event['recurrence_type'] = 0
            elif 'FREQ=WEEKLY' in recurrence_rule:
                ms_outlook_export_event['recurrence_type'] = 1
            elif 'FREQ=MONTHLY' in recurrence_rule:
                # type 3 = monthly-nth (has BYDAY+BYSETPOS); type 2 = plain monthly
                ms_outlook_export_event['recurrence_type'] = 3 if (has_byday and has_bysetpos) else 2
            elif 'FREQ=YEARLY' in recurrence_rule:
                # type 6 = yearly-nth (has BYDAY+BYSETPOS); type 5 = plain yearly
                ms_outlook_export_event['recurrence_type'] = 6 if (has_byday and has_bysetpos) else 5
            else:
                ms_outlook_export_event['recurrence_type'] = 0

            if 'INTERVAL=' in recurrence_rule:
                ms_outlook_export_event['recurrence_interval'] = int(recurrence_rule.split('INTERVAL=')[1].split(';')[0])
            if 'UNTIL=' in recurrence_rule:
                recurrence_until = recurrence_rule.split('UNTIL=')[1].split(';')[0]
                recurrence_until_date = parser.parse(recurrence_until)
                recurrence_until_date = recurrence_until_date.date()
                ms_outlook_export_event['recurrence_end'] = recurrence_until_date.strftime('%Y-%m-%d')
            # FIX: parse BYDAY from the RRULE and convert back to an Outlook
            # DayOfWeekMask so that weekly patterns written to Outlook
            # preserve the correct days of the week.
            # FIX: parse BYDAY back to DayOfWeekMask for weekly (type 1),
            # monthly-nth (type 3), and yearly-nth (type 6)
            if has_byday:
                byday_value = recurrence_rule.split('BYDAY=')[1].split(';')[0]
                day_mask = _byday_to_day_mask(byday_value)
                if day_mask:
                    ms_outlook_export_event['recurrence_day_of_week_mask'] = day_mask
            # FIX END
            # FIX: parse BYSETPOS back to Outlook Instance for monthly-nth
            # (type 3) and yearly-nth (type 6); iCalendar -1 maps to
            # Outlook Instance 5 meaning "last"
            if has_bysetpos:
                bysetpos_raw = int(recurrence_rule.split('BYSETPOS=')[1].split(';')[0])
                ol_instance = _BYSETPOS_TO_OL_INSTANCE.get(bysetpos_raw, bysetpos_raw)
                ms_outlook_export_event['recurrence_instance'] = ol_instance

            # FIX: parse BYMONTH back to MonthOfYear for yearly (type 5)
            # and yearly-nth (type 6)
            if has_bymonth:
                ms_outlook_export_event['recurrence_month_of_year'] = int(recurrence_rule.split('BYMONTH=')[1].split(';')[0])

        ms_outlook_export_event.update(self.ms_outlook_only)
        return ms_outlook_export_event

    def to_dict(self) -> dict:
        return {
                'shared_uid'             : self.shared_uid,
                'shared_subject'         : self.shared_subject,
                'shared_description'     : self.shared_description,
                'shared_location'        : self.shared_location,
                'shared_start_date'      : self.shared_start_date,
                'shared_end_date'        : self.shared_end_date,

                # 'shared_organizer'       : self.shared_organizer,
                # 'shared_attendees'       : self.shared_attendees,

                'shared_recurrence'      : self.shared_recurrence,
                'shared_reminder_minutes': self.shared_reminder_minutes,
                'shared_visibility'      : self.shared_visibility,
                'shared_status'          : self.shared_status,
                'ms_outlook_only'        : self.ms_outlook_only,
                'g_calendar_only'        : self.g_calendar_only}

    def __repr__(self) -> str:
        return json.dumps(self.to_dict(),
                          indent=4,
                          default=str)

    __str__ = __repr__


if __name__ == '__main__':
    kind = 'm'
    if kind == 'e':
        g_calendar_connection = GoogleCalendarConnector()
        g_calendar_events = g_calendar_connection.get_all_sub_instances_g_calendar()
        for event_id, event_data in g_calendar_events.items():
            calendar_event = CalendarInstance()
            calendar_event.import_g_calendar(event_data)
            event_result = calendar_event.export_g_calendar()
            print(event_result)
            break
    if kind == 'm':
        ms_outlook_connection = MicrosoftOutlookConnector()
        ms_outlook_events = ms_outlook_connection.get_all_instances_ms_outlook()
        for event_id, event_data in ms_outlook_events.items():
            calendar_event = CalendarInstance()
            calendar_event.import_ms_outlook(event_data)
            event_result = calendar_event.export_ms_outlook()
            print(event_result)
            break
