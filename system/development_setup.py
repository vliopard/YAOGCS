from datetime import date
from datetime import datetime
from datetime import timedelta

from connector.event_mapping import EventMapping
from connector.g_calendar import GoogleCalendarConnector
from connector.ms_outlook import MicrosoftOutlookHelper
from system.tools import line_number
from system.tools import print_display
from system.tools import utc_to_outlook_local


class MessageSetup:
    def __init__(self):
        self.g_calendar_connection = GoogleCalendarConnector()

    def create_g_calendar_single_event(self,
                                       start_time: str,
                                       end_time: str,
                                       start_day: str,
                                       title: str,
                                       location: str,
                                       body: str,

                                       # attendees: list

                                       ):
        date_start = f'{start_day}T{start_time}:00'
        date_end = f'{start_day}T{end_time}:00'
        event = {
                'summary'    : title,
                'location'   : location,
                'description': body,
                'start'      : {
                        'dateTime': date_start,
                        'timeZone': 'America/Sao_Paulo'},
                'end'        : {
                        'dateTime': date_end,
                        'timeZone': 'America/Sao_Paulo'}, }
        """
        'attendees'  : [{
                'email': email} for email in attendees],
        """

        event = self.g_calendar_connection.g_calendar_insert_instance(event)
        htm_link = event.get('htmlLink')
        print_display(f'{line_number()} Single Google Calendar event created: [{htm_link}]')

    def create_g_calendar_daily_recurrence(self,
                                           start_time: str,
                                           end_time: str,
                                           start_day: str,
                                           end_day: str,
                                           title: str,
                                           location: str,
                                           body: str,

                                           # attendees: list,

                                           ):
        date_start = f'{start_day}T{start_time}:00'
        date_end = f'{start_day}T{end_time}:00'
        the_end = end_day.replace('-',
                                  '')
        recurrence_rule = f'RRULE:FREQ=DAILY;UNTIL={the_end}T235959Z'
        event = {
                'summary'    : title,
                'location'   : location,
                'description': body,
                'start'      : {
                        'dateTime': date_start,
                        'timeZone': 'America/Sao_Paulo'},
                'end'        : {
                        'dateTime': date_end,
                        'timeZone': 'America/Sao_Paulo'},
                'recurrence' : [recurrence_rule], }
        """
        'attendees'  : [{
                'email': email} for email in attendees],
        """
        event = self.g_calendar_connection.g_calendar_insert_instance(event)
        htm_link = event.get('htmlLink')
        print_display(f'{line_number()} Recurring Google Calendar event created: [{htm_link}]')

    def create_ms_outlook_single_event(self,
                                       start_time: str,
                                       end_time: str,
                                       start_day: str,
                                       title: str,
                                       location: str,
                                       body: str,

                                       # attendees: list

                                       ):
        ms_outlook_helper = MicrosoftOutlookHelper()
        appointment = ms_outlook_helper.ms_outlook_create_item()
        start_dt = datetime.strptime(f'{start_day} {start_time}',
                                     '%Y-%m-%d %H:%M')
        end_dt = datetime.strptime(f'{start_day} {end_time}',
                                   '%Y-%m-%d %H:%M')
        appointment.Subject = title
        appointment.Location = location
        appointment.Body = body
        appointment.Start = utc_to_outlook_local(start_dt)
        appointment.End = utc_to_outlook_local(end_dt)

        '''
        for email in attendees:
            recipient = appointment.Recipients.Add(email)
            recipient.Type = 1  # Required attendee
        appointment.Recipients.ResolveAll()
        '''

        appointment.Save()
        # appointment.Send()
        print_display(f'{line_number()} Single Outlook calendar event created successfully.')

    def create_ms_outlook_daily_recurrence(self,
                                           start_time: str,
                                           end_time: str,
                                           start_day: str,
                                           end_day: str,
                                           title: str,
                                           location: str,
                                           body: str,

                                           # attendees: list

                                           ):
        ms_outlook_helper = MicrosoftOutlookHelper()
        appointment = ms_outlook_helper.ms_outlook_create_item()
        start_dt = datetime.strptime(f'{start_day} {start_time}',
                                     '%Y-%m-%d %H:%M')
        end_dt = datetime.strptime(f'{start_day} {end_time}',
                                   '%Y-%m-%d %H:%M')
        appointment.Subject = title
        appointment.Location = location
        appointment.Body = body
        appointment.Start = utc_to_outlook_local(start_dt)
        appointment.End = utc_to_outlook_local(end_dt)

        """
        for email in attendees:
            recipient = appointment.Recipients.Add(email)
            recipient.Type = 1  # 1 = Required
        appointment.Recipients.ResolveAll()
        """

        recurrence = appointment.GetRecurrencePattern()
        recurrence.RecurrenceType = 0  # 0 = olRecursDaily
        recurrence.PatternStartDate = datetime.strptime(start_day,
                                                        '%Y-%m-%d')
        recurrence.PatternEndDate = datetime.strptime(end_day,
                                                      '%Y-%m-%d')

        appointment.Save()
        # appointment.Send()
        print_display(f'{line_number()} Recurring Outlook calendar event created successfully.')

    # NEW SETUP START
    def create_ms_outlook_weekly_recurrence(self,
                                            start_time: str,
                                            end_time: str,
                                            start_day: str,
                                            end_day: str,
                                            title: str,
                                            location: str,
                                            body: str,
                                            days_of_week: list,
                                            interval: int = 1,

                                            # attendees: list

                                            ):
        """
        Creates a weekly recurring Outlook appointment.

        days_of_week: list of day name strings, e.g. ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
                      Accepted values: 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'
        interval:     how many weeks between occurrences, default 1
        """
        # Outlook DayOfWeekMask bit values
        day_mask_map = {
                'Sunday'   : 1,
                'Monday'   : 2,
                'Tuesday'  : 4,
                'Wednesday': 8,
                'Thursday' : 16,
                'Friday'   : 32,
                'Saturday' : 64,
        }
        day_of_week_mask = 0
        for day in days_of_week:
            day_of_week_mask |= day_mask_map.get(day,
                                                 0)
        ms_outlook_helper = MicrosoftOutlookHelper()
        appointment = ms_outlook_helper.ms_outlook_create_item()
        start_dt = datetime.strptime(f'{start_day} {start_time}',
                                     '%Y-%m-%d %H:%M')
        end_dt = datetime.strptime(f'{start_day} {end_time}',
                                   '%Y-%m-%d %H:%M')
        appointment.Subject = title
        appointment.Location = location
        appointment.Body = body
        appointment.Start = utc_to_outlook_local(start_dt)
        appointment.End = utc_to_outlook_local(end_dt)

        """
        for email in attendees:
            recipient = appointment.Recipients.Add(email)
            recipient.Type = 1  # 1 = Required
        appointment.Recipients.ResolveAll()
        """

        recurrence = appointment.GetRecurrencePattern()
        recurrence.RecurrenceType = 1  # 1 = olRecursWeekly
        recurrence.DayOfWeekMask = day_of_week_mask
        recurrence.Interval = interval
        recurrence.PatternStartDate = datetime.strptime(start_day,
                                                        '%Y-%m-%d')
        recurrence.PatternEndDate = datetime.strptime(end_day,
                                                      '%Y-%m-%d')

        appointment.Save()
        # appointment.Send()
        print_display(f'{line_number()} Weekly recurring Outlook calendar event created successfully.')

    def create_ms_outlook_monthly_recurrence(self,
                                             start_time: str,
                                             end_time: str,
                                             start_day: str,
                                             end_day: str,
                                             title: str,
                                             location: str,
                                             body: str,
                                             interval: int = 1,

                                             # attendees: list

                                             ):
        """
        Creates a plain monthly recurring Outlook appointment (olRecursMonthly, type 2).
        The event repeats on the same date each month as defined by start_day.

        interval: how many months between occurrences, default 1
        """
        ms_outlook_helper = MicrosoftOutlookHelper()
        appointment = ms_outlook_helper.ms_outlook_create_item()
        start_dt = datetime.strptime(f'{start_day} {start_time}',
                                     '%Y-%m-%d %H:%M')
        end_dt = datetime.strptime(f'{start_day} {end_time}',
                                   '%Y-%m-%d %H:%M')
        appointment.Subject = title
        appointment.Location = location
        appointment.Body = body
        appointment.Start = utc_to_outlook_local(start_dt)
        appointment.End = utc_to_outlook_local(end_dt)

        """
        for email in attendees:
            recipient = appointment.Recipients.Add(email)
            recipient.Type = 1  # 1 = Required
        appointment.Recipients.ResolveAll()
        """

        recurrence = appointment.GetRecurrencePattern()
        recurrence.RecurrenceType = 2  # 2 = olRecursMonthly
        recurrence.Interval = interval
        recurrence.PatternStartDate = datetime.strptime(start_day,
                                                        '%Y-%m-%d')
        recurrence.PatternEndDate = datetime.strptime(end_day,
                                                      '%Y-%m-%d')

        appointment.Save()
        # appointment.Send()
        print_display(f'{line_number()} Monthly recurring Outlook calendar event created successfully.')

    def create_ms_outlook_monthly_nth_recurrence(self,
                                                 start_time: str,
                                                 end_time: str,
                                                 start_day: str,
                                                 end_day: str,
                                                 title: str,
                                                 location: str,
                                                 body: str,
                                                 day_of_week: str,
                                                 instance: int,
                                                 interval: int = 1,

                                                 # attendees: list

                                                 ):
        """
        Creates a monthly-nth recurring Outlook appointment (olRecursMonthNth, type 3).
        The event repeats on the Nth weekday of each month, e.g. '3rd Thursday'.

        day_of_week: one day name string, e.g. 'Thursday'
                     Accepted values: 'Sunday', 'Monday', 'Tuesday', 'Wednesday',
                                      'Thursday', 'Friday', 'Saturday'
        instance:    which occurrence within the month: 1=first, 2=second, 3=third,
                     4=fourth, 5=last
        interval:    how many months between occurrences, default 1
        """
        day_mask_map = {
                'Sunday'   : 1,
                'Monday'   : 2,
                'Tuesday'  : 4,
                'Wednesday': 8,
                'Thursday' : 16,
                'Friday'   : 32,
                'Saturday' : 64,
        }
        day_of_week_mask = day_mask_map.get(day_of_week,
                                            0)
        ms_outlook_helper = MicrosoftOutlookHelper()
        appointment = ms_outlook_helper.ms_outlook_create_item()
        start_dt = datetime.strptime(f'{start_day} {start_time}',
                                     '%Y-%m-%d %H:%M')
        end_dt = datetime.strptime(f'{start_day} {end_time}',
                                   '%Y-%m-%d %H:%M')
        appointment.Subject = title
        appointment.Location = location
        appointment.Body = body
        appointment.Start = utc_to_outlook_local(start_dt)
        appointment.End = utc_to_outlook_local(end_dt)

        """
        for email in attendees:
            recipient = appointment.Recipients.Add(email)
            recipient.Type = 1  # 1 = Required
        appointment.Recipients.ResolveAll()
        """

        recurrence = appointment.GetRecurrencePattern()
        recurrence.RecurrenceType = 3  # 3 = olRecursMonthNth
        recurrence.DayOfWeekMask = day_of_week_mask
        recurrence.Instance = instance
        recurrence.Interval = interval
        recurrence.PatternStartDate = datetime.strptime(start_day,
                                                        '%Y-%m-%d')
        recurrence.PatternEndDate = datetime.strptime(end_day,
                                                      '%Y-%m-%d')

        appointment.Save()
        # appointment.Send()
        print_display(f'{line_number()} Monthly-nth recurring Outlook calendar event created successfully.')

    def create_ms_outlook_yearly_recurrence(self,
                                            start_time: str,
                                            end_time: str,
                                            start_day: str,
                                            end_day: str,
                                            title: str,
                                            location: str,
                                            body: str,

                                            # attendees: list

                                            ):
        """
        Creates a plain yearly recurring Outlook appointment (olRecursYearly, type 5).
        The event repeats on the same date and month each year as defined by start_day.
        """
        ms_outlook_helper = MicrosoftOutlookHelper()
        appointment = ms_outlook_helper.ms_outlook_create_item()
        start_dt = datetime.strptime(f'{start_day} {start_time}',
                                     '%Y-%m-%d %H:%M')
        end_dt = datetime.strptime(f'{start_day} {end_time}',
                                   '%Y-%m-%d %H:%M')
        appointment.Subject = title
        appointment.Location = location
        appointment.Body = body
        appointment.Start = utc_to_outlook_local(start_dt)
        appointment.End = utc_to_outlook_local(end_dt)

        """
        for email in attendees:
            recipient = appointment.Recipients.Add(email)
            recipient.Type = 1  # 1 = Required
        appointment.Recipients.ResolveAll()
        """

        recurrence = appointment.GetRecurrencePattern()
        recurrence.RecurrenceType = 5  # 5 = olRecursYearly
        recurrence.PatternStartDate = datetime.strptime(start_day,
                                                        '%Y-%m-%d')
        recurrence.PatternEndDate = datetime.strptime(end_day,
                                                      '%Y-%m-%d')

        appointment.Save()
        # appointment.Send()
        print_display(f'{line_number()} Yearly recurring Outlook calendar event created successfully.')

    def create_ms_outlook_yearly_nth_recurrence(self,
                                                start_time: str,
                                                end_time: str,
                                                start_day: str,
                                                end_day: str,
                                                title: str,
                                                location: str,
                                                body: str,
                                                day_of_week: str,
                                                instance: int,
                                                month_of_year: int,

                                                # attendees: list

                                                ):
        """
        Creates a yearly-nth recurring Outlook appointment (olRecursYearNth, type 6).
        The event repeats on the Nth weekday of a specific month each year,
        e.g. 'last Friday of November'.

        day_of_week:    one day name string, e.g. 'Friday'
                        Accepted values: 'Sunday', 'Monday', 'Tuesday', 'Wednesday',
                                         'Thursday', 'Friday', 'Saturday'
        instance:       which occurrence within the month: 1=first, 2=second, 3=third,
                        4=fourth, 5=last
        month_of_year:  integer month number, e.g. 11 for November
        """
        day_mask_map = {
                'Sunday'   : 1,
                'Monday'   : 2,
                'Tuesday'  : 4,
                'Wednesday': 8,
                'Thursday' : 16,
                'Friday'   : 32,
                'Saturday' : 64,
        }
        day_of_week_mask = day_mask_map.get(day_of_week,
                                            0)
        ms_outlook_helper = MicrosoftOutlookHelper()
        appointment = ms_outlook_helper.ms_outlook_create_item()
        start_dt = datetime.strptime(f'{start_day} {start_time}',
                                     '%Y-%m-%d %H:%M')
        end_dt = datetime.strptime(f'{start_day} {end_time}',
                                   '%Y-%m-%d %H:%M')
        appointment.Subject = title
        appointment.Location = location
        appointment.Body = body
        appointment.Start = utc_to_outlook_local(start_dt)
        appointment.End = utc_to_outlook_local(end_dt)

        """
        for email in attendees:
            recipient = appointment.Recipients.Add(email)
            recipient.Type = 1  # 1 = Required
        appointment.Recipients.ResolveAll()
        """

        recurrence = appointment.GetRecurrencePattern()
        recurrence.RecurrenceType = 6  # 6 = olRecursYearNth
        recurrence.DayOfWeekMask = day_of_week_mask
        recurrence.Instance = instance
        recurrence.MonthOfYear = month_of_year
        recurrence.PatternStartDate = datetime.strptime(start_day,
                                                        '%Y-%m-%d')
        recurrence.PatternEndDate = datetime.strptime(end_day,
                                                      '%Y-%m-%d')

        appointment.Save()
        # appointment.Send()
        print_display(f'{line_number()} Yearly-nth recurring Outlook calendar event created successfully.')
    # NEW SETUP END

    def create_single_event(self,
                            start_time: str,
                            end_time: str,
                            start_day: str,
                            title: str,
                            location: str,
                            body: str,

                            # attendees: list,

                            default='g_calendar'):
        if default == 'g_calendar':
            self.create_g_calendar_single_event(start_time,
                                                end_time,
                                                start_day,
                                                title,
                                                location,
                                                body,

                                                # attendees
                                                )
        else:
            self.create_ms_outlook_single_event(start_time,
                                                end_time,
                                                start_day,
                                                title,
                                                location,
                                                body,

                                                # attendees

                                                )

    def create_daily_recurrence(self,
                                start_time: str,
                                end_time: str,
                                start_day: str,
                                end_day: str,
                                title: str,
                                location: str,
                                body: str,

                                # attendees: list,

                                default='g_calendar'):
        if default == 'g_calendar':
            self.create_g_calendar_daily_recurrence(start_time,
                                                    end_time,
                                                    start_day,
                                                    end_day,
                                                    title,
                                                    location,
                                                    body,

                                                    # attendees

                                                    )
        else:
            self.create_ms_outlook_daily_recurrence(start_time,
                                                    end_time,
                                                    start_day,
                                                    end_day,
                                                    title,
                                                    location,
                                                    body,

                                                    # attendees

                                                    )

    def setup_mockup_appointments(self,
                                  event_title,
                                  side='ms_outlook',
                                  enabled=False):
        today = date.today()
        monday = today - timedelta(days=today.weekday())
        first_day = monday.day
        last_day = first_day + 4
        current_year_month = monday.strftime('%Y-%m')
        if enabled:
            past_days = (date.today() - timedelta(days=50)).strftime('%Y-%m-%d')

            self.create_daily_recurrence(start_time='12:00',
                                         end_time='13:00',
                                         start_day=past_days,
                                         end_day=f'{current_year_month}-{last_day:02d}',
                                         title=f'Evento X {event_title} Recurrence',
                                         location='Conference Room / Teams',
                                         body='Daily project sync meeting.',

                                         # attendees=['user1@me.con',
                                         #            'user2@me.con'],

                                         default=side)

            self.create_daily_recurrence(start_time='19:00',
                                         end_time='20:00',
                                         start_day=f'{current_year_month}-{first_day:02d}',
                                         end_day=f'{current_year_month}-{last_day:02d}',
                                         title=f'Evento 1 {event_title} Recurrence',
                                         location='Conference Room / Teams',
                                         body='Daily project sync meeting.',

                                         # attendees=['user1@me.con',
                                         #            'user2@me.con'],

                                         default=side)
            self.create_daily_recurrence(start_time='20:00',
                                         end_time='21:00',
                                         start_day=f'{current_year_month}-{first_day:02d}',
                                         end_day=f'{current_year_month}-{last_day:02d}',
                                         title=f'Evento 2 {event_title} Recurrence',
                                         location='Conference Room / Teams',
                                         body='Daily project sync meeting.',

                                         # attendees=['user1@me.con',
                                         #            'user2@me.con'],

                                         default=side)
            self.create_single_event(start_time='14:00',
                                     end_time='15:00',
                                     start_day=f'{current_year_month}-{first_day:02d}',
                                     title=f'Evento 0_{event_title} Single',
                                     location='Room 402 / Teams',
                                     body='Review of project milestones and next steps.',

                                     # attendees=['user1@me.con',
                                     #            'user2@me.con'],

                                     default=side)
            self.create_single_event(start_time='15:00',
                                     end_time='16:00',
                                     start_day=f'{current_year_month}-{(first_day + 1):02d}',
                                     title=f'Evento 1_{event_title} Single',
                                     location='Room 402 / Teams',
                                     body='Review of project milestones and next steps.',

                                     # attendees=['user1@me.con',
                                     #            'user2@me.con'],

                                     default=side)
            self.create_single_event(start_time='16:00',
                                     end_time='17:00',
                                     start_day=f'{current_year_month}-{(first_day + 2):02d}',
                                     title=f'Evento 2_{event_title} Single',
                                     location='Room 402 / Teams',
                                     body='Review of project milestones and next steps.',

                                     # attendees=['user1@me.con',
                                     #            'user2@me.con'],

                                     default=side)
            self.create_single_event(start_time='17:00',
                                     end_time='18:00',
                                     start_day=f'{current_year_month}-{(first_day + 3):02d}',
                                     title=f'Evento 3_{event_title} Single',
                                     location='Room 402 / Teams',
                                     body='Review of project milestones and next steps.',

                                     # attendees=['user1@me.con',
                                     #            'user2@me.con'],

                                     default=side)
            self.create_single_event(start_time='18:00',
                                     end_time='19:00',
                                     start_day=f'{current_year_month}-{(first_day + 4):02d}',
                                     title=f'Evento 4_{event_title} Single',
                                     location='Room 402 / Teams',
                                     body='Review of project milestones and next steps.',

                                     # attendees=['user1@me.con',
                                     #            'user2@me.con'],

                                     default=side)

    # NEW SETUP START
            if side == 'ms_outlook':
                # Weekly Mon-Fri
                self.create_ms_outlook_weekly_recurrence(
                        start_time='09:00',
                        end_time='09:30',
                        start_day=f'{current_year_month}-{first_day:02d}',
                        end_day=f'{current_year_month}-{last_day:02d}',
                        title=f'Evento W1_{event_title} Weekly Mon-Fri',
                        location='Conference Room / Teams',
                        body='Weekly standup Mon to Fri.',
                        days_of_week=['Monday',
                                      'Tuesday',
                                      'Wednesday',
                                      'Thursday',
                                      'Friday'])

                # Weekly on Monday and Wednesday only
                self.create_ms_outlook_weekly_recurrence(
                        start_time='10:00',
                        end_time='10:30',
                        start_day=f'{current_year_month}-{first_day:02d}',
                        end_day=f'{current_year_month}-{last_day:02d}',
                        title=f'Evento W2_{event_title} Weekly Mon+Wed',
                        location='Conference Room / Teams',
                        body='Biweekly sync on Monday and Wednesday.',
                        days_of_week=['Monday',
                                      'Wednesday'])

                # Plain monthly — same date each month
                self.create_ms_outlook_monthly_recurrence(
                        start_time='11:00',
                        end_time='12:00',
                        start_day=f'{current_year_month}-{first_day:02d}',
                        end_day=(monday + timedelta(days=180)).strftime('%Y-%m-%d'),
                        title=f'Evento M1_{event_title} Monthly',
                        location='Room 402 / Teams',
                        body='Monthly review on the same date.')

                # Monthly-nth — 3rd Thursday of each month
                self.create_ms_outlook_monthly_nth_recurrence(
                        start_time='14:00',
                        end_time='15:00',
                        start_day=f'{current_year_month}-{first_day:02d}',
                        end_day=(monday + timedelta(days=180)).strftime('%Y-%m-%d'),
                        title=f'Evento M2_{event_title} Monthly 3rd Thursday',
                        location='Room 402 / Teams',
                        body='Monthly review on the 3rd Thursday.',
                        day_of_week='Thursday',
                        instance=3)

                # Plain yearly — same date every year
                self.create_ms_outlook_yearly_recurrence(
                        start_time='09:00',
                        end_time='10:00',
                        start_day=f'{current_year_month}-{first_day:02d}',
                        end_day=(monday + timedelta(days=730)).strftime('%Y-%m-%d'),
                        title=f'Evento Y1_{event_title} Yearly',
                        location='Room 402 / Teams',
                        body='Annual review on the same date.')

                # Yearly-nth — last Friday of November each year
                self.create_ms_outlook_yearly_nth_recurrence(
                        start_time='15:00',
                        end_time='16:00',
                        start_day=f'{current_year_month}-{first_day:02d}',
                        end_day=(monday + timedelta(days=730)).strftime('%Y-%m-%d'),
                        title=f'Evento Y2_{event_title} Yearly Last Friday Nov',
                        location='Room 402 / Teams',
                        body='Annual review on the last Friday of November.',
                        day_of_week='Friday',
                        instance=5,
                        month_of_year=11)
    # NEW SETUP END

if __name__ == '__main__':
    event_title = 'OUTLOOK|31|G_CAL'
    calendar_samples = 'ms_outlook'
    # calendar_samples = 'g_calendar'
    generate_calendar_samples = True
    reset_event_mapping_file = True
    if reset_event_mapping_file:
        event_mapping = EventMapping()
        event_mapping.clear_map()
    message_setup = MessageSetup()
    message_setup.setup_mockup_appointments(event_title,
                                            side=calendar_samples,
                                            enabled=generate_calendar_samples)
