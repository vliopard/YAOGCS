from datetime import datetime

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
        recurrence.RecurrenceType = 0  # 0 = Daily
        recurrence.PatternStartDate = datetime.strptime(start_day,
                                                        '%Y-%m-%d')
        recurrence.PatternEndDate = datetime.strptime(end_day,
                                                      '%Y-%m-%d')

        appointment.Save()
        # appointment.Send()
        print_display(f'{line_number()} Recurring Outlook calendar event created successfully.')

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
        first_day = 16
        last_day = first_day + 4
        if enabled:
            self.create_daily_recurrence(start_time='19:00',
                                         end_time='20:00',
                                         start_day=f'2026-02-{first_day:02d}',
                                         end_day=f'2026-02-{last_day:02d}',
                                         title=f'Evento 1 {event_title} Recurrence',
                                         location='Conference Room / Teams',
                                         body='Daily project sync meeting.',

                                         # attendees=['user1@me.con',
                                         #            'user2@me.con'],

                                         default=side)
            self.create_daily_recurrence(start_time='20:00',
                                         end_time='21:00',
                                         start_day=f'2026-02-{first_day:02d}',
                                         end_day=f'2026-02-{last_day:02d}',
                                         title=f'Evento 2 {event_title} Recurrence',
                                         location='Conference Room / Teams',
                                         body='Daily project sync meeting.',

                                         # attendees=['user1@me.con',
                                         #            'user2@me.con'],

                                         default=side)
            self.create_single_event(start_time='14:00',
                                     end_time='15:00',
                                     start_day=f'2026-02-{first_day:02d}',
                                     title=f'Evento 0_{event_title} Single',
                                     location='Room 402 / Teams',
                                     body='Review of project milestones and next steps.',

                                     # attendees=['user1@me.con',
                                     #            'user2@me.con'],

                                     default=side)
            self.create_single_event(start_time='15:00',
                                     end_time='16:00',
                                     start_day=f'2026-02-{(first_day + 1):02d}',
                                     title=f'Evento 1_{event_title} Single',
                                     location='Room 402 / Teams',
                                     body='Review of project milestones and next steps.',

                                     # attendees=['user1@me.con',
                                     #            'user2@me.con'],

                                     default=side)
            self.create_single_event(start_time='16:00',
                                     end_time='17:00',
                                     start_day=f'2026-02-{(first_day + 2):02d}',
                                     title=f'Evento 2_{event_title} Single',
                                     location='Room 402 / Teams',
                                     body='Review of project milestones and next steps.',

                                     # attendees=['user1@me.con',
                                     #            'user2@me.con'],

                                     default=side)
            self.create_single_event(start_time='17:00',
                                     end_time='18:00',
                                     start_day=f'2026-02-{(first_day + 3):02d}',
                                     title=f'Evento 3_{event_title} Single',
                                     location='Room 402 / Teams',
                                     body='Review of project milestones and next steps.',

                                     # attendees=['user1@me.con',
                                     #            'user2@me.con'],

                                     default=side)
            self.create_single_event(start_time='18:00',
                                     end_time='19:00',
                                     start_day=f'2026-02-{(first_day + 4):02d}',
                                     title=f'Evento 4_{event_title} Single',
                                     location='Room 402 / Teams',
                                     body='Review of project milestones and next steps.',

                                     # attendees=['user1@me.con',
                                     #            'user2@me.con'],

                                     default=side)


if __name__ == '__main__':
    event_title = 'OUTLOOK=15=G_CAL'
    calendar_samples = 'ms_outlook'
    # calendar_samples = 'g_calendar'
    generate_calendar_samples = False
    reset_event_mapping_file = True
    if reset_event_mapping_file:
        event_mapping = EventMapping()
        event_mapping.clear_map()
    message_setup = MessageSetup()
    message_setup.setup_mockup_appointments(event_title,
                                            side=calendar_samples,
                                            enabled=generate_calendar_samples)
