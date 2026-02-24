from connector.event_mapping import EventMapping
from connector.g_calendar import GoogleCalendarConnector

if __name__ == '__main__':
    event_mapping = EventMapping()
    event_mapping.reset()
    local_g_calendar_connection = GoogleCalendarConnector()
    events = local_g_calendar_connection.get_g_calendar_events()
    for event in events:
        print(f'RESULT [{local_g_calendar_connection.g_calendar_delete_event(event)}]')
