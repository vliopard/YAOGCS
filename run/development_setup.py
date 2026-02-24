from connector.event_mapping import EventMapping
from run.message_setup import MessageSetup

if __name__ == '__main__':
    event_title = 'OUTLOOK=15=G_CAL'
    calendar_samples = 'ms_outlook'
    #calendar_samples = 'g_calendar'
    generate_calendar_samples = False
    reset_event_mapping_file = True
    if reset_event_mapping_file:
        event_mapping = EventMapping()
        event_mapping.reset()
    message_setup = MessageSetup()
    message_setup.setup_mockup_appointments(event_title,
                                            side=calendar_samples,
                                            enabled=generate_calendar_samples)
