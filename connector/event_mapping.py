import json
import os
from enum import Enum
from pathlib import Path
from threading import Lock
from typing import Optional
from typing import Tuple

from utils.handling_time import utc_now
from utils.utils import line_number
from utils.utils import print_display


class EventSide(Enum):
    MS_OUTLOOK = 'ms_outlook'
    G_CALENDAR = 'g_calendar'


class EventMapping:
    VERSION = '1.0'

    def __init__(self):
        base_dir = Path(__file__).resolve().parent.parent
        database_dir = (base_dir / 'resources' / 'database').resolve()

        self.event_map_file = str(database_dir / 'event_map.json')
        self._lock = Lock()
        self._ensure_directory()
        self.event_map = self._load_map()

    def _ensure_directory(self):
        event_map_directory = os.path.dirname(self.event_map_file)
        if event_map_directory and not os.path.exists(event_map_directory):
            os.makedirs(event_map_directory,
                        exist_ok=True)

    def _get_default_structure(self) -> dict:
        return {
                'single_events'   : {},
                'recurrent_events': {},
                'metadata'        : {
                        'version'  : self.VERSION,
                        'last_sync': utc_now()}}

    def _load_map(self) -> dict:
        if not os.path.exists(self.event_map_file):
            return self._get_default_structure()

        try:
            with open(self.event_map_file,
                      'r',
                      encoding='utf-8') as file_reader:
                data = json.load(file_reader)
                if 'single_events' not in data or 'recurrent_events' not in data:
                    return self._get_default_structure()
                return data
        except (json.JSONDecodeError,
                IOError) as errors:
            backup_file = f'{self.event_map_file}.backup.{utc_now()}'
            if os.path.exists(self.event_map_file):
                os.rename(self.event_map_file,
                          backup_file)
            print_display(f'{line_number()} Warning: Corrupted mapping file backed up to {backup_file}. Error: {errors}')
            return self._get_default_structure()

    def _save_map(self):
        temp_file = f'{self.event_map_file}.tmp'
        try:
            self.event_map['metadata']['last_sync'] = utc_now()
            with open(temp_file,
                      'w',
                      encoding='utf-8') as f:
                json.dump(self.event_map,
                          f,
                          indent=4,
                          ensure_ascii=False)
            os.replace(temp_file,
                       self.event_map_file)
        except Exception as exception:
            if os.path.exists(temp_file):
                os.remove(temp_file)
            raise IOError(f'Failed to save mapping: {exception}')

    def _identify_side(self,
                       event_id: str,
                       mapping_dict: dict) -> Optional[EventSide]:
        if event_id in mapping_dict:
            return EventSide.MS_OUTLOOK
        for ms_outlook_id, g_calendar_id in mapping_dict.items():
            if g_calendar_id == event_id:
                return EventSide.G_CALENDAR
        return None

    def reset(self):
        with self._lock:
            self.event_map = self._get_default_structure()
            self._save_map()
            print_display(f'{line_number()} Event mapping cleared. Reset to empty state.')

    def add_single_event(self,
                         ms_outlook_id: str,
                         g_calendar_id: str) -> bool:
        with self._lock:
            single_events = self.event_map['single_events']
            if ms_outlook_id in single_events:
                return False
            single_events[ms_outlook_id] = g_calendar_id
            self._save_map()
            return True

    def get_single_event_pair(self,
                              event_id: str) -> Optional[Tuple[str, Optional[str]]]:
        with self._lock:
            single_events = self.event_map['single_events']
            side = self._identify_side(event_id,
                                       single_events)
            if side == EventSide.MS_OUTLOOK:
                return (event_id,
                        single_events[event_id])
            elif side == EventSide.G_CALENDAR:
                for ms_outlook_id, g_calendar_id in single_events.items():
                    if g_calendar_id == event_id:
                        return ms_outlook_id, g_calendar_id
            return None

    def remove_single_event(self,
                            event_id: str) -> bool:
        with self._lock:
            single_events = self.event_map['single_events']
            side = self._identify_side(event_id,
                                       single_events)
            if side == EventSide.MS_OUTLOOK:
                del single_events[event_id]
                self._save_map()
                return True
            elif side == EventSide.G_CALENDAR:
                for ms_outlook_id, g_calendar_id in single_events.items():
                    if g_calendar_id == event_id:
                        del single_events[ms_outlook_id]
                        self._save_map()
                        return True
            return False

    def add_recurrent_master(self,
                             ms_outlook_master_id: str,
                             g_calendar_master_id: str) -> bool:
        with self._lock:
            recurrent_events = self.event_map['recurrent_events']
            if ms_outlook_master_id in recurrent_events:
                return False
            recurrent_events[ms_outlook_master_id] = {
                    'g_calendar_master_id': g_calendar_master_id,
                    'instances'           : {}}
            self._save_map()
            return True

    def add_recurrent_instance(self,
                               master_id: str,
                               ms_outlook_instance_id: str,
                               g_calendar_instance_id: str) -> bool:
        with self._lock:
            recurrent_events = self.event_map['recurrent_events']
            ms_outlook_master_id = self._find_recurrent_master(master_id)
            if not ms_outlook_master_id:
                return False
            recurrent_events[ms_outlook_master_id]['instances'][ms_outlook_instance_id] = g_calendar_instance_id
            self._save_map()
            return True

    def _find_recurrent_master(self,
                               master_id: str) -> Optional[str]:
        recurrent_events = self.event_map['recurrent_events']
        if master_id in recurrent_events:
            return master_id
        for ms_outlook_master_id, ms_outlook_data in recurrent_events.items():
            if ms_outlook_data['g_calendar_master_id'] == master_id:
                return ms_outlook_master_id
        return None

    def get_recurrent_master_pair(self,
                                  master_id: str) -> Optional[Tuple[str, Optional[str]]]:
        with self._lock:
            ms_outlook_master_id = self._find_recurrent_master(master_id)
            if not ms_outlook_master_id:
                return None
            recurrent_events = self.event_map['recurrent_events']
            g_calendar_master_id = recurrent_events[ms_outlook_master_id]['g_calendar_master_id']
            return ms_outlook_master_id, g_calendar_master_id

    def remove_recurrence(self,
                          instance_id: str) -> bool:
        with self._lock:
            recurrent_events = self.event_map['recurrent_events']
            for ms_outlook_master_id, ms_outlook_data in recurrent_events.items():
                if ms_outlook_master_id == instance_id:
                    del recurrent_events[ms_outlook_master_id]
                    self._save_map()
                    return True
            return False

    def g_calendar_remove_recurrence(self,
                                     instance_id: str) -> bool:
        with self._lock:
            recurrent_events = self.event_map['recurrent_events']
            r_items = recurrent_events.items()
            for ms_outlook_master_id, ms_outlook_data in r_items:
                if ms_outlook_data['g_calendar_master_id'] == instance_id:
                    del recurrent_events[ms_outlook_master_id]
                    self._save_map()
                    return True
            return False

    def remove_recurrent_instance(self,
                                  instance_id: str) -> bool:
        with self._lock:
            recurrent_events = self.event_map['recurrent_events']
            for ms_outlook_master_id, ms_outlook_data in recurrent_events.items():
                ms_outlook_instances = ms_outlook_data['instances']
                side = self._identify_side(instance_id,
                                           ms_outlook_instances)
                if side == EventSide.MS_OUTLOOK:
                    if instance_id in ms_outlook_instances:
                        del ms_outlook_instances[instance_id]
                        if not ms_outlook_instances:
                            del recurrent_events[ms_outlook_master_id]
                        self._save_map()
                        return True
                elif side == EventSide.G_CALENDAR:
                    for ms_outlook_instance_id, g_calendar_instance_id in list(ms_outlook_instances.items()):
                        if g_calendar_instance_id == instance_id:
                            del ms_outlook_instances[ms_outlook_instance_id]
                            if not ms_outlook_instances:
                                del recurrent_events[ms_outlook_master_id]
                            self._save_map()
                            return True
            return False

    def get_all_mappings(self) -> dict:
        with self._lock:
            return json.loads(json.dumps(self.event_map))
