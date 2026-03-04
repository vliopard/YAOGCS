import json
import os
from enum import Enum
from pathlib import Path
from threading import Lock
from typing import Optional
from typing import Tuple

from system.tools import line_number
from system.tools import print_box
from system.tools import print_display
from system.tools import utc_now


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
                'single_events'     : {},
                'single_events_meta': {},
                'recurrent_events'  : {},
                'metadata'          : {
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
                # migrate existing maps that predate single_events_meta
                if 'single_events_meta' not in data:
                    data['single_events_meta'] = dict()
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
            # On Windows, os.replace can fail with PermissionError if the
            # destination file has a read-only attribute or is briefly locked.
            # Explicitly clear the read-only flag before replacing.
            if os.path.exists(self.event_map_file):
                os.chmod(self.event_map_file,
                         0o666)
            os.replace(temp_file,
                       self.event_map_file)
        except Exception as exception:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except OSError:
                    pass
            raise IOError(f'Failed to save mapping: {exception}')

    def clear_map(self):
        with self._lock:
            self.event_map = self._get_default_structure()
            self._save_map()
            print_display(f'{line_number()} Event mapping cleared. Reset to empty state.')

    def _identify_side(self,
                       instance_id: str,
                       mapping_dict: dict) -> Optional[EventSide]:
        if instance_id in mapping_dict:
            return EventSide.MS_OUTLOOK
        for ms_outlook_id, g_calendar_id in mapping_dict.items():
            if g_calendar_id == instance_id:
                return EventSide.G_CALENDAR
        return None

    def _find_recurrent_master(self,
                               master_id: str) -> Optional[str]:
        recurrent_events = self.event_map['recurrent_events']
        if master_id in recurrent_events:
            return master_id
        for ms_outlook_master_id, ms_outlook_data in recurrent_events.items():
            if ms_outlook_data['g_calendar_master_id'] == master_id:
                return ms_outlook_master_id
        return None

    def get_all_instances(self) -> dict:
        with self._lock:
            return json.loads(json.dumps(self.event_map))

    def get_instance_pair(self,
                          event_id: str) -> Optional[Tuple[str, Optional[str]]]:
        with self._lock:
            print_box(f'{line_number()} [EVENT MAPPING] recovering: [{event_id}]')
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

    def get_recurrent_pair(self,
                           master_id: str) -> Optional[Tuple[str, Optional[str]]]:
        with self._lock:
            print_box(f'{line_number()} [EVENT MAPPING] get recurrent pair: [{master_id}]')
            ms_outlook_master_id = self._find_recurrent_master(master_id)
            if not ms_outlook_master_id:
                return None
            recurrent_events = self.event_map['recurrent_events']
            g_calendar_master_id = recurrent_events[ms_outlook_master_id]['g_calendar_master_id']
            return ms_outlook_master_id, g_calendar_master_id

    def insert_instance(self,
                        ms_outlook_id: str,
                        g_calendar_id: str,
                        instance_name: str = None) -> bool:
        with self._lock:
            print_box(f'{line_number()} [EVENT MAPPING] inserting instance: [{ms_outlook_id}]')
            single_events = self.event_map['single_events']
            if ms_outlook_id in single_events:
                return False
            single_events[ms_outlook_id] = g_calendar_id
            if instance_name:
                self.event_map['single_events_meta'][ms_outlook_id] = f'[{instance_name}]'
            self._save_map()
            return True

    def insert_recurrence(self,
                          ms_outlook_master_id: str,
                          g_calendar_master_id: str,
                          instance_name: str = None) -> bool:
        with self._lock:
            print_box(f'{line_number()} [EVENT MAPPING] inserting recurrence: [{ms_outlook_master_id}]')
            recurrent_events = self.event_map['recurrent_events']
            if ms_outlook_master_id in recurrent_events:
                return False
            recurrent_events[ms_outlook_master_id] = {
                    'g_calendar_master_id': g_calendar_master_id,
                    'instance_name'       : f'[{instance_name}]',
                    'instances'           : {}}
            self._save_map()
            return True

    def insert_occurrence(self,
                          master_id: str,
                          ms_outlook_instance_id: str,
                          g_calendar_instance_id: str) -> bool:
        with self._lock:
            print_box(f'{line_number()} [EVENT MAPPING] inserting occurrence: [{ms_outlook_instance_id}]')
            recurrent_events = self.event_map['recurrent_events']
            ms_outlook_master_id = self._find_recurrent_master(master_id)
            if not ms_outlook_master_id:
                return False
            recurrent_events[ms_outlook_master_id]['instances'][ms_outlook_instance_id] = g_calendar_instance_id
            self._save_map()
            return True

    def remove_instance(self,
                        event_id: str) -> bool:
        with self._lock:
            print_box(f'{line_number()} [EVENT MAPPING] removing instance: [{event_id}]')
            single_events = self.event_map['single_events']
            single_events_meta = self.event_map['single_events_meta']
            side = self._identify_side(event_id,
                                       single_events)
            if side == EventSide.MS_OUTLOOK:
                del single_events[event_id]
                single_events_meta.pop(event_id,
                                       None)
                self._save_map()
                return True
            elif side == EventSide.G_CALENDAR:
                for ms_outlook_id, g_calendar_id in single_events.items():
                    if g_calendar_id == event_id:
                        del single_events[ms_outlook_id]
                        single_events_meta.pop(ms_outlook_id,
                                               None)
                        self._save_map()
                        return True
            return False

    def remove_g_calendar_recurrence(self,
                                     g_calendar_instance_id: str) -> bool:
        with self._lock:
            print_box(f'{line_number()} [EVENT MAPPING] removing [Google Calendar] recurrence: [{g_calendar_instance_id}]')
            recurrent_events = self.event_map['recurrent_events']
            r_items = recurrent_events.items()
            for ms_outlook_master_id, ms_outlook_data in r_items:
                if ms_outlook_data['g_calendar_master_id'] == g_calendar_instance_id:
                    del recurrent_events[ms_outlook_master_id]
                    self._save_map()
                    return True
            return False

    def remove_ms_outlook_recurrence(self,
                                     ms_outlook_instance_id: str) -> bool:
        with self._lock:
            print_box(f'{line_number()} [EVENT MAPPING] removing [Microsoft Outlook] recurrence: [{ms_outlook_instance_id}]')
            recurrent_events = self.event_map['recurrent_events']
            for ms_outlook_master_id, ms_outlook_data in recurrent_events.items():
                if ms_outlook_master_id == ms_outlook_instance_id:
                    del recurrent_events[ms_outlook_master_id]
                    self._save_map()
                    return True
            return False

    def remove_generic_occurrence(self,
                                  generic_instance_id: str) -> bool:
        with self._lock:
            print_box(f'{line_number()} [EVENT MAPPING] removing generic: [{generic_instance_id}]')
            recurrent_events = self.event_map['recurrent_events']
            for ms_outlook_master_id, ms_outlook_data in recurrent_events.items():
                ms_outlook_instances = ms_outlook_data['instances']
                side = self._identify_side(generic_instance_id,
                                           ms_outlook_instances)
                if side == EventSide.MS_OUTLOOK:
                    if generic_instance_id in ms_outlook_instances:
                        del ms_outlook_instances[generic_instance_id]
                        if not ms_outlook_instances:
                            del recurrent_events[ms_outlook_master_id]
                        self._save_map()
                        return True
                elif side == EventSide.G_CALENDAR:
                    for ms_outlook_instance_id, g_calendar_instance_id in list(ms_outlook_instances.items()):
                        if g_calendar_instance_id == generic_instance_id:
                            del ms_outlook_instances[ms_outlook_instance_id]
                            if not ms_outlook_instances:
                                del recurrent_events[ms_outlook_master_id]
                            self._save_map()
                            return True
            return False
