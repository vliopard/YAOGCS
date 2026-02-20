import gc
import inspect
import os
import threading
import time
from datetime import datetime
from datetime import timedelta
from functools import wraps
from os import name as os_name
from platform import system
from sys import platform as sys_platform

import pywintypes

import utils.constants as constants
from utils.screen import get_terminal_width

_gui_log_queue = None


class PauseToken:
    """Cooperative pause token passed into sync routines.
    Each loop calls check() between iterations - blocks if paused, returns
    instantly if running. Raise PauseToken.Interrupted to abort early."""

    class Interrupted(Exception):
        pass

    def __init__(self, event: threading.Event):
        self._event = event

    def check(self):
        """Block while paused. If still paused after waking, raise Interrupted."""
        if not self._event.is_set():
            self._event.wait()
            if not self._event.is_set():
                raise PauseToken.Interrupted()

    @property
    def is_paused(self):
        return not self._event.is_set()


def set_log_queue(q):
    global _gui_log_queue
    _gui_log_queue = q


def print_debug(*arguments):
    if constants.DEBUG_MODE:
        print(constants.NEW_LINE)
        double_line()
        double_line(constants.TEXT_DEBUG_MESSAGE_START)
        double_line()
        print_display(*arguments)
        double_line()
        double_line(constants.TEXT_DEBUG_MESSAGE_END)
        double_line()
        print(constants.NEW_LINE)


def print_box(msg):
    print_underline()
    print_display(msg)
    print_overline()


def print_display(*arguments):
    argument_count = len(arguments)
    if argument_count == 0:
        return

    display_text = constants.SYMBOL_EMPTY
    for element in arguments:
        display_text = display_text + str(element) + constants.SYMBOL_BLANK
    display_text = display_text.rstrip()

    if display_text == constants.DEFAULT_L:
        return

    if constants.RUN_GUI:
        if _gui_log_queue is not None:
            _gui_log_queue.put(display_text)
    else:
        print(display_text)


def timed(func):
    @wraps(func)
    def wrapper(*args,
                **kwargs):
        start_time = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        time_start = time.time()
        result = func(*args,
                      **kwargs)
        time_end = time.time()
        end_time = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        time_report = [f'Start time: {start_time}',
                       f'End time:   {end_time}',
                       f'Function {func.__name__} ran in {timedelta(seconds=(time_end - time_start))}']
        print(f'{section_line(constants.SYMBOL_EQ, get_terminal_width() - 35)}')
        for time_detail in time_report:
            print(time_detail)
        print(f'{section_line(constants.SYMBOL_EQ, get_terminal_width() - 35)}')
        return result

    return wrapper


def section_line(style,
                 size):
    return style * size


def print_underline():
    print_display(constants.SYMBOL_UNDERLINE * (get_terminal_width()))


def print_overline():
    print_display(constants.SYMBOL_OVERLINE * (get_terminal_width()))


def double_line(message=None):
    if message:
        message = f'== {message} '
        print_display(message + constants.SYMBOL_EQ * (get_terminal_width() - len(message)))
    else:
        print_display(constants.SYMBOL_EQ * (get_terminal_width()))


def get_system():
    return system()


def object_serializer(original_object):
    if isinstance(original_object,
                  (str,
                   int,
                   float,
                   bool,
                   type(None))):
        return original_object
    if isinstance(original_object,
                  datetime):
        return original_object.isoformat()
    if isinstance(original_object,
                  dict):
        return {str(serial_key): object_serializer(serial_value) for serial_key, serial_value in original_object.items()}
    if isinstance(original_object,
                  (list,
                   tuple,
                   set)):
        return [object_serializer(serial_value) for serial_value in original_object]
    return str(original_object)


def _release(com_object_for_deletion):
    try:
        print_debug(com_object_for_deletion)
        del com_object_for_deletion
    except ValueError as value_error:
        print_display(f'{line_number()} ValueError during COM release: [{value_error}]')
        pass


def com_object_to_dictionary(com_object):
    dictionary_data = dict()
    for com_object_attribute in dir(com_object):
        if com_object_attribute.startswith('_'):
            continue
        if com_object_attribute == 'GetInspector':
            continue
        try:
            com_object_value = getattr(com_object,
                                       com_object_attribute)
            if not callable(com_object_value):
                dictionary_data[com_object_attribute] = com_object_value
        except pywintypes.com_error as com_error_type:
            print_display(f'{line_number()} ValueError for attribute [{com_object_attribute}]: [{com_error_type}]')
    _release(com_object)
    gc.collect()
    return dictionary_data


def parse_rule(rule: str):
    if ':' in rule:
        rule = rule.split(':',
                          1)[1]
    rule_parts = rule.split(';')
    rule = {rule_part.split('=')[0].strip().lower(): rule_part.split('=')[1].strip() for rule_part in rule_parts if '=' in rule_part}
    if 'freq' in rule:
        if rule['freq'] == 'DAILY':
            if 'interval' not in rule:
                rule['interval'] = '1'
    return dict(sorted(rule.items()))


def compare_rule(rule_one: str,
                 rule_two: str) -> bool:
    if isinstance(rule_one,
                  list):
        rule_one = rule_one[0]
    if isinstance(rule_two,
                  list):
        rule_two = rule_two[0]
    print_display(f'{line_number()} [{rule_one}]')
    print_display(f'{line_number()} [{rule_two}]')
    if not rule_one and not rule_two:
        return True
    rule_one_string = parse_rule(rule_one)
    rule_two_string = parse_rule(rule_two)
    print_display(f'{line_number()} [{rule_one_string}]')
    print_display(f'{line_number()} [{rule_two_string}]')
    return rule_one_string == rule_two_string


def get_nested_value(data,
                     path):
    for key in path.split('.'):
        if isinstance(data,
                      dict):
            data = data.get(key,
                            {})
        else:
            return None
    return data


def sort_json_list(data_list,
                   sort_path,
                   reverse=False):
    return sorted(data_list,
                  key=lambda x: get_nested_value(x,
                                                 sort_path),
                  reverse=reverse, )


def line_number():
    frame = inspect.currentframe().f_back
    filename = os.path.basename(frame.f_code.co_filename)
    func_name = frame.f_code.co_name
    line_no = frame.f_lineno
    return f'[{line_no:04d}:{filename}:{func_name}()]'


def get_platform():
    function_name = 'GET PLATFORM:'
    print_display(f'{line_number()} {constants.DEBUG_MARKER} {function_name} STARTED')
    platforms = {
            constants.PLATFORM_LINUX0: constants.OS_LINUX,
            constants.PLATFORM_LINUX1: constants.OS_LINUX,
            constants.PLATFORM_LINUX2: constants.OS_LINUX,
            constants.PLATFORM_DARWIN: constants.OS_X,
            constants.PLATFORM_WIN32 : constants.OS_WINDOWS}
    if sys_platform not in platforms:
        print_display(f'{line_number()} {constants.DEBUG_MARKER} {function_name} UNDESIRED END')
        print_display(f'{line_number()} {function_name} RETURN [{sys_platform}]')
        return sys_platform
    print_display(f'{line_number()} {constants.DEBUG_MARKER} {function_name} NORMAL END')
    return platforms[sys_platform]


def is_windows():
    if os_name == constants.OS_WINDOWS_NT:
        return True
    return False


def clear_screen():
    os.system('cls' if get_platform() == constants.OS_WINDOWS else 'clear')
