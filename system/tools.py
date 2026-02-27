import calendar
import gc
import inspect
import logging
import os
import platform
import re
import shlex
import struct
import subprocess
import time
from datetime import datetime
from datetime import timedelta
from datetime import timezone
from functools import wraps
from os import name as os_name
from platform import system
from sys import platform as sys_platform

import pywintypes
from dateutil import parser

import system.constants as constants

_gui_log_queue = None
_OUTLOOK_MAX_DATE = datetime(2080,
                             1,
                             1,
                             tzinfo=timezone.utc)
_OL_RECURRENCE_TYPE_TO_DELTA = {
        0: lambda interval: (lambda dt: dt + timedelta(days=interval)),
        1: lambda interval: (lambda dt: dt + timedelta(weeks=interval)),
        2: lambda interval: (lambda dt: _add_months(dt,
                                                    interval)),
        3: lambda interval: (lambda dt: _add_months(dt,
                                                    interval)),
        5: lambda interval: (lambda dt: _add_years(dt,
                                                   interval)),
        6: lambda interval: (lambda dt: _add_years(dt,
                                                   interval)), }


def set_log_queue(queue):
    global _gui_log_queue
    _gui_log_queue = queue


def measure_time(measured_function):
    @wraps(measured_function)
    def wrapper(*args,
                **kwargs):
        start_time = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        time_start = time.time()
        result = measured_function(*args,
                                   **kwargs)
        time_end = time.time()
        end_time = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        time_report = [f'Start time: {start_time}',
                       f'End time:   {end_time}',
                       f'Function {measured_function.__name__} ran in {timedelta(seconds=(time_end - time_start))}']
        print(f'{print_section_line(constants.SYMBOL_EQ, get_terminal_width() - 35)}')
        for time_detail in time_report:
            print(time_detail)
        print(f'{print_section_line(constants.SYMBOL_EQ, get_terminal_width() - 35)}')
        return result

    return wrapper


def trim_id(identification):
    if len(identification) > 30:
        return f'{identification[8:18]}|{identification[-20:-9]}'
    return identification


def release_com_object_memory(com_object_for_deletion):
    try:
        del com_object_for_deletion
    except ValueError as value_error:
        print_display(f'{line_number()} ValueError during COM release: [{value_error}]')
        pass


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


def get_master_id(text: str) -> str:
    return re.sub(r'_\d{8}T\d{6}Z$',
                  '',
                  text)


def get_system():
    return system()


def is_windows():
    if os_name == constants.OS_WINDOWS_NT:
        return True
    return False


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


def convert_com_object_to_dictionary(com_object):
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
    release_com_object_memory(com_object)
    gc.collect()
    return dictionary_data


def convert_object_to_string(original_object):
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
        return {str(serial_key): convert_object_to_string(serial_value) for serial_key, serial_value in original_object.items()}
    if isinstance(original_object,
                  (list,
                   tuple,
                   set)):
        return [convert_object_to_string(serial_value) for serial_value in original_object]
    return str(original_object)


def get_terminal_width():
    return (get_terminal_size()[0]) - 1


def _get_terminal_size_t_put():
    try:
        cols = int(subprocess.check_call(shlex.split('tput cols')))
        rows = int(subprocess.check_call(shlex.split('tput lines')))
        return cols, rows
    except Exception:
        pass


def get_terminal_size():
    current_os = platform.system()
    tuple_xy = None
    if current_os == 'Windows':
        tuple_xy = _get_terminal_size_windows()
        if tuple_xy is None:
            tuple_xy = _get_terminal_size_t_put()
    if current_os in ['Linux',
                      'Darwin'] or current_os.startswith('CYGWIN'):
        tuple_xy = _get_terminal_size_linux()
    if tuple_xy is None:
        tuple_xy = (80,
                    25)
    return tuple_xy


def _get_terminal_size_windows():
    try:
        from ctypes import windll, create_string_buffer
        standard_handle = windll.kernel32.GetStdHandle(-12)
        string_buffer = create_string_buffer(22)
        if windll.kernel32.GetConsoleScreenBufferInfo(standard_handle,
                                                      string_buffer):
            (buffer_x,
             buffer_y,
             current_x,
             current_y,
             w_attr,
             left,
             top,
             right,
             bottom,
             max_x,
             max_y) = struct.unpack('hhhhHhhhhhh',
                                    string_buffer.raw)
            size_x = right - left + 1
            size_y = bottom - top + 1
            return size_x, size_y
    except Exception:
        pass


def _get_terminal_size_linux():
    def ioctl_gwinsz(local_file_descriptor):
        try:
            import fcntl
            import termios
            return struct.unpack('hh',
                                 fcntl.ioctl(local_file_descriptor,
                                             termios.TIOCGWINSZ,
                                             '1234'))
        except Exception:
            pass

    coordinates = ioctl_gwinsz(0) or ioctl_gwinsz(1) or ioctl_gwinsz(2)
    if not coordinates:
        try:
            file_descriptor = os.open(os.ctermid(),
                                      os.O_RDONLY)
            coordinates = ioctl_gwinsz(file_descriptor)
            os.close(file_descriptor)
        except Exception:
            pass
    if not coordinates:
        try:
            coordinates = (os.environ['LINES'],
                           os.environ['COLUMNS'])
        except Exception:
            return None
    return int(coordinates[1]), int(coordinates[0])


def clear_screen():
    os.system('cls' if get_platform() == constants.OS_WINDOWS else 'clear')


def print_section_line(style,
                       size):
    return style * size


def print_double_line(message=None):
    if message:
        message = f'== {message} '
        print_display(message + constants.SYMBOL_EQ * (get_terminal_width() - len(message)))
    else:
        print_display(constants.SYMBOL_EQ * (get_terminal_width()))


def print_underline():
    print_display(constants.SYMBOL_UNDERLINE * (get_terminal_width()))


def print_overline():
    print_display(constants.SYMBOL_OVERLINE * (get_terminal_width()))


def print_box(message):
    print_underline()
    print_display(message)
    print_overline()


def print_debug(*arguments):
    if constants.DEBUG_MODE:
        print(constants.NEW_LINE)
        print_double_line()
        print_double_line(constants.TEXT_DEBUG_MESSAGE_START)
        print_double_line()
        print_display(*arguments)
        print_double_line()
        print_double_line(constants.TEXT_DEBUG_MESSAGE_END)
        print_double_line()
        print(constants.NEW_LINE)


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
        # if _gui_log_queue is not None:
        #     _gui_log_queue.put(display_text)
        display_text = " ".join(map(str,
                                    arguments)).rstrip()
        logger = logging.getLogger('CalendarSync Logger')
        logger.info(display_text)
    else:
        print(display_text)


def strip_symbols(message_with_symbols):
    return str(message_with_symbols).replace(':',
                                             '_').replace(' ',
                                                          '_').replace('+',
                                                                       '_').replace('-',
                                                                                    '_')


def utc_now():
    return datetime.now(timezone.utc).replace(tzinfo=None).isoformat()


def utc_to_outlook_local(dt_utc):
    if dt_utc.tzinfo is None:
        dt_utc = dt_utc.replace(tzinfo=timezone.utc)
    dt_local = dt_utc.astimezone()
    return dt_local.replace(tzinfo=None)


def convert_to_utc(date_time_value):
    if isinstance(date_time_value,
                  str):
        date_time = datetime.fromisoformat(date_time_value.replace('Z',
                                                                   '+00:00'))
    elif isinstance(date_time_value,
                    datetime):
        date_time = date_time_value
    else:
        raise TypeError(f'Unsupported type: [{type(date_time_value)}]')
    if date_time.tzinfo is None:
        local_tz = datetime.now().astimezone().tzinfo
        date_time = date_time.replace(tzinfo=local_tz)
    to_utc_value = date_time.astimezone(timezone.utc)
    print_display(f'{line_number()} [{date_time_value}] [{to_utc_value}]')
    return to_utc_value


def convert_to_local(date_time_value):
    if isinstance(date_time_value,
                  str):
        date_time = datetime.fromisoformat(date_time_value.replace('Z',
                                                                   '+00:00'))
    elif isinstance(date_time_value,
                    datetime):
        date_time = date_time_value
    else:
        raise TypeError(f'Unsupported type: {type(date_time_value)}')
    if date_time.tzinfo is None:
        date_time = date_time.replace(tzinfo=timezone.utc)
    return date_time.astimezone()


def convert_to_str(date_time_value):
    if isinstance(date_time_value,
                  dict):
        date_time_value = date_time_value.get('dateTime',
                                              '')
    if isinstance(date_time_value,
                  datetime):
        date_time_value = date_time_value.isoformat()
    return date_time_value


def time_when(time_shift):
    utc_time_now = datetime.now(timezone.utc)
    if time_shift < 0:
        utc_time_begin = utc_time_now - timedelta(days=abs(time_shift))
    elif time_shift > 0:
        utc_time_begin = utc_time_now + timedelta(days=time_shift)
    else:
        utc_time_begin = utc_time_now
    return utc_time_begin.isoformat().replace('+00:00',
                                              'Z')


def time_min():
    return time_when(-constants.DAY_PAST)


def time_max():
    return time_when(constants.DAY_NEXT)


def _add_months(date_time,
                date_months):
    date_month = date_time.month - 1 + date_months
    date_year = date_time.year + date_month // 12
    date_month = date_month % 12 + 1
    date_day = min(date_time.day,
                   calendar.monthrange(date_year,
                                       date_month)[1])
    return date_time.replace(year=date_year,
                             month=date_month,
                             day=date_day)


def _add_years(date_time,
               date_years):
    return _add_months(date_time,
                       date_years * 12)


def remove_timezone_info(date_time):
    if date_time is None:
        return None
    if isinstance(date_time,
                  str):
        date_time = parser.parse(date_time)
    wall_clock = date_time.replace(tzinfo=None)
    return wall_clock.strftime('%Y-%m-%dT%H:%M:%SZ')


def extract_date_full(text: str) -> str | None:
    match = re.search(r'(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z',
                      text)
    if not match:
        return None
    year, month, day, hour, minute, second = match.groups()
    return f'{year}-{month}-{day}-{hour}-{minute}-{second}'
