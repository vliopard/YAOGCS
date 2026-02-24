import calendar
import re
from datetime import datetime
from datetime import timedelta
from datetime import timezone

from dateutil import parser

from utils.utils import line_number
from utils.utils import print_display

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
_OUTLOOK_MAX_DATE = datetime(4500,
                             1,
                             1,
                             tzinfo=timezone.utc)


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


def utc_now():
    return datetime.now(timezone.utc).replace(tzinfo=None).isoformat()


def convert_to_str(date_time_value):
    if isinstance(date_time_value,
                  dict):
        date_time_value = date_time_value.get('dateTime',
                                              '')
    if isinstance(date_time_value,
                  datetime):
        date_time_value = date_time_value.isoformat()
    return date_time_value


def time_when(shift):
    utc_time_now = datetime.now(timezone.utc)
    if shift < 0:
        utc_time_begin = utc_time_now - timedelta(days=abs(shift))
    elif shift > 0:
        utc_time_begin = utc_time_now + timedelta(days=shift)
    else:
        utc_time_begin = utc_time_now
    return utc_time_begin.isoformat().replace('+00:00',
                                              'Z')


def time_min():
    return time_when(-18)


def time_max():
    return time_when(180)


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
