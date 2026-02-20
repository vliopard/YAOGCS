import ctypes
import os
import sys
import traceback

_base = os.path.dirname(os.path.abspath(__file__))
_log_path = os.path.join(_base,
                         'launch_errors.log')
_log_file = open(_log_path,
                 'a',
                 encoding='utf-8')
sys.stdout = _log_file
sys.stderr = _log_file

sys.path.insert(0,
                _base)

_MUTEX_NAME = 'CalendarSync_SingleInstance'
_mutex = ctypes.windll.kernel32.CreateMutexW(None,
                                             False,
                                             _MUTEX_NAME)
if ctypes.windll.kernel32.GetLastError() == 183:
    ctypes.windll.kernel32.CloseHandle(_mutex)
    sys.exit(0)

try:
    import utils.constants as constants
    from system.gui import main_gui, _log_queue
    from utils.utils import set_log_queue

    set_log_queue(_log_queue)
    constants.RUN_GUI = True
    main_gui()
except Exception:
    traceback.print_exc()
    _log_file.flush()
