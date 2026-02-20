import ctypes
import json
import logging
import queue
import ssl
import socket
import threading
import time
from datetime import datetime
from pathlib import Path
from sys import stdin as sys_standard_in
from time import sleep

import pystray
import pythoncom
import wx
from PIL import Image

import utils.constants as constants
from connector.g_calendar import GoogleCalendarConnector
from connector.ms_outlook import MicrosoftOutlookConnector
from system.observer import SystemObserver
from system.routines import sync_outlook_to_google
from utils.utils import is_windows
from utils.utils import line_number
from utils.utils import print_debug
from utils.utils import print_display
from utils.utils import PauseToken

thread_started = time.time()
system_tray_icon = None
_log_queue = queue.Queue()
_tray_queue = queue.Queue()
_pause_event = threading.Event()
_pause_event.set()  # starts in running state
wx_app_ref = [None]

if is_windows():
    import msvcrt
else:
    import termios
    import atexit
    from select import select


class KBHit:
    def __init__(self):
        if is_windows():
            pass
        else:
            self.fd = sys_standard_in.fileno()
            self.new_term = termios.tcgetattr(self.fd)
            self.old_term = termios.tcgetattr(self.fd)
            self.new_term[3] = (self.new_term[3] & ~termios.ICANON & ~termios.ECHO)
            termios.tcsetattr(self.fd,
                              termios.TCSAFLUSH,
                              self.new_term)
            atexit.register(self.set_normal_term)

    def set_normal_term(self):
        if is_windows():
            pass
        else:
            termios.tcsetattr(self.fd,
                              termios.TCSAFLUSH,
                              self.old_term)

    @staticmethod
    def get_character():
        if is_windows():
            return msvcrt.getch().decode(constants.UTF8)
        else:
            return sys_standard_in.read(1)

    @staticmethod
    def get_arrow_key():
        if is_windows():
            msvcrt.getch()
            character = msvcrt.getch()
            arrows = [72,
                      77,
                      80,
                      75]
        else:
            character = sys_standard_in.read(3)[2]
            arrows = [65,
                      67,
                      66,
                      68]
        return arrows.index(ord(character.decode(constants.UTF8)))

    @staticmethod
    def keyboard_hit():
        if is_windows():
            return msvcrt.kbhit()
        else:
            dr, dw, de = select([sys_standard_in],
                                [],
                                [],
                                0)
            return dr != []

    def check(self):
        if self.keyboard_hit():
            character = self.get_character()
            if ord(character) == constants.SLASH:
                self.set_normal_term()
                return True
        return False


def update_icon(task_event):
    global thread_started
    global system_tray_icon

    paused_state = False
    check_out = False
    tray_icon_thread_running = True
    moving_icons = [constants.MOVE_ICO1,
                    constants.MOVE_ICO2,
                    constants.MOVE_ICO3,
                    constants.MOVE_ICO4]
    loop_interaction = 0

    while tray_icon_thread_running:
        time.sleep(1)

        while not task_event.empty():
            value = task_event.get()
            if value == constants.PAUSE:
                system_tray_icon.icon = Image.open(constants.ICON_DONE)
                paused_state = True
            if value == constants.CONTINUE:
                paused_state = False
            if value == constants.TERMINATE:
                tray_icon_thread_running = False
            if value == constants.RECHECK:
                check_out = True
            task_event.task_done()

        check_time_now = time.time() - thread_started
        if check_out and check_time_now > constants.MAX_SECONDS:
            system_tray_icon.icon = Image.open(constants.ICON_DONE)
            paused_state = True
            check_out = False
            with task_event.mutex:
                task_event.queue.clear()

        if not paused_state and tray_icon_thread_running:
            system_tray_icon.icon = Image.open(moving_icons[loop_interaction])
            loop_interaction = (loop_interaction + 1) % len(moving_icons)


def _observer_bridge(tray_q):
    pythoncom.CoInitialize()
    system_observer = SystemObserver()
    last_sync = 0
    _CHUNK_SECONDS = 5

    try:
        while True:
            # Block here if paused - zero CPU spin, wakes instantly on resume
            if not _pause_event.is_set():
                print_display(f'{line_number()} Sync paused, waiting for resume...')
                tray_q.put(constants.PAUSE)
                _pause_event.wait()
                tray_q.put(constants.CONTINUE)
                print_display(f'{line_number()} Sync resumed.')

            # Keep system awake
            ctypes.windll.kernel32.SetThreadExecutionState(system_observer.continuous | system_observer.system_required | system_observer.display_required)
            tray_q.put(constants.CONTINUE)

            # Sleep in small chunks so pause is detected quickly mid-wait / the loop stays responsive and logs keep flowing
            if not system_observer.first_sleep:
                tray_q.put(constants.PAUSE)
                elapsed = 0
                while elapsed < system_observer.sleep_timeout:
                    sleep(_CHUNK_SECONDS)
                    elapsed += _CHUNK_SECONDS
                    if not _pause_event.is_set():
                        break
                tray_q.put(constants.CONTINUE)
            else:
                system_observer.first_sleep = False

            # Re-check pause after sleep in case it was set mid-wait
            if not _pause_event.is_set():
                continue

            now = time.time()
            antes = datetime.fromtimestamp(now + system_observer.time_out()).strftime('%Y.%m.%d %p %I:%M:%S')
            nls = now - last_sync
            print_display(f'{line_number()} [{now}]-[{last_sync}]>=[{nls}][{system_observer.time_out()}]')

            if nls >= system_observer.time_out():
                current_time = datetime.now().strftime('%Y.%m.%d %p %I:%M:%S')
                print_display(f'[{current_time}] Syncing Outlook to Google...')
                _SYNC_MAX_RETRIES = 3
                _SYNC_BASE_DELAY = 10.0
                _TRANSIENT = (
                    ssl.SSLError,
                    ssl.SSLEOFError,
                    ConnectionResetError,
                    ConnectionAbortedError,
                    socket.timeout,
                    TimeoutError,
                    OSError,
                )
                sync_succeeded = False
                sync_delay = _SYNC_BASE_DELAY
                for sync_attempt in range(1, _SYNC_MAX_RETRIES + 1):
                    try:
                        connection_ms_outlook = MicrosoftOutlookConnector()
                        connection_g_calendar = GoogleCalendarConnector()
                        pause_token = PauseToken(_pause_event)
                        sync_outlook_to_google(connection_ms_outlook,
                                               connection_g_calendar,
                                               pause_token)
                        sync_succeeded = True
                        break
                    except PauseToken.Interrupted:
                        print_display(f'{line_number()} Sync interrupted by pause.')
                        break
                    except _TRANSIENT as net_error:
                        if sync_attempt < _SYNC_MAX_RETRIES:
                            print_display(f'{line_number()} [{current_time}] Transient network error (attempt {sync_attempt}/{_SYNC_MAX_RETRIES}): {net_error}, retrying in {sync_delay:.0f}s...')
                            sleep(sync_delay)
                            sync_delay *= 2
                        else:
                            print_display(f'{line_number()} [{current_time}] Sync failed after {_SYNC_MAX_RETRIES} attempts: {net_error}')
                    except Exception as sync_error:
                        print_display(f'{line_number()} [{current_time}] ERROR during sync: {sync_error}')
                        break
                if sync_succeeded:
                    last_sync = now

            print_display(f'[{antes}] NEXT Syncing Outlook to Google...')
    except KeyboardInterrupt:
        tray_q.put(constants.PAUSE)
        system_observer.system_original_state()
    except Exception as fatal_error:
        print_display(f'{line_number()} [FATAL] Observer loop crashed: {fatal_error}')
        tray_q.put(constants.PAUSE)
    finally:
        tray_q.put(constants.PAUSE)
        pythoncom.CoUninitialize()


def tray_icon_click(_,
                    selected_tray_item):
    global system_tray_icon
    tray_label = str(selected_tray_item)
    if tray_label == constants.LABEL_DONE:
        system_tray_icon.icon = Image.open(constants.ICON_DONE)
    elif tray_label == constants.LABEL_ERROR:
        system_tray_icon.icon = Image.open(constants.ICON_ERROR)
    elif tray_label in (constants.LABEL_PAUSE, constants.LABEL_RESUME):
        if _pause_event.is_set():
            _pause_event.clear()
            system_tray_icon.icon = Image.open(constants.ICON_PAUSE)
            print_display(f'{line_number()} Sync paused by user.')
        else:
            _pause_event.set()
            system_tray_icon.icon = Image.open(constants.ICON_DONE)
            print_display(f'{line_number()} Sync resumed by user.')
        system_tray_icon.update_menu()
    elif tray_label == 'Logs':
        system_tray_icon.icon = Image.open(constants.ICON_LOG)
        if wx_app_ref[0] is not None:
            import wx
            wx.CallAfter(wx_app_ref[0].open_log_window)
    elif tray_label == constants.LABEL_EXIT:
        _tray_queue.put(constants.TERMINATE)
        system_tray_icon.stop()
        if wx_app_ref[0] is not None:
            import wx
            wx.CallAfter(wx_app_ref[0].ExitMainLoop)


def _run_tray():
    global system_tray_icon
    state = False
    system_tray_image = Image.open(constants.ICON_DONE)
    system_tray_icon = pystray.Icon(f'{constants.LABEL_MAIN} 1',
                                    system_tray_image,
                                    constants.LABEL_MAIN,
                                    menu=pystray.Menu(pystray.MenuItem(constants.LABEL_DONE,
                                                                       tray_icon_click,
                                                                       checked=lambda item: state),
                                                      pystray.MenuItem(lambda item: constants.LABEL_PAUSE if _pause_event.is_set() else constants.LABEL_RESUME,
                                                                       tray_icon_click,
                                                                       checked=lambda item: state),
                                                      pystray.MenuItem(constants.LABEL_ERROR,
                                                                       tray_icon_click,
                                                                       checked=lambda item: state),
                                                      pystray.MenuItem('Logs',
                                                                       tray_icon_click,
                                                                       checked=lambda item: state),
                                                      pystray.MenuItem(constants.LABEL_EXIT,
                                                                       tray_icon_click,
                                                                       checked=lambda item: state)))
    tray_thread = threading.Thread(target=update_icon,
                                   args=(_tray_queue,),
                                   daemon=True,
                                   name='TrayIconThread')
    tray_thread.start()
    observer_thread = threading.Thread(target=_observer_bridge,
                                       args=(_tray_queue,),
                                       daemon=True,
                                       name='ObserverThread')
    observer_thread.start()
    system_tray_icon.run()


class WxTextCtrlHandler(logging.Handler):
    def __init__(self,
                 text_ctrl):
        super().__init__()
        self.text_ctrl = text_ctrl

    def emit(self,
             record):
        record_message = self.format(record)
        wx.CallAfter(self.text_ctrl.AppendText,
                     record_message + '\n')


_WINDOW_STATE_FILE = str((Path(__file__).resolve().parent.parent / 'resources' / 'database' / 'log_window_state.json').resolve())
_WINDOW_DEFAULT_SIZE = (800, 500)
_WINDOW_DEFAULT_POS = wx.DefaultPosition


def _load_window_state():
    try:
        with open(_WINDOW_STATE_FILE, 'r', encoding='utf-8') as f:
            state = json.load(f)
            return (state['x'], state['y']), (state['width'], state['height'])
    except Exception:
        return None, None


def _save_window_state(x, y, width, height):
    try:
        Path(_WINDOW_STATE_FILE).parent.mkdir(parents=True, exist_ok=True)
        with open(_WINDOW_STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump({'x': x, 'y': y, 'width': width, 'height': height}, f, indent=4)
    except Exception:
        pass


class LogFrame(wx.Frame):
    _POLL_INTERVAL_MS = 200

    def __init__(self,
                 log_queue,
                 history,
                 on_close_cb=None):
        saved_pos, saved_size = _load_window_state()
        pos = wx.Point(*saved_pos) if saved_pos else _WINDOW_DEFAULT_POS
        size = saved_size if saved_size else _WINDOW_DEFAULT_SIZE
        super().__init__(None,
                         title='Log Viewer',
                         pos=pos,
                         size=size)
        self._log_queue = log_queue
        self._history = history
        self._on_close_cb = on_close_cb
        wx_panel = wx.Panel(self)
        self.log_ctrl = wx.TextCtrl(wx_panel,
                                    style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2)
        self.clear_btn = wx.Button(wx_panel,
                                   label='Clear')
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        button_sizer.Add(self.clear_btn,
                         0,
                         wx.ALL,
                         5)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(self.log_ctrl,
                       1,
                       wx.EXPAND | wx.ALL,
                       5)
        main_sizer.Add(button_sizer,
                       0,
                       wx.ALIGN_RIGHT)
        wx_panel.SetSizer(main_sizer)
        self.clear_btn.Bind(wx.EVT_BUTTON,
                            self.on_clear)
        self.Bind(wx.EVT_CLOSE,
                  self._on_close)
        self._setup_logging()
        if self._history:
            self.log_ctrl.AppendText('\n'.join(self._history) + '\n')
        self._poll_timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER,
                  self._on_poll_timer,
                  self._poll_timer)
        self._poll_timer.Start(self._POLL_INTERVAL_MS)

    def _on_close(self,
                  event):
        print_debug(event)
        self._poll_timer.Stop()
        if not self.IsIconized() and not self.IsMaximized():
            pos = self.GetPosition()
            size = self.GetSize()
            _save_window_state(pos.x, pos.y, size.width, size.height)
        if self._on_close_cb:
            self._on_close_cb()
        self.Destroy()

    def _setup_logging(self):
        self.logger = logging.getLogger('WxLogger')
        self.logger.setLevel(logging.DEBUG)
        if not self.logger.handlers:
            text_control_handler = WxTextCtrlHandler(self.log_ctrl)
            log_formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s',
                                              datefmt='%Y-%m-%d %H:%M:%S')
            text_control_handler.setFormatter(log_formatter)
            self.logger.addHandler(text_control_handler)

    def _on_poll_timer(self,
                       event):
        print_debug(event)
        try:
            while True:
                log_message = self._log_queue.get_nowait()
                log_string = str(log_message)
                self._history.append(log_string)
                self.log_ctrl.AppendText(log_string + '\n')
                self._log_queue.task_done()
        except queue.Empty:
            pass

    def on_clear(self,
                 event):
        print_debug(event)
        self._history.clear()
        self.log_ctrl.Clear()


class _QueueDrainer(threading.Thread):
    def __init__(self,
                 log_queue,
                 history,
                 frame_alive_fn):
        super().__init__(daemon=True,
                         name='QueueDrainerThread')
        self._log_queue = log_queue
        self._history = history
        self._frame_alive = frame_alive_fn

    def run(self):
        while True:
            time.sleep(0.2)
            if self._frame_alive():
                continue
            try:
                while True:
                    log_message = self._log_queue.get_nowait()
                    self._history.append(str(log_message))
                    self._log_queue.task_done()
            except queue.Empty:
                pass


class LogApp(wx.App):
    def __init__(self,
                 log_queue,
                 on_window_close_cb=None):
        self._log_queue = log_queue
        self._on_window_close_cb = on_window_close_cb
        self._frame = None
        self._history = []
        super().__init__(False)

    def OnInit(self):
        self._hidden = wx.Frame(None)
        self.SetTopWindow(self._hidden)
        drainer = _QueueDrainer(self._log_queue,
                                self._history,
                                frame_alive_fn=self._frame_is_alive)
        drainer.start()
        return True

    def _frame_is_alive(self):
        try:
            return self._frame is not None and self._frame.IsShown()
        except RuntimeError:
            return False

    def _on_frame_closed(self):
        self._frame = None
        if self._on_window_close_cb:
            self._on_window_close_cb()

    def open_log_window(self):
        try:
            if self._frame and self._frame.IsShown():
                self._frame.Raise()
                return
        except RuntimeError:
            pass

        self._frame = LogFrame(log_queue=self._log_queue,
                               history=self._history,
                               on_close_cb=self._on_frame_closed)
        self._frame.Show()
        self._frame.Raise()


def _on_log_window_closed():
    if system_tray_icon is not None:
        try:
            system_tray_icon.icon = Image.open(constants.ICON_DONE)
        except Exception:
            pass


def main_gui():
    tray_thread = threading.Thread(target=_run_tray,
                                   daemon=True,
                                   name='PystrayThread')
    tray_thread.start()

    main_application = LogApp(log_queue=_log_queue,
                              on_window_close_cb=_on_log_window_closed)
    wx_app_ref[0] = main_application
    main_application.MainLoop()
    wx_app_ref[0] = None


def main_log():
    main_application = LogApp(log_queue=_log_queue)
    main_application.open_log_window()
    main_application.MainLoop()
