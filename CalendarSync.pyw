# CalendarSync.pyw - Run with pythonw.exe (no console window)
# Dependencies: pip install pystray pillow
import ctypes
import json
import logging
import os
import queue
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import scrolledtext

import pystray
import pythoncom
from PIL import Image
from PIL import ImageDraw
from pystray import MenuItem as Item

import system.constants as constants
from system.sync_tasks import SyncTask
from system.tools import print_display
from system.tools import line_number

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

constants.RUN_GUI = True

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
INTERVAL_OBSERVER = 280  # 5 minutes in seconds
INTERVAL_SYNC_JOB = 1800  # 30 minutes in seconds
ANIM_FRAMES = 8  # animation frames for spinning arc

APP_NAME = 'CalendarSync'
APP_VERSION = '1.0.0'
APP_DESCRIPTION = 'Automatically syncs your calendar events\nin the background at scheduled intervals.'
APP_AUTHOR = 'Vincent Liopard.'
APP_COMPANY = 'OTDS H Co.'

# ---------------------------------------------------------------------------
# Global Tkinter root (created on main thread)
# ---------------------------------------------------------------------------
root = tk.Tk()
root.withdraw()

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------
paused = threading.Event()
paused.set()  # set = NOT paused (running)
stop_event = threading.Event()

# ---------------------------------------------------------------------------
# Logger
# ---------------------------------------------------------------------------
log_lines = []
log_callbacks = []


class ListHandler(logging.Handler):
    def emit(self,
             record):
        line = self.format(record)
        log_lines.append(line)
        for call_back in list(log_callbacks):
            try:
                call_back(line)
            except Exception:
                pass


class PauseToken:
    class Interrupted(Exception):
        pass

    def __init__(self,
                 event: threading.Event):
        self._event = event

    def check(self):
        if not self._event.is_set():
            self._event.wait()
            if not self._event.is_set():
                raise PauseToken.Interrupted()

    @property
    def is_paused(self):
        return not self._event.is_set()


class SystemObserver:
    def __init__(self):
        self.enabled = True
        self.continuous = 0x80000000
        self.system_required = 0x00000001
        self.display_required = 0x00000002

    def system_observer_state(self):
        if self.enabled:
            print_display(f'{line_number()} System observing state...')
            ctypes.windll.kernel32.SetThreadExecutionState(self.continuous | self.system_required | self.display_required)

    def system_original_state(self):
        if self.enabled:
            print_display(f'{line_number()} System continuous system state...')
            ctypes.windll.kernel32.SetThreadExecutionState(self.continuous)


logger = logging.getLogger('CalendarSync Logger')
logger.setLevel(logging.DEBUG)
_handler = ListHandler()
_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s',
                                        '%H:%M:%S'))
logger.addHandler(_handler)

# ---------------------------------------------------------------------------
# Log Viewer Memory (Persistent)
# ---------------------------------------------------------------------------
SETTINGS_FILE = str((Path(__file__).resolve().parent / 'resources' / 'database' / 'settings.json').resolve())

_log_win = None


def save_settings():
    """Writes the window position to a physical file."""
    global _viewer_geom
    try:
        # Create the directory if it doesn't exist
        os.makedirs(os.path.dirname(SETTINGS_FILE),
                    exist_ok=True)
        with open(SETTINGS_FILE,
                  "w") as f:
            json.dump({
                    "viewer_geom": _viewer_geom},
                    f)
    except Exception:
        pass


def load_settings():
    """Reads the window position from the physical file on startup."""
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE,
                      "r") as f:
                data = json.load(f)
                # Use .get() with a fallback default string
                return data.get("viewer_geom",
                                '700x400')
        except Exception:
            pass
    return '700x400'  # Default if no file exists


# Load the memory immediately when the script starts
_viewer_geom = load_settings()


# ---------------------------------------------------------------------------
# Pause-aware helpers
# ---------------------------------------------------------------------------
def interruptible_sleep(seconds: float,
                        interval: float = 0.5):
    """Sleep for `seconds`, honoring pause and stop at each interval tick."""
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if stop_event.is_set():
            raise StopIteration('stop requested')
        if not paused.is_set():
            paused.wait()
            deadline = time.monotonic() + (deadline - time.monotonic())
        remaining = deadline - time.monotonic()
        time.sleep(min(interval,
                       max(remaining,
                           0)))


def check_pause():
    """Block if paused; raise StopIteration if stop was requested."""
    if stop_event.is_set():
        raise StopIteration('stop requested')
    if not paused.is_set():
        paused.wait()
    if stop_event.is_set():
        raise StopIteration('stop requested')


# ---------------------------------------------------------------------------
# Icon image helpers
# ---------------------------------------------------------------------------
def _base_img():
    return Image.new('RGBA',
                     (64,
                      64),
                     (0,
                      0,
                      0,
                      0))


def make_icon_green():
    img = _base_img()
    ImageDraw.Draw(img).ellipse([4,
                                 4,
                                 60,
                                 60],
                                fill=(0,
                                      200,
                                      80))
    return img


def make_icon_red():
    img = _base_img()
    ImageDraw.Draw(img).ellipse([4,
                                 4,
                                 60,
                                 60],
                                fill=(220,
                                      40,
                                      40))
    return img


def make_icon_anim(frame: int):
    img = _base_img()
    d = ImageDraw.Draw(img)
    d.ellipse([4,
               4,
               60,
               60],
              fill=(30,
                    120,
                    255))
    start = int(frame * (360 / ANIM_FRAMES))
    d.arc([8,
           8,
           56,
           56],
          start=start,
          end=start + 120,
          fill=(255,
                255,
                255),
          width=6)
    return img


# ---------------------------------------------------------------------------
# Icon manager — the ONE AND ONLY thread that writes tray.icon
#
# All other threads call _set_icon_state(state) to request a change.
# Valid states:  'animate' | 'pause' | 'idle' | 'stop'
# Priority rule: pause > animate > idle
# ---------------------------------------------------------------------------
icon_queue = queue.Queue()
_active_jobs = 0
_active_jobs_lock = threading.Lock()


def _icon_manager(tray):
    frame = 0
    state = 'idle'
    tray.icon = make_icon_green()

    while True:
        # Drain queue — keep only the latest command
        try:
            while True:
                state = icon_queue.get_nowait()
        except queue.Empty:
            pass

        if state == 'stop':
            break
        elif state == 'pause':
            tray.icon = make_icon_red()
        elif state == 'animate':
            tray.icon = make_icon_anim(frame % ANIM_FRAMES)
            frame += 1
        else:  # 'idle'
            tray.icon = make_icon_green()

        time.sleep(0.15)


def _set_icon_state(state: str):
    """Post an icon state command. Safe to call from any thread."""
    icon_queue.put(state)


# ---------------------------------------------------------------------------
# Job wrapper — increments active counter, drives icon state
# ---------------------------------------------------------------------------
def _job_wrapper(func,
                 *args,
                 **kwargs):
    global _active_jobs
    with _active_jobs_lock:
        _active_jobs += 1
        _set_icon_state('animate')
    try:
        func(*args,
             **kwargs)
    finally:
        with _active_jobs_lock:
            _active_jobs -= 1
            if _active_jobs == 0:
                _set_icon_state('pause' if not paused.is_set() else 'idle')


# ---------------------------------------------------------------------------
# Worker functions — replace bodies with real logic
# ---------------------------------------------------------------------------
def function_observer():
    logger.info('[Observer] started')
    system_observer = SystemObserver()
    try:
        check_pause()
        system_observer.system_observer_state()
        interruptible_sleep(3)
    except StopIteration:
        logger.warning('[Observer] interrupted')
        system_observer.system_original_state()
    logger.info('[Observer] Cycled')


def function_sync_job():
    logger.info('[Sync Job] started')
    pythoncom.CoInitialize()
    try:
        check_pause()
        sync_task = SyncTask()
        sync_task.sync_task()
        interruptible_sleep(4)
    except StopIteration:
        logger.warning('[Sync Job] interrupted')
    finally:
        # Always uninitialize COM to clean up resources
        pythoncom.CoUninitialize()
    logger.info('[Sync Job] Cycled')


# ---------------------------------------------------------------------------
# Scheduler — lightweight loop, never blocks on actual work
# ---------------------------------------------------------------------------
def main_loop():
    last_observer = 0.0
    last_sync_job = 0.0
    running_observer = threading.Event()
    running_sync_job = threading.Event()

    def run_observer():
        running_observer.set()
        _job_wrapper(function_observer)
        running_observer.clear()

    def run_sync_job():
        running_sync_job.set()
        _job_wrapper(function_sync_job)
        running_sync_job.clear()

    while not stop_event.is_set():
        # Pause handling
        if not paused.is_set():
            _set_icon_state('pause')
            paused.wait()
            _set_icon_state('animate' if _active_jobs > 0 else 'idle')

        now = time.monotonic()

        # Stage 1 — Observer Task every 5 minutes
        if now - last_observer >= INTERVAL_OBSERVER:
            last_observer = time.monotonic()
            if not running_observer.is_set():
                logger.info('Scheduling [Observer]...')
                threading.Thread(target=run_observer,
                                 daemon=True).start()
                # Log next Observer run
                next_obs = datetime.fromtimestamp(time.time() + INTERVAL_OBSERVER).strftime('%Y.%m.%d %p %I:%M:%S')
                logger.info(f'Next [[OBSERVER]] scheduled for [{next_obs}]')
            else:
                logger.warning('Observer Task still running — skipping this cycle')

        # Stage 2 — Sync Job every 30 minutes
        if now - last_sync_job >= INTERVAL_SYNC_JOB:
            last_sync_job = time.monotonic()
            if not running_sync_job.is_set():
                logger.info('Scheduling [Sync Job]...')
                threading.Thread(target=run_sync_job,
                                 daemon=True).start()
                next_sync = datetime.fromtimestamp(time.time() + INTERVAL_SYNC_JOB).strftime('%Y.%m.%d %p %I:%M:%S')
                logger.info(f'Next [[SYNC JOB]] scheduled for [{next_sync}]')
            else:
                logger.warning('Sync Job still running — skipping this cycle')

        try:
            interruptible_sleep(10)
        except StopIteration:
            break


# ---------------------------------------------------------------------------
# Log Viewer
# ---------------------------------------------------------------------------

def open_log_viewer():
    root.after(0,
               _create_or_raise_log_viewer)


def _create_or_raise_log_viewer():
    global _log_win, _viewer_geom

    if _log_win is not None:
        try:
            if _log_win.winfo_exists():
                _log_win.lift()
                _log_win.focus_force()
                return
        except Exception:
            pass

    log_viewer_window = tk.Toplevel(root)
    log_viewer_window.title('Log Viewer')
    log_viewer_window.geometry(_viewer_geom if _viewer_geom else '700x400')
    log_viewer_window.protocol('WM_DELETE_WINDOW',
                               lambda: _on_log_close(log_viewer_window))

    log_viewer_window.autoscroll_enabled = tk.BooleanVar(value=True)

    ctrl_frame = tk.Frame(log_viewer_window,
                          bg='#2d2d2d')
    ctrl_frame.pack(fill=tk.X,
                    side=tk.TOP)

    def toggle_scroll():
        current = log_viewer_window.autoscroll_enabled.get()
        log_viewer_window.autoscroll_enabled.set(not current)
        btn_text = "Auto-Scroll: ON" if not current else "Auto-Scroll: OFF"
        scroll_btn.config(text=btn_text,
                          fg="#00c850" if not current else "#8888aa")

    scroll_btn = tk.Button(ctrl_frame,
                           text="Auto-Scroll: ON",
                           font=('Segoe UI',
                                 8),
                           bg='#2d2d2d',
                           fg='#00c850',
                           relief='flat',
                           command=toggle_scroll,
                           cursor='hand2')
    scroll_btn.pack(side=tk.LEFT,
                    padx=5,
                    pady=2)

    clear_btn = tk.Button(ctrl_frame,
                          text="Clear Log",
                          font=('Segoe UI',
                                8),
                          bg='#2d2d2d',
                          fg='#d4d4d4',
                          relief='flat',
                          command=lambda: on_clear_log(None,
                                                       None),
                          cursor='hand2')
    clear_btn.pack(side=tk.LEFT,
                   padx=5,
                   pady=2)

    txt = scrolledtext.ScrolledText(log_viewer_window,
                                    state='disabled',
                                    wrap='word',
                                    font=('Consolas',
                                          8),
                                    bg='#1e1e1e',
                                    fg='#d4d4d4',
                                    insertbackground='white',
                                    borderwidth=0)
    txt.pack(fill=tk.BOTH,
             expand=True,
             padx=4,
             pady=4)

    _append_to_text(txt,
                    log_lines,
                    autoscroll=True)

    def _safe_append(local_window,
                     local_text,
                     local_line):
        if not local_window.winfo_exists() or not local_text.winfo_exists():
            return
        _append_to_text(local_text,
                        [local_line],
                        local_window.autoscroll_enabled.get())

    def on_new_line(line):
        if not log_viewer_window.winfo_exists():
            return
        log_viewer_window.after(0,
                                lambda: _safe_append(log_viewer_window,
                                                     txt,
                                                     line))

    log_callbacks.append(on_new_line)
    log_viewer_window._log_cb = on_new_line
    log_viewer_window._log_txt = txt
    _log_win = log_viewer_window


def _on_log_close(win):
    global _log_win, _viewer_geom
    _viewer_geom = win.geometry()
    save_settings()
    try:
        log_callbacks.remove(win._log_cb)
    except ValueError:
        pass
    win.destroy()
    _log_win = None


def _append_to_text(txt_widget,
                    lines,
                    autoscroll=True):
    # Se o widget foi destruído, não faça nada
    if not txt_widget.winfo_exists():
        return

    try:
        txt_widget.configure(state='normal')
        for line in lines:
            txt_widget.insert(tk.END,
                              f'{line}\n')
        txt_widget.configure(state='disabled')

        if autoscroll:
            txt_widget.see(tk.END)
    except tk.TclError:
        # O widget morreu entre o exists() e o insert()
        pass


def _clear_text_widget(txt_widget):
    txt_widget.configure(state='normal')
    txt_widget.delete('1.0',
                      tk.END)
    txt_widget.configure(state='disabled')


# ---------------------------------------------------------------------------
# About Window
# ---------------------------------------------------------------------------
_about_win = None


def open_about():
    root.after(0,
               _create_or_raise_about)


def _create_or_raise_about():
    global _about_win

    if _about_win is not None:
        try:
            if _about_win.winfo_exists():
                _about_win.lift()
                _about_win.focus_force()
                return
        except Exception:
            pass

    background_color = '#1a1a2e'
    card_color = '#16213e'
    accent_color = '#0f3460'
    green_color = '#00c850'
    foreground_bright = '#e8e8f0'
    foreground_dim = '#8888aa'
    border_color = '#0f3460'

    about_window = tk.Toplevel(root)
    about_window.title(f'About {APP_NAME}')
    about_window.resizable(False,
                           False)

    w, h = 400, 400
    sw = about_window.winfo_screenwidth()
    sh = about_window.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    about_window.geometry(f'{w}x{h}+{x}+{y}')
    about_window.configure(bg=background_color)
    about_window.protocol('WM_DELETE_WINDOW',
                          lambda: _on_about_close(about_window))

    # ── top accent bar ──────────────────────────────────────────────────────
    accent_bar = tk.Frame(about_window,
                          bg=green_color,
                          height=4)
    accent_bar.pack(fill=tk.X,
                    side=tk.TOP)

    # ── icon + name row ─────────────────────────────────────────────────────
    header_frame = tk.Frame(about_window,
                            bg=background_color)
    header_frame.pack(fill=tk.X,
                      padx=28,
                      pady=(22,
                            0))

    # Small canvas icon (green dot mirroring the tray icon)
    icon_canvas = tk.Canvas(header_frame,
                            width=42,
                            height=42,
                            bg=background_color,
                            highlightthickness=0)
    icon_canvas.pack(side=tk.LEFT,
                     padx=(0,
                           14))
    icon_canvas.create_oval(3,
                            3,
                            39,
                            39,
                            fill=green_color,
                            outline='')

    name_frame = tk.Frame(header_frame,
                          bg=background_color)
    name_frame.pack(side=tk.LEFT,
                    anchor='w')

    tk.Label(name_frame,
             text=APP_NAME,
             font=('Segoe UI',
                   18,
                   'bold'),
             bg=background_color,
             fg=foreground_bright).pack(anchor='w')

    tk.Label(name_frame,
             text=f'Version {APP_VERSION}',
             font=('Segoe UI',
                   9),
             bg=background_color,
             fg=foreground_dim).pack(anchor='w')

    # ── divider ─────────────────────────────────────────────────────────────
    div = tk.Frame(about_window,
                   bg=border_color,
                   height=1)
    div.pack(fill=tk.X,
             padx=28,
             pady=(18,
                   0))

    # ── description card ────────────────────────────────────────────────────
    card = tk.Frame(about_window,
                    bg=card_color,
                    bd=0,
                    highlightthickness=1,
                    highlightbackground=accent_color)
    card.pack(fill=tk.X,
              padx=28,
              pady=(16,
                    0))

    tk.Label(card,
             text=APP_DESCRIPTION,
             font=('Segoe UI',
                   10),
             bg=card_color,
             fg=foreground_bright,
             justify=tk.LEFT,
             wraplength=320).pack(anchor='w',
                                  padx=14,
                                  pady=12)

    # ── metadata grid ───────────────────────────────────────────────────────
    meta_frame = tk.Frame(about_window,
                          bg=background_color)
    meta_frame.pack(fill=tk.X,
                    padx=28,
                    pady=(16,
                          0))

    def _meta_row(parent,
                  label,
                  value):
        row = tk.Frame(parent,
                       bg=background_color)
        row.pack(fill=tk.X,
                 pady=2)
        tk.Label(row,
                 text=label,
                 font=('Segoe UI',
                       9),
                 bg=background_color,
                 fg=foreground_dim,
                 width=14,
                 anchor='w').pack(side=tk.LEFT)
        tk.Label(row,
                 text=value,
                 font=('Segoe UI',
                       9),
                 bg=background_color,
                 fg=foreground_bright,
                 anchor='w').pack(side=tk.LEFT)

    _meta_row(meta_frame,
              'Author',
              APP_AUTHOR)
    _meta_row(meta_frame,
              'Company',
              APP_COMPANY)
    _meta_row(meta_frame,
              'Observer interval',
              f'Every {INTERVAL_OBSERVER // 60} min')
    _meta_row(meta_frame,
              'Sync interval',
              f'Every {INTERVAL_SYNC_JOB // 60} min')

    # ── close button ────────────────────────────────────────────────────────
    btn_frame = tk.Frame(about_window,
                         bg=background_color)
    btn_frame.pack(pady=(18,
                         0))

    close_btn = tk.Button(btn_frame,
                          text='Close',
                          font=('Segoe UI',
                                9),
                          bg=accent_color,
                          fg=foreground_bright,
                          activebackground=green_color,
                          activeforeground='#000000',
                          relief='flat',
                          cursor='hand2',
                          padx=24,
                          pady=6,
                          bd=0,
                          command=lambda: _on_about_close(about_window))
    close_btn.pack()

    _about_win = about_window


def _on_about_close(win):
    global _about_win
    win.destroy()
    _about_win = None


# ---------------------------------------------------------------------------
# Tray menu callbacks
# ---------------------------------------------------------------------------
def on_pause_resume(icon,
                    item):
    if paused.is_set():
        paused.clear()
        _set_icon_state('pause')
        logger.info('Paused by user')
    else:
        paused.set()
        logger.info('Resumed by user')
        _set_icon_state('animate' if _active_jobs > 0 else 'idle')


def is_paused(item):
    return not paused.is_set()


def on_log_viewer(icon,
                  item):
    open_log_viewer()


def on_clear_log(icon,
                 item):
    log_lines.clear()
    if _log_win is not None:
        try:
            if _log_win.winfo_exists():
                root.after(0,
                           _clear_text_widget,
                           _log_win._log_txt)
        except Exception:
            pass


def on_about(icon,
             item):
    open_about()


def on_quit(icon,
            item):
    stop_event.set()
    paused.set()
    _set_icon_state('stop')
    icon.stop()
    root.after(0,
               root.destroy)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    menu = pystray.Menu(Item('Paused',
                             on_pause_resume,
                             checked=is_paused),
                        Item('Log Viewer',
                             on_log_viewer),
                        Item('Clear Log',
                             on_clear_log),
                        pystray.Menu.SEPARATOR,
                        Item('About',
                             on_about),
                        pystray.Menu.SEPARATOR,
                        Item('Quit',
                             on_quit))

    icon = pystray.Icon(name='CalendarSync',
                        icon=make_icon_green(),
                        title='CalendarSync',
                        menu=menu)

    # Icon manager is the ONLY thread that writes tray.icon
    threading.Thread(target=_icon_manager,
                     args=(icon,),
                     daemon=True).start()

    # Scheduler
    threading.Thread(target=main_loop,
                     daemon=True).start()

    # pystray on its own thread; Tkinter owns the main thread
    threading.Thread(target=icon.run,
                     daemon=True).start()

    logger.info('CalendarSync Started...')
    root.mainloop()


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print_display(f'{line_number()} Bye...')
        pass
