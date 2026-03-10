# system/settings_screen.py
import json
import os
import tkinter as tk
from pathlib import Path

import system.constants as constants

SETTINGS_FILE = str((Path(__file__).resolve().parent.parent / 'resources' / 'database' / 'settings.json').resolve())

# Keys managed by this screen
_SETTING_KEYS = ('day_past',
                 'day_next',
                 'interval_observer',
                 'interval_sync_job',
                 'settings_geom')

_settings_win = None

screen_size = '450x420'

# ---------------------------------------------------------------------------
# Persistence helpers
# ---------------------------------------------------------------------------

def load_runtime_settings():
    """Load persisted settings and apply them to constants at startup."""
    if not os.path.exists(SETTINGS_FILE):
        return
    try:
        with open(SETTINGS_FILE,
                  'r') as f:
            data = json.load(f)
        if 'day_past' in data:
            constants.DAY_PAST = int(data['day_past'])
        if 'day_next' in data:
            constants.DAY_NEXT = int(data['day_next'])
        if 'interval_observer' in data:
            constants.INTERVAL_OBSERVER = int(data['interval_observer'])
        if 'interval_sync_job' in data:
            constants.INTERVAL_SYNC_JOB = int(data['interval_sync_job'])
    except Exception:
        pass


def _save_runtime_settings(extra: dict | None = None):
    """Merge runtime settings into the shared settings.json file."""
    try:
        os.makedirs(os.path.dirname(SETTINGS_FILE),
                    exist_ok=True)
        existing = {}
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE,
                      'r') as f:
                existing = json.load(f)
        existing.update({
                'day_past'         : constants.DAY_PAST,
                'day_next'         : constants.DAY_NEXT,
                'interval_observer': constants.INTERVAL_OBSERVER,
                'interval_sync_job': constants.INTERVAL_SYNC_JOB, })
        if extra:
            existing.update(extra)
        with open(SETTINGS_FILE,
                  'w') as f:
            json.dump(existing,
                      f,
                      indent=2)
    except Exception:
        pass


def _load_settings_geom() -> str:
    if not os.path.exists(SETTINGS_FILE):
        return screen_size
    try:
        with open(SETTINGS_FILE,
                  'r') as f:
            data = json.load(f)
        return data.get('settings_geom',
                        screen_size)
    except Exception:
        return screen_size


# ---------------------------------------------------------------------------
# Public entry point (mirrors open_log_viewer pattern)
# ---------------------------------------------------------------------------

def open_settings(root: tk.Tk):
    root.after(0,
               lambda: _create_or_raise_settings(root))


def _create_or_raise_settings(root: tk.Tk):
    global _settings_win

    if _settings_win is not None:
        try:
            if _settings_win.winfo_exists():
                _settings_win.lift()
                _settings_win.focus_force()
                return
        except Exception:
            pass

    # ── colour palette (matches About window) ──────────────────────────────
    bg = '#1a1a2e'
    card = '#16213e'
    accent = '#0f3460'
    green = '#00c850'
    fg_bright = '#e8e8f0'
    fg_dim = '#8888aa'
    entry_bg = '#0d1117'
    error_color = '#ff4444'

    geom = _load_settings_geom()

    win = tk.Toplevel(root)
    win.title('Settings')
    win.geometry(geom)
    win.resizable(False,
                  False)
    win.configure(bg=bg)
    win.protocol('WM_DELETE_WINDOW',
                 lambda: _on_settings_close(win))
    _settings_win = win

    # ── top accent bar ──────────────────────────────────────────────────────
    tk.Frame(win,
             bg=green,
             height=4).pack(fill=tk.X,
                            side=tk.TOP)

    # ── title ───────────────────────────────────────────────────────────────
    hdr = tk.Frame(win,
                   bg=bg)
    hdr.pack(fill=tk.X,
             padx=24,
             pady=(18,
                   0))
    tk.Label(hdr,
             text='Settings',
             font=('Segoe UI',
                   16,
                   'bold'),
             bg=bg,
             fg=fg_bright).pack(side=tk.LEFT)
    tk.Label(hdr,
             text='Live — applied immediately',
             font=('Segoe UI',
                   8),
             bg=bg,
             fg=fg_dim).pack(side=tk.LEFT,
                             padx=(10,
                                   0),
                             pady=(6,
                                   0))

    # ── divider ─────────────────────────────────────────────────────────────
    tk.Frame(win,
             bg=accent,
             height=1).pack(fill=tk.X,
                            padx=24,
                            pady=(12,
                                  0))

    # ── card ────────────────────────────────────────────────────────────────
    card_frame = tk.Frame(win,
                          bg=card,
                          bd=0,
                          highlightthickness=1,
                          highlightbackground=accent)
    card_frame.pack(fill=tk.X,
                    padx=24,
                    pady=(14,
                          0))

    entries: dict[str, tk.Entry] = {}
    error_labels: dict[str, tk.Label] = {}

    fields = [('day_past',
               'Day Past',
               'days',
               lambda v: v > 0),
            ('day_next',
             'Day Next',
             'days',
             lambda v: v > 0),
            ('interval_observer',
             'Observer Interval',
             'seconds',
             lambda v: v >= 10),
            ('interval_sync_job',
             'Sync Job Interval',
             'seconds',
             lambda v: v >= 60), ]

    def _current_value(key: str) -> int:
        return {
                'day_past'         : constants.DAY_PAST,
                'day_next'         : constants.DAY_NEXT,
                'interval_observer': constants.INTERVAL_OBSERVER,
                'interval_sync_job': constants.INTERVAL_SYNC_JOB}[key]

    grid = tk.Frame(card_frame,
                    bg=card)
    grid.pack(fill=tk.X,
              padx=14,
              pady=(10,
                    0))
    grid.columnconfigure(0,
                         minsize=160)  # label column — fixed width, no clipping
    grid.columnconfigure(1,
                         minsize=100)  # entry column
    grid.columnconfigure(2,
                         weight=1)  # error label — takes remaining space

    for row_idx, (key,
                  label,
                  unit,
                  _) in enumerate(fields):
        pad_y = (4,
                 4)

        # label + unit stacked in a plain Frame
        lbl_cell = tk.Frame(grid,
                            bg=card)
        lbl_cell.grid(row=row_idx,
                      column=0,
                      sticky='w',
                      padx=(0,
                            12),
                      pady=pad_y)
        tk.Label(lbl_cell,
                 text=label,
                 font=('Segoe UI',
                       9),
                 bg=card,
                 fg=fg_bright,
                 anchor='w').pack(anchor='w')
        tk.Label(lbl_cell,
                 text=f'({unit})',
                 font=('Segoe UI',
                       7),
                 bg=card,
                 fg=fg_dim,
                 anchor='w').pack(anchor='w')

        # entry — centered vertically in its row using sticky='ns' on a container
        ent_cell = tk.Frame(grid,
                            bg=card)
        ent_cell.grid(row=row_idx,
                      column=1,
                      sticky='nsw',
                      pady=pad_y)
        var = tk.StringVar(value=str(_current_value(key)))
        ent = tk.Entry(ent_cell,
                       textvariable=var,
                       width=10,
                       bg=entry_bg,
                       fg=fg_bright,
                       insertbackground=fg_bright,
                       relief='flat',
                       font=('Consolas',
                             9),
                       highlightthickness=1,
                       highlightbackground=accent,
                       highlightcolor=green)
        ent.place(relx=0,
                  rely=0.5,
                  anchor='w')  # vertically center inside cell
        ent_cell.update_idletasks()
        ent_cell.config(width=ent.winfo_reqwidth(),
                        height=lbl_cell.winfo_reqheight())
        entries[key] = ent

        # inline error label
        err_lbl = tk.Label(grid,
                           text='',
                           font=('Segoe UI',
                                 8),
                           bg=card,
                           fg=error_color,
                           anchor='w')
        err_lbl.grid(row=row_idx,
                     column=2,
                     sticky='w',
                     pady=pad_y)
        error_labels[key] = err_lbl

    # spacer at bottom of card
    tk.Frame(card_frame,
             bg=card,
             height=4).pack()

    # ── hint label ──────────────────────────────────────────────────────────
    hint_var = tk.StringVar(value='')
    hint_lbl = tk.Label(win,
                        textvariable=hint_var,
                        font=('Segoe UI',
                              8),
                        bg=bg,
                        fg=fg_dim)
    hint_lbl.pack(pady=(8,
                        0))

    # ── buttons ─────────────────────────────────────────────────────────────
    btn_frame = tk.Frame(win,
                         bg=bg)
    btn_frame.pack(pady=(10,
                         16))

    def _apply():
        validators = {f[0]: f[3] for f in fields}
        all_ok = True
        new_vals: dict[str, int] = {}

        for key, ent in entries.items():
            raw = ent.get().strip()
            try:
                val = int(raw)
                if not validators[key](val):
                    raise ValueError('out of range')
                error_labels[key].config(text='')
                new_vals[key] = val
            except ValueError:
                error_labels[key].config(text='✗ invalid')
                all_ok = False

        if not all_ok:
            hint_var.set('Fix errors above before applying.')
            return

        # Apply to live constants
        constants.DAY_PAST = new_vals['day_past']
        constants.DAY_NEXT = new_vals['day_next']
        constants.INTERVAL_OBSERVER = new_vals['interval_observer']
        constants.INTERVAL_SYNC_JOB = new_vals['interval_sync_job']

        _save_runtime_settings()
        hint_var.set('✓ Settings applied and saved.')
        win.after(2500,
                  lambda: hint_var.set(''))

    def _reset():
        defaults = {
                'day_past'         : 18,
                'day_next'         : 180,
                'interval_observer': 280,
                'interval_sync_job': 60 * 60 * 2}
        for key, ent in entries.items():
            ent.delete(0,
                       tk.END)
            ent.insert(0,
                       str(defaults[key]))
            error_labels[key].config(text='')
        hint_var.set('Defaults loaded — press Apply to save.')

    _btn_style = dict(font=('Segoe UI',
                            9),
                      relief='flat',
                      cursor='hand2',
                      padx=20,
                      pady=5,
                      bd=0)

    tk.Button(btn_frame,
              text='Apply',
              bg=green,
              fg='#000000',
              activebackground='#00ff6a',
              activeforeground='#000000',
              command=_apply,
              **_btn_style).pack(side=tk.LEFT,
                                 padx=6)

    tk.Button(btn_frame,
              text='Reset Defaults',
              bg=accent,
              fg=fg_bright,
              activebackground='#1a4a80',
              activeforeground=fg_bright,
              command=_reset,
              **_btn_style).pack(side=tk.LEFT,
                                 padx=6)

    tk.Button(btn_frame,
              text='Close',
              bg='#2d2d2d',
              fg=fg_dim,
              activebackground='#3d3d3d',
              activeforeground=fg_bright,
              command=lambda: _on_settings_close(win),
              **_btn_style).pack(side=tk.LEFT,
                                 padx=6)


def _on_settings_close(win: tk.Toplevel):
    global _settings_win
    geom = win.geometry()
    _save_runtime_settings(extra={
            'settings_geom': geom})
    win.destroy()
    _settings_win = None
