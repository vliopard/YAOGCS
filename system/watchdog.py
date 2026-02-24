import ctypes

from utils.utils import line_number
from utils.utils import print_display


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
