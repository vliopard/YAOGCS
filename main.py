import os
import sys

import utils.constants as constants
from system.gui import _log_queue
from system.gui import main_gui
from utils.utils import set_log_queue

if __name__ == '__main__':
    set_log_queue(_log_queue)
    constants.RUN_GUI = True
    sys.stdout = open(os.devnull,
                      'w')
    sys.stderr = open(os.devnull,
                      'w')
    main_gui()
