import sys

import filelock

from .gui.gui import gui_main
from .logging_ import setup_logging
from .utils import create_app_dirs, get_file_lock, get_logs_dir


# Acquire file lock to limit the program to a single instance
try:
    lock = get_file_lock()
    lock.acquire()
except filelock.Timeout:
    print('Another instance is already running')
    sys.exit(1)


# Create application directories and setup logging early
try:
    create_app_dirs()
except Exception:
    import logging
    from pathlib import Path

    setup_logging(Path('.'))

    log = logging.getLogger()
    log.exception('Failed to create application directories')

    sys.exit(1)

logs_dir = get_logs_dir()
setup_logging(logs_dir)


gui_main()
