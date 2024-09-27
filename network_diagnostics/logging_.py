import atexit
import logging
import logging.config
import logging.handlers
import queue
import copy

import datetime


class DatetimeFormatter(logging.Formatter):

    converter = datetime.datetime.fromtimestamp

    def formatTime(self, record, datefmt=None):  # noqa: N802
        converter = self.converter(record.created).astimezone()
        if datefmt:
            ts = converter.strftime(datefmt)
        else:
            t = converter.strftime('%Y-%m-%dT%H:%M:%S')
            ts = f'{t}.{int(record.msecs):03d}'
        return ts


_logging_config = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'verbose': {
            '()': DatetimeFormatter,
            'format': '%(asctime)s %(levelname)s %(name)s: %(message)s',
            'datefmt': '%Y-%m-%dT%H:%M:%S.%f%z'
        }
    },
    'handlers': {
        'stderr': {
            'class': 'logging.StreamHandler',
            'level': 'DEBUG',
            'formatter': 'verbose',
            'stream': 'ext://sys.stderr'
        },
        'file': {
            'class': 'logging.handlers.RotatingFileHandler',
            'level': 'DEBUG',
            'formatter': 'verbose',
            'filename': '{}/log.log',
            'maxBytes': 1024 * 1024,  # 1 MiB
            'backupCount': 2
        },
        'queue_handler': {
            'class': 'logging.handlers.QueueHandler',
            'handlers': [
                'stderr',
                'file'
            ],
            'respect_handler_level': True
        }
    },
    'root': {
        'level': 'DEBUG',
        'handlers': [
            'queue_handler'
        ]
    }
}


def setup_logging(logs_dir):
    logging.getLogger('cfspeedtest').propagate = False
    logging.getLogger('requests').propagate = False
    logging.getLogger('urllib3').propagate = False

    config = copy.deepcopy(_logging_config)
    for handler in config.get('handlers', {}).values():
        for key, value in handler.items():
            if key == 'filename':
                handler[key] = value.format(logs_dir)

    logging.config.dictConfig(config)

    queue_handler = logging.getHandlerByName('queue_handler')
    if queue_handler is not None:
        queue_handler.listener.start()
        atexit.register(queue_handler.listener.stop)


_diagnostics_handlers = {
    'root': {
        'level': 'DEBUG',
        'format': '%(asctime)s %(levelname)s %(name)s: %(message)s',
        'datefmt': '%Y-%m-%dT%H:%M:%S.%f%z',
        'filename': '{}/general.log'
    },
    'icmp': {
        'level': 'DEBUG',
        'format': '%(asctime)s %(levelname)s: %(message)s',
        'datefmt': '%Y-%m-%dT%H:%M:%S.%f%z',
        'filename': '{}/icmp.log'
    },
    'tests': {
        'level': 'DEBUG',
        'format': '%(asctime)s %(levelname)s: %(message)s',
        'datefmt': '%Y-%m-%dT%H:%M:%S.%f%z',
        'filename': '{}/tests.log'
    },
    'event-log': {
        'level': 'DEBUG',
        'format': '%(asctime)s %(levelname)s: %(message)s',
        'datefmt': '%Y-%m-%dT%H:%M:%S.%f%z',
        'filename': '{}/event-log.log'
    }
}


def setup_diagnostics_logging(diagnostics_logs_dir):
    handlers = copy.deepcopy(_diagnostics_handlers)
    for handler in handlers.values():
        for key, value in handler.items():
            if key == 'filename':
                handler[key] = value.format(diagnostics_logs_dir)

    queue_listener = None
    queue_handler = None

    root_logger = logging.getLogger()

    def start():
        nonlocal queue_listener
        nonlocal queue_handler

        level_names_mapping = logging.getLevelNamesMapping()
        diagnostics_handlers = []

        for name, config in handlers.items():
            handler = logging.FileHandler(config['filename'], encoding='utf-8')
            handler.setLevel(level_names_mapping[config['level']])
            handler.set_name(name)

            formatter = DatetimeFormatter(config['format'], config['datefmt'])
            handler.setFormatter(formatter)

            # Restrict each handler to its respective logger, except for the root logger,
            # which is allowed to handle all records
            if name != 'root':
                filter_ = logging.Filter(name)
                handler.addFilter(filter_)

            current_logger = logging.getLogger(name)
            current_logger.propagate = False

            diagnostics_handlers.append(handler)

        # Create and configure the diagnostics queue listener and queue handler
        queue_ = queue.Queue()
        queue_listener = logging.handlers.QueueListener(queue_, *diagnostics_handlers, respect_handler_level=True)
        queue_handler = logging.handlers.QueueHandler(queue_)
        queue_handler.set_name('diagnostics_queue_handler')

        # Add the diagnostics queue handler to the root logger to allow handling records for all loggers
        root_logger.addHandler(queue_handler)

        # Add the diagnostics queue handler to each non-propagating diagnostics logger individually
        # to allow diagnostics-only log records to be written to files without being sent to other handlers
        for name in handlers.keys():
            current_logger = logging.getLogger(name)
            current_logger.addHandler(queue_handler)

        queue_listener.start()

    def stop():
        root_logger.removeHandler(queue_handler)
        for name in handlers.keys():
            current_logger = logging.getLogger(name)
            current_logger.removeHandler(queue_handler)

        queue_listener.stop()

        # Close the diagnostics handlers to flush the buffers and allow reading the files
        queue_handler.close()
        for handler in queue_listener.handlers:
            handler.close()

    return start, stop
