from logging import Handler
from PySide6.QtCore import QObject, Signal


class _SignalProxy(QObject):

    message = Signal(str)


class SignalHandler(Handler):

    def __init__(self, widget, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._widget = widget
        self._signal_proxy = _SignalProxy()
        self.message = self._signal_proxy.message

    def emit(self, record):
        self.message.emit(self.format(record))
