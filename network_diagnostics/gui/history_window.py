from datetime import datetime

from PySide6.QtWidgets import QListWidgetItem, QMainWindow

from .generated.history_window import Ui_MainWindow


class HistoryWindow(QMainWindow):

    def __init__(self, main_window, *args, **kwargs):
        super().__init__(main_window, *args, **kwargs)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.main_window = main_window

        self.ui.export_button.clicked.connect(self._on_export_button)
        self.ui.browse_button.clicked.connect(self._on_browse_button)
        self.ui.delete_button.clicked.connect(self._on_delete_button)
        self.ui.delete_all_button.clicked.connect(self._on_delete_all_button)
        self.ui.diagnostics_history_list.itemSelectionChanged.connect(self._on_history_list_selection_change)

        self.history = None
        self.selected_index = None

    def update_history(self, history):
        self.history = history

        self.ui.export_button.setEnabled(False)
        self.ui.browse_button.setEnabled(False)
        self.ui.delete_button.setEnabled(False)
        self.ui.delete_all_button.setEnabled(False)
        self.ui.diagnostics_history_list.clear()

        if history:
            self.ui.delete_all_button.setEnabled(True)

        for name, path in history:
            name_split = name.split('.')
            if len(name_split) >= 2:
                # Name should be file-safe ISO 8601 string + number
                try:
                    dt_utc = datetime.strptime(name_split[0], '%Y-%m-%dT%H-%M-%S%z')
                except ValueError:
                    item_name = name_split[1]
                else:
                    dt_str = dt_utc.astimezone().strftime('%Y-%m-%d %H:%M:%S')
                    item_name = f'{dt_str} #{name_split[1]}'

            else:
                # Name should be file-safe ISO 8601 string
                try:
                    dt_utc = datetime.strptime(name, '%Y-%m-%dT%H-%M-%S%z')
                except ValueError:
                    item_name = name
                else:
                    item_name = dt_utc.astimezone().strftime('%Y-%m-%d %H:%M:%S')

            item = QListWidgetItem(item_name)
            self.ui.diagnostics_history_list.addItem(item)

    def get_selected_index(self):
        selected_indexes = self.ui.diagnostics_history_list.selectedIndexes()
        if selected_indexes:
            return selected_indexes[0].row()
        return None

    # UI handlers

    def _on_export_button(self):
        name, path = self.history[self.selected_index]
        self.main_window.export_diagnostics(name, path)

    def _on_browse_button(self):
        name, path = self.history[self.selected_index]
        self.main_window.browse_diagnostics(path)

    def _on_delete_button(self):
        name, path = self.history[self.selected_index]
        success = self.main_window.delete_diagnostics(name, path)
        if not success:
            return

        del self.history[self.selected_index]
        self.ui.diagnostics_history_list.takeItem(self.selected_index)

    def _on_delete_all_button(self):
        success = self.main_window.delete_all_diagnostics()
        if not success:
            return

        self.history.clear()
        self.ui.diagnostics_history_list.clear()

        self.ui.export_button.setEnabled(False)
        self.ui.browse_button.setEnabled(False)
        self.ui.delete_button.setEnabled(False)
        self.ui.delete_all_button.setEnabled(False)

    def _on_history_list_selection_change(self):
        self.selected_index = self.get_selected_index()
        if self.selected_index is not None:
            self.ui.export_button.setEnabled(True)
            self.ui.browse_button.setEnabled(True)
            self.ui.delete_button.setEnabled(True)
            return
        self.ui.export_button.setEnabled(False)
        self.ui.browse_button.setEnabled(False)
        self.ui.delete_button.setEnabled(False)
