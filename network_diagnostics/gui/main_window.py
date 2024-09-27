import logging
import platform
import shutil
import tarfile

from PySide6.QtCore import QCoreApplication, Signal, qVersion
from PySide6.QtGui import QClipboard, Qt
from PySide6.QtWidgets import QFileDialog, QMainWindow, QMessageBox, QTableWidgetItem
from PySide6 import __version__ as __pyside_version__
from qdarktheme import load_palette, load_stylesheet

from ..diagnostics import Diagnostics, NetworkInterface, RTTData
from ..utils import get_diagnostics_dir, open_path_in_explorer, parse_cs2_ncs
from ..version import __version__
from .generated.main_window import Ui_MainWindow
from .history_window import HistoryWindow
from .logging_ import SignalHandler


log = logging.getLogger('gui')


class MainWindow(QMainWindow):

    _s_on_diagnostics_stop = Signal()
    _s_on_interruption_start = Signal(int)
    _s_on_interruption_end = Signal()
    _s_on_interface_update = Signal(NetworkInterface)
    _s_on_interface_stats_update = Signal(dict)
    _s_on_icmp_gateway_test_start = Signal(str)
    _s_on_icmp_gateway_test_update = Signal(RTTData)
    _s_on_icmp_external_test_start = Signal(str)
    _s_on_icmp_external_test_update = Signal(RTTData)
    _s_on_icmp_cs2_test_start = Signal(str)
    _s_on_icmp_cs2_test_update = Signal(RTTData)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Setup UI logging handler

        signal_handler = SignalHandler(self.ui.log_field, level=logging.INFO)
        signal_handler.message.connect(self.ui.log_field.appendPlainText)
        formatter = logging.Formatter('%(asctime)s %(levelname)s: %(message)s', '%H:%M:%S')
        signal_handler.setFormatter(formatter)

        logging.getLogger().addHandler(signal_handler)

        self._signal_handler = signal_handler

        # Connect signals from Diagnostics callbacks
        self._s_on_diagnostics_stop.connect(self._on_diagnostics_stop)
        self._s_on_interruption_start.connect(self._on_interruption_start)
        self._s_on_interruption_end.connect(self._on_interruption_end)
        self._s_on_interface_update.connect(self._on_interface_update)
        self._s_on_interface_stats_update.connect(self._on_interface_stats_update)
        self._s_on_icmp_gateway_test_start.connect(self._on_icmp_gateway_test_start)
        self._s_on_icmp_gateway_test_update.connect(self._on_icmp_gateway_test_update)
        self._s_on_icmp_external_test_start.connect(self._on_icmp_external_test_start)
        self._s_on_icmp_external_test_update.connect(self._on_icmp_external_test_update)
        self._s_on_icmp_cs2_test_start.connect(self._on_icmp_cs2_test_start)
        self._s_on_icmp_cs2_test_update.connect(self._on_icmp_cs2_test_update)

        # Connect UI widgets signals
        self.ui.action_browse_all_diagnostics.triggered.connect(self.browse_all_diagnostics)
        self.ui.action_delete_all_diagnostics.triggered.connect(self.delete_all_diagnostics)
        self.ui.action_exit.triggered.connect(self.close)

        self.ui.action_cloudflare_dns.triggered.connect(lambda: self.set_external_server('1.1.1.1'))
        self.ui.action_google_dns.triggered.connect(lambda: self.set_external_server('8.8.8.8'))
        self.ui.action_quad9_dns.triggered.connect(lambda: self.set_external_server('9.9.9.9'))

        self.ui.action_theme_dark.triggered.connect(lambda: self.change_theme('dark'))
        self.ui.action_theme_light.triggered.connect(lambda: self.change_theme('light'))

        self.ui.action_about.triggered.connect(self._on_menu_about)

        self.ui.diagnostics_toggle_button.toggled.connect(self._on_diagnostics_toggle_button)
        self.ui.export_last_button.clicked.connect(self._on_export_last_button)
        self.ui.history_button.clicked.connect(self._on_history_button)
        self.ui.cs2_ncs_field.textChanged.connect(self._on_cs2_ncs_field_change)
        self.ui.cs2_ncs_paste_button.clicked.connect(self._on_cs2_ncs_paste_button)
        self.ui.cs2_ncs_clear_button.clicked.connect(self._on_cs2_ncs_clear_button)

        # Initialize RTT table cells
        for row in range(3):
            for column in range(7):
                self.ui.rtt_table.setItem(row, column, QTableWidgetItem())

        self._exit_requested = False

        self._history_window = HistoryWindow(self, Qt.WindowType.Dialog)
        self._history_window.setWindowModality(Qt.WindowModality.WindowModal)

        self._external_server = '1.1.1.1'
        self._cs2_server = None

        self._diagnostics = Diagnostics(
            self._external_server,
            cb_on_stop=lambda: self._s_on_diagnostics_stop.emit(),
            cb_on_interruption_start=lambda total_interruptions: self._s_on_interruption_start.emit(total_interruptions),
            cb_on_interruption_end=lambda: self._s_on_interruption_end.emit(),
            cb_on_interface_update=lambda interface: self._s_on_interface_update.emit(interface),
            cb_on_interface_stats_update=lambda interface_stats: self._s_on_interface_stats_update.emit(interface_stats),
            cb_on_icmp_gateway_test_start=lambda host: self._s_on_icmp_gateway_test_start.emit(host),
            cb_on_icmp_gateway_test_update=lambda rtt_data: self._s_on_icmp_gateway_test_update.emit(rtt_data),
            cb_on_icmp_external_test_start=lambda host: self._s_on_icmp_external_test_start.emit(host),
            cb_on_icmp_external_test_update=lambda rtt_data: self._s_on_icmp_external_test_update.emit(rtt_data),
            cb_on_icmp_cs2_test_start=lambda host: self._s_on_icmp_cs2_test_start.emit(host),
            cb_on_icmp_cs2_test_update=lambda rtt_data: self._s_on_icmp_cs2_test_update.emit(rtt_data)
        )

    def change_theme(self, theme):
        app = QCoreApplication.instance()
        stylesheet = load_stylesheet(theme)
        app.setStyleSheet(stylesheet)
        palette = load_palette(theme)
        app.setPalette(palette)

    def set_external_server(self, server):
        if self._diagnostics.running:
            QMessageBox.critical(self, 'Error', 'Cannot change external server while diagnostics are running.')
            return

        self._external_server = server
        self._diagnostics.set_icmp_external_test_server(self._external_server)

    def export_diagnostics(self, name, path):
        if self._diagnostics.running:
            QMessageBox.critical(self, 'Error', 'Cannot export diagnostics while diagnostics are currently running.')
            return False

        tar_file_name = f'{name}.tar.xz'
        save_file, _ = QFileDialog.getSaveFileName(self, 'Save File', tar_file_name)
        if not save_file:
            return False

        log.debug(f"Exporting diagnostics '{name}' to '{save_file}'")

        try:
            with tarfile.open(save_file, 'w:xz', format=tarfile.PAX_FORMAT) as tar_file:
                tar_file.add(path, name)
        except Exception as e:
            log.debug(f"Error creating TAR file for diagnostics '{name}'", exc_info=True)
            log.error(f"Failed to export diagnostics '{name}': {e}")
            QMessageBox.critical(self, 'Error', f"Failed to export diagnostics '{name}': {e}")
            return False

        log.info(f"Successfully exported diagnostics '{name}'")
        return True

    def browse_diagnostics(self, path):
        try:
            open_path_in_explorer(str(path))
        except OSError as e:
            log.debug('Error opening diagnostics directory in explorer', exc_info=True)
            log.error(f'Failed to open diagnostics directory in explorer: {e}')
            QMessageBox.critical(self, 'Error', f'Failed to open diagnostics directory in explorer: {e}')

    def browse_all_diagnostics(self):
        diagnostics_dir = get_diagnostics_dir()
        try:
            open_path_in_explorer(str(diagnostics_dir))
        except OSError as e:
            log.debug('Error opening all diagnostics directory in explorer', exc_info=True)
            log.error(f'Failed to open all diagnostics directory in explorer: {e}')
            QMessageBox.critical(self, 'Error', f'Failed to open all diagnostics directory in explorer: {e}')

    def delete_diagnostics(self, name, path):
        if self._diagnostics.running:
            QMessageBox.critical(self, 'Error', 'Cannot delete diagnostics while diagnostics are currently running.')
            return False

        result = QMessageBox.question(self, 'Confirm', 'Are you sure you want to permanently delete this directory?')
        if result != QMessageBox.StandardButton.Yes:
            return False

        try:
            log.debug(f"Deleting diagnostics directory '{path}'")
            shutil.rmtree(str(path))
        except OSError as e:
            log.debug(f"Error deleting diagnostics directory '{path}'", exc_info=True)
            log.error(f"Failed to delete diagnostics directory '{path}': {e}")
            QMessageBox.critical(self, 'Error', f"Failed to delete diagnostics directory '{path}': {e}")
            return False

        log.info(f"Successfully deleted diagnostics '{name}'")
        return True

    def delete_all_diagnostics(self):
        if self._diagnostics.running:
            QMessageBox.critical(self, 'Error', 'Cannot delete diagnostics while diagnostics are currently running.')
            return False

        result = QMessageBox.question(self, 'Confirm', 'Are you sure you want to permanently delete all diagnostics?')
        if result != QMessageBox.StandardButton.Yes:
            return False

        log.debug('Deleting all diagnostics')
        for name, path in self._diagnostics.get_diagnostics_history():
            try:
                log.debug(f"Deleting diagnostics directory '{path}'")
                shutil.rmtree(str(path))
            except OSError as e:
                log.debug(f"Error deleting diagnostics directory '{path}'", exc_info=True)
                log.error(f"Failed to delete diagnostics directory '{path}': {e}")
                QMessageBox.critical(self, 'Error', f"Failed to delete diagnostics directory '{path}': {e}")
                return False

        log.info('Successfully deleted all diagnostics')
        return True

    # UI handlers

    def _on_menu_about(self):
        QMessageBox.about(
            self,
            'About',
            (
                f'Network Diagnostics {__version__}\n\n'
                'Licensed under the GNU General Public License v3.0.\n\n'
                'Version information:\n'
                f'Windows {platform.version()} {platform.architecture()[0]}\n'
                f'Python {platform.python_version()}\n'
                f'PySide {__pyside_version__}\n'
                f'Qt {qVersion()}'
            )
        )

    def _on_diagnostics_toggle_button(self, checked):
        if checked:
            try:
                self._diagnostics.start()
            except Exception as e:
                log.debug('Error starting diagnostics', exc_info=True)
                log.error(f'Failed to start diagnostics: {e}')
                QMessageBox.critical(self, 'Error', f'Failed to start diagnostics: {e}')
                self.ui.diagnostics_toggle_button.setChecked(False)
                return

            self.ui.diagnostics_toggle_button.setText('Stop')
            self.ui.export_last_button.setEnabled(False)
            self.ui.history_button.setEnabled(False)
            self.ui.interruptions_value_label.setText('0')
            self.ui.statusbar.showMessage('Diagnostics running')

        elif self._diagnostics.running:
            self._diagnostics.stop(False)

            self.setEnabled(False)
            self.ui.statusbar.showMessage('Stopping diagnostics...')

    def _on_export_last_button(self):
        try:
            diagnostics_history = self._diagnostics.get_diagnostics_history()
        except Exception as e:
            log.debug('Failed to get diagnostics history', exc_info=True)
            log.error(f'Failed to get diagnostics history: {e}')
            QMessageBox.critical(self, 'Error', f'Failed to get diagnostics history: {e}')
            return

        if not diagnostics_history:
            QMessageBox.information(self, 'Info', 'No diagnostics found')
            return

        name, path = diagnostics_history[0]
        self.export_diagnostics(name, path)

    def _on_history_button(self):
        try:
            diagnostics_history = self._diagnostics.get_diagnostics_history()
        except Exception as e:
            log.debug('Failed to get diagnostics history', exc_info=True)
            log.error(f'Failed to get diagnostics history: {e}')
            QMessageBox.critical(self, 'Error', f'Failed to get diagnostics history: {e}')
            return

        if not diagnostics_history:
            QMessageBox.information(self, 'Info', 'No diagnostics found')
            return

        try:
            self._history_window.update_history(diagnostics_history)
        except Exception as e:
            log.debug('Error updating history with diagnostics data', exc_info=True)
            log.error(f'Error updating history with diagnostics data: {e}')
            QMessageBox.critical(self, 'Error', f'Error updating history with diagnostics data: {e}')
            return

        self._history_window.show()

    def _on_cs2_ncs_field_change(self):
        cs2_console_output = self.ui.cs2_ncs_field.toPlainText()

        if not cs2_console_output.strip():
            self._cs2_server = None

            # Reset console parse status
            self.ui.cs2_ncs_parse_value_label.setText('Waiting for input...')
            self.ui.cs2_ncs_parse_value_label.setStyleSheet('')

            # Reset connection status values
            self.ui.cs2_ncs_game_server_value_label.setText('-')
            self.ui.cs2_ncs_primary_relay_value_label.setText('-')
            self.ui.cs2_ncs_backup_relay_value_label.setText('-')

            self._diagnostics.set_icmp_cs2_test_server(self._cs2_server)

            return

        try:
            cs2_ncs = parse_cs2_ncs(cs2_console_output)
        except ValueError as e:
            self._cs2_server = None

            # Set console parse status to error
            self.ui.cs2_ncs_parse_value_label.setText('Incorrect or incomplete CS2 console output. Please try again.')
            self.ui.cs2_ncs_parse_value_label.setStyleSheet('QLabel { color: #ff4040; }')

            # Reset connection status values
            self.ui.cs2_ncs_game_server_value_label.setText('-')
            self.ui.cs2_ncs_primary_relay_value_label.setText('-')
            self.ui.cs2_ncs_backup_relay_value_label.setText('-')

            self._diagnostics.set_icmp_cs2_test_server(self._cs2_server)

            return

        self._cs2_server = cs2_ncs['primary_relay_address']

        # Set console parse status to success
        self.ui.cs2_ncs_parse_value_label.setText('CS2 server connection information available.')
        self.ui.cs2_ncs_parse_value_label.setStyleSheet('QLabel { color: #10b010; }')

        # Populate connection status values

        self.ui.cs2_ncs_game_server_value_label.setText(cs2_ncs['server_location'])

        pr_location = cs2_ncs['primary_relay_location']
        pr_address = cs2_ncs['primary_relay_address']
        pr_port = cs2_ncs['primary_relay_port']
        pr_latency_front = cs2_ncs['primary_relay_latency_front']
        pr_latency_back = cs2_ncs['primary_relay_latency_back']
        pr_latency_total = pr_latency_front + pr_latency_back
        self.ui.cs2_ncs_primary_relay_value_label.setText(
            f'{pr_location}, {pr_address}:{pr_port}, {pr_latency_total} ({pr_latency_front}+{pr_latency_back}) ms'
        )

        br_location = cs2_ncs['backup_relay_location']
        br_address = cs2_ncs['backup_relay_address']
        br_port = cs2_ncs['backup_relay_port']
        br_latency_front = cs2_ncs['backup_relay_latency_front']
        br_latency_back = cs2_ncs['backup_relay_latency_back']
        br_latency_total = br_latency_front + br_latency_back
        self.ui.cs2_ncs_backup_relay_value_label.setText(
            f'{br_location}, {br_address}:{br_port}, {br_latency_total} ({br_latency_front}+{br_latency_back}) ms'
        )

        self._diagnostics.set_icmp_cs2_test_server(self._cs2_server)

    def _on_cs2_ncs_paste_button(self):
        clipboard = QClipboard(self)
        clipboard_text = clipboard.text()
        self.ui.cs2_ncs_field.setPlainText(clipboard_text)

    def _on_cs2_ncs_clear_button(self):
        self.ui.cs2_ncs_field.setPlainText('')

    # Event handlers

    def closeEvent(self, event):  # noqa: N802
        if not self._diagnostics.running:
            event.accept()
            return

        event.ignore()

        result = QMessageBox.question(
            self,
            'Confirm',
            'Diagnostics are currently running. Would you like to stop the diagnostics and exit?'
        )
        if result == QMessageBox.StandardButton.Yes:
            self._exit_requested = True

            self._diagnostics.stop(False)

            self.setEnabled(False)

            self.ui.statusbar.showMessage('Stopping diagnostics...')

    # Diagnostics handlers

    def _on_diagnostics_stop(self):
        if self._exit_requested:
            self.close()
            return

        self.setEnabled(True)

        self.ui.diagnostics_toggle_button.setText('Start')
        self.ui.export_last_button.setEnabled(True)
        self.ui.history_button.setEnabled(True)

        self.ui.active_interface_value_label.setText('-')
        self.ui.link_status_value_label.setText('-')
        self.ui.link_speed_value_label.setText('-')
        self.ui.internet_connectivity_value_label.setText('-')
        self.ui.interruptions_value_label.setText('-')

        self.ui.statusbar.clearMessage()

    def _on_interruption_start(self, total_interruptions):
        self.ui.internet_connectivity_label.setText('Unstable')
        self.ui.interruptions_value_label.setText(str(total_interruptions))

    def _on_interruption_end(self):
        self.ui.internet_connectivity_label.setText('Stable')

    def _on_interface_update(self, interface):
        self.ui.active_interface_value_label.setText(interface.name)

    def _on_interface_stats_update(self, interface_stats):
        status = 'Up' if interface_stats['up'] else 'Down'
        self.ui.link_status_value_label.setText(status)

        speed = interface_stats['speed']
        duplex = interface_stats['duplex_str']
        self.ui.link_speed_value_label.setText(f'{speed} Mbps ({duplex} duplex)')

        self.ui.internet_connectivity_value_label.setText('Stable')

    def _on_icmp_gateway_test_start(self, host):
        self.ui.rtt_table.item(0, 0).setText('-')
        self.ui.rtt_table.item(0, 1).setText('-')
        self.ui.rtt_table.item(0, 2).setText('-')
        self.ui.rtt_table.item(0, 3).setText('-')
        self.ui.rtt_table.item(0, 4).setText('-')
        self.ui.rtt_table.item(0, 5).setText('-')
        self.ui.rtt_table.item(0, 6).setText('-')

    def _on_icmp_gateway_test_update(self, rtt_data):
        self.ui.rtt_table.item(0, 0).setText(f'{rtt_data.average:.3f}' if rtt_data.average is not None else '-')
        self.ui.rtt_table.item(0, 1).setText(f'{rtt_data.minimum:.3f}' if rtt_data.minimum is not None else '-')
        self.ui.rtt_table.item(0, 2).setText(f'{rtt_data.maximum:.3f}' if rtt_data.maximum is not None else '-')
        self.ui.rtt_table.item(0, 3).setText(f'{rtt_data.jitter:.3f}' if rtt_data.jitter is not None else '-')
        self.ui.rtt_table.item(0, 4).setText(f'{rtt_data.sent}')
        self.ui.rtt_table.item(0, 5).setText(f'{rtt_data.received}')
        self.ui.rtt_table.item(0, 6).setText(f'{rtt_data.loss * 100:.1f}%')

    def _on_icmp_external_test_start(self, host):
        self.ui.rtt_table.item(1, 0).setText('-')
        self.ui.rtt_table.item(1, 1).setText('-')
        self.ui.rtt_table.item(1, 2).setText('-')
        self.ui.rtt_table.item(1, 3).setText('-')
        self.ui.rtt_table.item(1, 4).setText('-')
        self.ui.rtt_table.item(1, 5).setText('-')
        self.ui.rtt_table.item(1, 6).setText('-')

    def _on_icmp_external_test_update(self, rtt_data):
        self.ui.rtt_table.item(1, 0).setText(f'{rtt_data.average:.3f}' if rtt_data.average is not None else '-')
        self.ui.rtt_table.item(1, 1).setText(f'{rtt_data.minimum:.3f}' if rtt_data.minimum is not None else '-')
        self.ui.rtt_table.item(1, 2).setText(f'{rtt_data.maximum:.3f}' if rtt_data.maximum is not None else '-')
        self.ui.rtt_table.item(1, 3).setText(f'{rtt_data.jitter:.3f}' if rtt_data.jitter is not None else '-')
        self.ui.rtt_table.item(1, 4).setText(f'{rtt_data.sent}')
        self.ui.rtt_table.item(1, 5).setText(f'{rtt_data.received}')
        self.ui.rtt_table.item(1, 6).setText(f'{rtt_data.loss * 100:.1f}%')

    def _on_icmp_cs2_test_start(self, host):
        self.ui.rtt_table.item(2, 0).setText('-')
        self.ui.rtt_table.item(2, 1).setText('-')
        self.ui.rtt_table.item(2, 2).setText('-')
        self.ui.rtt_table.item(2, 3).setText('-')
        self.ui.rtt_table.item(2, 4).setText('-')
        self.ui.rtt_table.item(2, 5).setText('-')
        self.ui.rtt_table.item(2, 6).setText('-')

    def _on_icmp_cs2_test_update(self, rtt_data):
        self.ui.rtt_table.item(2, 0).setText(f'{rtt_data.average:.3f}' if rtt_data.average is not None else '-')
        self.ui.rtt_table.item(2, 1).setText(f'{rtt_data.minimum:.3f}' if rtt_data.minimum is not None else '-')
        self.ui.rtt_table.item(2, 2).setText(f'{rtt_data.maximum:.3f}' if rtt_data.maximum is not None else '-')
        self.ui.rtt_table.item(2, 3).setText(f'{rtt_data.jitter:.3f}' if rtt_data.jitter is not None else '-')
        self.ui.rtt_table.item(2, 4).setText(f'{rtt_data.sent}')
        self.ui.rtt_table.item(2, 5).setText(f'{rtt_data.received}')
        self.ui.rtt_table.item(2, 6).setText(f'{rtt_data.loss * 100:.1f}%')
