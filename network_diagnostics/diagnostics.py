import datetime
import logging
import threading
import time
import xml.dom.minidom

import cfspeedtest
import icmplib
import icmplib.utils
import netifaces
import psutil
import win32evtlog
import subprocess

from .logging_ import setup_diagnostics_logging
from .utils import get_diagnostics_dir, create_firewall_rules, remove_firewall_rules


log = logging.getLogger('diagnostics')
log_icmp = logging.getLogger('icmp')
log_tests = logging.getLogger('tests')
log_event_log = logging.getLogger('event-log')


class Diagnostics:

    def __init__(
            self,
            external_test_server,
            cb_on_stop=None,
            cb_on_interruption_start=None,
            cb_on_interruption_end=None,
            cb_on_interface_update=None,
            cb_on_interface_stats_update=None,
            cb_on_icmp_gateway_test_start=None,
            cb_on_icmp_gateway_test_update=None,
            cb_on_icmp_external_test_start=None,
            cb_on_icmp_external_test_update=None,
            cb_on_icmp_cs2_test_start=None,
            cb_on_icmp_cs2_test_update=None
        ):
        self._cb_on_stop = cb_on_stop
        self._cb_on_interruption_start = cb_on_interruption_start
        self._cb_on_interruption_end = cb_on_interruption_end
        self._cb_on_interface_update = cb_on_interface_update
        self._cb_on_interface_stats_update = cb_on_interface_stats_update
        self._cb_on_icmp_gateway_test_start = cb_on_icmp_gateway_test_start
        self._cb_on_icmp_gateway_test_update = cb_on_icmp_gateway_test_update
        self._cb_on_icmp_external_test_start = cb_on_icmp_external_test_start
        self._cb_on_icmp_external_test_update = cb_on_icmp_external_test_update
        self._cb_on_icmp_cs2_test_start = cb_on_icmp_cs2_test_start
        self._cb_on_icmp_cs2_test_update = cb_on_icmp_cs2_test_update

        self._diagnostics_dir = get_diagnostics_dir()
        self._stop_diagnostics_logging = None

        self._running = False

        self._active_interruption = False
        self._total_interruptions = 0
        self._internet_connectivity = True

        self._watchdog_thread = None
        self._watchdog_interval = 1.0
        self._watchdog_interruption_tests_interval = 1.0

        self._interface = None
        self._interface_stats = None
        self._interface_changed = False
        self._interface_update_thread = None
        self._interface_update_running = False
        self._interface_update_interval = 1.0

        self._icmp_gateway_test = None
        self._icmp_gateway_test_thread = None
        self._icmp_gateway_test_seq_loss = 0

        self._icmp_external_test = None
        self._icmp_external_test_server = external_test_server
        self._icmp_external_test_server_changed = False
        self._icmp_external_test_thread = None
        self._icmp_external_test_seq_loss = 0

        self._icmp_cs2_test = None
        self._icmp_cs2_test_server = None
        self._icmp_cs2_test_server_changed = False
        self._icmp_cs2_test_thread = None

        self._icmp_late_time = 0.5
        self._icmp_max_seq_loss = 1
        self._icmp_interval = 1.0
        self._icmp_timeout = 2.0

        self._traceroute = None
        self._traceroute_thread = None

        self._speed_test = None
        self._speed_test_thread = None
        self._speed_test_min_duration = 10.0

        self._event_log_listener = None
        self._event_log_listener_thread = None

    @property
    def running(self):
        if self._watchdog_thread and self._watchdog_thread.is_alive():
            return True
        return False

    def start(self):
        log.debug('Diagnostics start')

        if self._watchdog_thread and self._watchdog_thread.is_alive():
            log.debug('Attempted to start diagnostics while thread is alive')
            return
        self._running = True

        timestamp = datetime.datetime.now().astimezone().strftime('%Y-%m-%dT%H-%M-%S%z')
        current_diagnostics_dir = self._diagnostics_dir / timestamp
        i = 2
        try:
            while current_diagnostics_dir.exists():
                current_diagnostics_dir = self._diagnostics_dir / f'{timestamp}.{i}'
                i += 1
        except:
            self._running = False
            raise

        try:
            create_firewall_rules()
        except:
            self._running = False
            raise

        log.debug(f"Creating current diagnostics directory '{current_diagnostics_dir}'")
        try:
            current_diagnostics_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            raise OSError(
                e.errno,
                'Failed to create current diagnostics directory',
                str(current_diagnostics_dir)
            ) from e
        except:
            self._running = False
            raise

        try:
            log.info('Starting diagnostics logging')
            start, stop = setup_diagnostics_logging(current_diagnostics_dir)
            self._stop_diagnostics_logging = stop
            start()
        except:
            self._running = False
            raise

        def wrapper():
            try:
                self._watchdog_run()
            finally:
                self._clean_up()

        self._watchdog_thread = threading.Thread(target=wrapper, name='watchdog', daemon=True)
        self._watchdog_thread.start()

        log.info('Diagnostics started')

    def stop(self, blocking=True):
        log.debug(f'Diagnostics stop, blocking: {blocking}')

        self._running = False
        if blocking and self._watchdog_thread:
            self._watchdog_thread.join()

    def set_icmp_external_test_server(self, server):
        if server is None:
            raise ValueError('External server cannot be None')
        self._icmp_external_test_server = server
        self._icmp_external_test_server_changed = True

    def set_icmp_cs2_test_server(self, server):
        if self._icmp_cs2_test_server == server:
            return
        self._icmp_cs2_test_server = server
        self._icmp_cs2_test_server_changed = True

    def get_diagnostics_history(self):
        return sorted(
            ((path.name, path) for path in self._diagnostics_dir.iterdir()),
            key=lambda item: item[0],
            reverse=True
        )

    def get_last_diagnostics(self):
        history = self.get_diagnostics_history()
        return history[0] if history else None

    def _watchdog_run(self):
        log.debug('Watchdog start')

        self._start_event_log_listener()
        self._start_interface_update()

        # Wait for the interface update to finish or exit early if requested
        while self._interface_update_thread and self._interface_update_thread.is_alive():
            if not self._running:
                log.debug('Stopping watchdog during initial interface update')
                self._stop_interface_update()
                return
            self._interface_update_thread.join(self._watchdog_interval)

        # Log current statistics
        self._log_netsat_statistics()

        # Run initial traceroute
        log.info('Starting initial traceroute')
        self._run_traceroute(self._icmp_external_test_server)

        # Active interface has been found and initial traceroute has started
        # Begin ICMP tests and connection monitoring

        self._start_icmp_gateway_test(self._interface.gateway_ipv4_address)

        if not self._icmp_external_test_server:
            log.error('No ICMP external test server specified')
            log.debug('Watchdog stop')
            return

        self._start_icmp_external_test(self._icmp_external_test_server)

        cs2_server = self._icmp_cs2_test_server
        if cs2_server:
            self._start_icmp_cs2_test(cs2_server)
        else:
            log.info('Not starting ICMP CS2 test because server is not specified')

        while self._running:
            # Check for connection to external server and begin advanced diagnostics on interruption

            if not self._internet_connectivity:
                if not self._active_interruption:
                    log.info('Internet connection lost')

                    self._active_interruption = True
                    self._total_interruptions += 1
                    self._on_interruption_start(self._total_interruptions)

            elif self._active_interruption:
                log.info('Internet connection established')

                self._active_interruption = False
                self._on_interruption_end()

            # Check for changes and restart components accordingly

            if self._interface_changed:
                gateway_address = self._interface.gateway_ipv4_address

                log.info('Active interface was updated')
                self._interface_changed = False

                self._stop_icmp_gateway_test()
                self._start_icmp_gateway_test(gateway_address)

                self._wait_traceroute()
                self._run_traceroute(self._icmp_external_test_server)

            if self._icmp_external_test_server_changed:
                log.info(f'ICMP test server was set to: {self._icmp_external_test_server}')
                self._icmp_external_test_server_changed = False

                self._stop_icmp_external_test()
                self._start_icmp_external_test(self._icmp_external_test_server)

                self._wait_traceroute()
                self._run_traceroute(self._icmp_external_test_server)

            if self._icmp_cs2_test_server_changed:
                cs2_server = self._icmp_cs2_test_server

                log.info(f'ICMP CS2 test server was set to: {cs2_server}')
                self._icmp_cs2_test_server_changed = False

                self._stop_icmp_cs2_test()

                if cs2_server is not None:
                    self._start_icmp_cs2_test(cs2_server)

                    self._wait_traceroute()
                    self._run_traceroute(cs2_server)

            time.sleep(self._watchdog_interval)

        log.debug('Watchdog stop')

    def _clean_up(self):
        log.debug('Clean up')

        self._running = False

        self._wait_traceroute()
        self._wait_speed_test()
        self._stop_icmp_gateway_test()
        self._stop_icmp_external_test()
        self._stop_icmp_cs2_test()
        self._stop_interface_update()
        self._stop_event_log_listener()

        self._active_interruption = False
        self._total_interruptions = 0

        self._watchdog_thread = None

        self._interface = None
        self._interface_stats = None
        self._interface_changed = False
        self._interface_update_thread = None

        self._icmp_gateway_test = None
        self._icmp_gateway_test_thread = None

        self._icmp_external_test = None
        self._icmp_external_test_thread = None

        self._icmp_cs2_test = None
        self._icmp_cs2_test_thread = None

        self._traceroute = None
        self._traceroute_thread = None

        self._speed_test = None
        self._speed_test_thread = None

        try:
            remove_firewall_rules()
        except OSError as e:
            log.debug('Error removing firewall rules', exc_info=True)
            log.error(f'Failed to remove firewall rules: {e}')

        log.info('Stopping diagnostics logging')
        if self._stop_diagnostics_logging:
            self._stop_diagnostics_logging()
        self._stop_diagnostics_logging = None

        if self._cb_on_stop is not None:
            self._cb_on_stop()

        log.info('Diagnostics stopped')

    def _log_netsat_statistics(self):
        log.debug('Log netstat statistics')

        try:
            log_tests.info(Netstat.get_statistics())
        except Exception as e:
            log.debug('Error getting netstat statistics', exc_info=True)
            log.error(f'Error getting netstat statistics: {e}')
            log_tests.error('Error getting netstat statistics', exc_info=True)

    def _on_interruption_start(self, total_interruptions):
        log.debug('On interruption start')

        if self._cb_on_interruption_start is not None:
            self._cb_on_interruption_start(total_interruptions)

        # Before starting more diagnostics, wait for all currently running diagnostics
        self._wait_traceroute()
        self._wait_speed_test()

        # Start searching for the active interface in case it changed
        self._start_interface_update()

        # Log current statistics
        self._log_netsat_statistics()

        # Keep retrying to run tests until either:
        # - The diagnostics are no longer running
        # - Connection has become stable again
        # - All tests have been completed successfully

        long_speed_test = False
        traceroute_succeeded = False
        short_speed_test_succeeded = False
        long_speed_test_succeeded = False

        while self._running and self._active_interruption:
            # Re-run traceroute if it was unsuccessful last time
            if not traceroute_succeeded:
                log.info('Starting traceroute')
                log_tests.info('Starting traceroute')
                self._run_traceroute(self._icmp_external_test_server, count=3, timeout=1.0, fast=True)

            # Re-run short speed test if it was unsuccessful last time
            if not short_speed_test_succeeded:
                log.info('Starting short speed test')
                log_tests.info('Starting short speed test')

                start = time.time()
                self._run_speed_test(SpeedTest.SHORT_TESTS)
                self._wait_speed_test()
                end = time.time()

                # Check if the speed test was successful and if long test should be ran
                if self._speed_test.results is not None:
                    short_speed_test_succeeded = True
                    if end - start < self._speed_test_min_duration:
                        long_speed_test = True

            # Re-run long speed test if it is necessary and was unsuccessful last time
            if not long_speed_test_succeeded and long_speed_test:
                log.info('Starting long speed test')
                log_tests.info('Starting long speed test')

                self._run_speed_test(SpeedTest.SHORT_TESTS)
                self._wait_speed_test()

                if self._speed_test.results is not None:
                    long_speed_test_succeeded = True

            # Check if the traceroute started before the speed test is successful
            if not traceroute_succeeded:
                self._wait_traceroute()
                if self._traceroute.hops is not None:
                    traceroute_succeeded = True

            if (
                traceroute_succeeded
                and short_speed_test_succeeded
                and long_speed_test_succeeded
            ):
                # All tests have completed successfully
                log.info('All tests on interruption finished')
                break
            else:
                # Wait before retrying tests
                time.sleep(self._watchdog_interruption_tests_interval)

    def _on_interruption_end(self):
        log.debug('On interruption end')

        if self._cb_on_interruption_end is not None:
            self._cb_on_interruption_end()

    def _start_interface_update(self):
        log.debug('Start interface update')

        if self._interface_update_thread and self._interface_update_thread.is_alive():
            log.debug('Attempted to start interface update while thread is alive')
            return
        self._interface_update_running = True

        self._interface_update_thread = threading.Thread(
            target=self._interface_update_run,
            name='interface-update',
            daemon=True
        )
        self._interface_update_thread.start()

    def _stop_interface_update(self):
        log.debug('Stop interface update')

        self._interface_update_running = False
        if self._interface_update_thread:
            self._interface_update_thread.join()

    def _interface_update_run(self):
        log.info('Interface update started')

        while self._interface_update_running:
            try:
                interface = NetworkInterface.from_default_gateway()
            except Exception:
                log.debug('Error getting active interface', exc_info=True)

                if not self._interface:
                    log.error('Failed to get active interface from default gateway, retrying')
                    time.sleep(self._interface_update_interval)
                    continue

                log.error('Failed to get active interface from default gateway')

            try:
                stats = interface.get_stats()
            except Exception:
                log.debug('Error getting active interface statistics', exc_info=True)
                log.error('Failed to get active interface statistics, retrying')
                time.sleep(self._interface_update_interval)
                continue

            if not self._interface:
                # First iteration where the active interface is found

                log.debug('Initial interface update')

                self._interface = interface
                self._interface_stats = stats
                self._on_interface_update(interface)
                self._on_interface_stats_update(stats)

            elif (
                self._interface.id != interface.id
                or self._interface.name != interface.name
                or self._interface.mac_address != interface.mac_address
                or self._interface.ipv4_address != interface.ipv4_address
                or self._interface.gateway_ipv4_address != interface.gateway_ipv4_address
            ):
                # Subsequent iterations where the active interface has changed

                log.debug(f"Active interface changed, ID: {interface.id}, name: {interface.name}, up: {stats['up']}")

                self._interface = interface
                self._interface_stats = stats
                self._interface_changed = True
                self._on_interface_update(interface)
                self._on_interface_stats_update(stats)

            elif self._interface_stats != stats:
                # Subsequent iterations where the active interface stats have changed

                log.debug(f"Active interface stats changed, up: {stats['up']}")

                self._interface_stats = stats
                self._on_interface_stats_update(stats)

            if self._interface_stats['up']:
                # The interface found is up and searching is no longer necessary

                log.debug('Interface is up, stopping update')

                self._interface_update_running = False
                break

            # Wait for the specified interval before attempting another search
            time.sleep(self._interface_update_interval)

        log.info('Interface update stopped')

    def _on_interface_update(self, interface):
        name = interface.name
        speed = self._interface_stats['speed']
        duplex = self._interface_stats['duplex_str']
        log.info(f'Active interface found: {name}, {speed} Mbps ({duplex} duplex)')

        if self._cb_on_interface_update is not None:
            self._cb_on_interface_update(interface)

    def _on_interface_stats_update(self, interface_stats):
        if self._cb_on_interface_stats_update is not None:
            self._cb_on_interface_stats_update(interface_stats)

    def _start_icmp_gateway_test(self, gateway_address):
        log.debug('Start ICMP gateway test')

        if self._icmp_gateway_test and self._icmp_gateway_test.running:
            log.debug('Attempted to start ICMP gateway test while previous test is running')
            return

        if self._cb_on_icmp_gateway_test_start is not None:
            self._cb_on_icmp_gateway_test_start(gateway_address)

        self._icmp_gateway_test = ICMPTest(
            gateway_address,
            self._icmp_interval,
            self._icmp_timeout,
            self._on_icmp_gateway_test_update,
            self._on_icmp_gateway_test_error
        )
        self._icmp_gateway_test_thread = threading.Thread(
            target=self._icmp_gateway_test.run,
            name='icmp-gateway-test',
            daemon=True
        )
        self._icmp_gateway_test_thread.start()

        log.info(f'ICMP gateway test started to: {gateway_address}')
        log_icmp.info(f'ICMP gateway test started to: {gateway_address}')

    def _stop_icmp_gateway_test(self):
        log.debug('Stop ICMP gateway test')

        if self._icmp_gateway_test:
            self._icmp_gateway_test.stop()
            self._icmp_gateway_test_thread.join()

            log_icmp.info(f'Gateway RTT data:\n{self._icmp_gateway_test.rtt_data}')

    def _on_icmp_gateway_test_update(self, rtt_data, lost):
        if lost:
            log_icmp.info(f'Gateway ({rtt_data.host}): Timeout')
        elif rtt_data.last >= self._icmp_late_time:
            log_icmp.info(f'Gateway ({rtt_data.host}): {rtt_data.last} (late)')
        else:
            log_icmp.info(f'Gateway ({rtt_data.host}): {rtt_data.last}')

        if self._cb_on_icmp_gateway_test_update is not None:
            self._cb_on_icmp_gateway_test_update(rtt_data)

    def _on_icmp_gateway_test_error(self, exception):
        log.debug('ICMP gateway test error', exc_info=True)
        log.error(f'ICMP gateway test error: {exception}')
        log_icmp.error('ICMP gateway test error', exc_info=True)
        log_icmp.info(f'Gateway RTT data (on error):\n{self._icmp_gateway_test.rtt_data}')

        time.sleep(self._icmp_interval)

        gateway_address = self._interface.gateway_ipv4_address
        if gateway_address is not None:
            self._start_icmp_gateway_test(gateway_address)

    def _start_icmp_external_test(self, external_server_address):
        log.debug('Start ICMP external test')

        if self._icmp_external_test and self._icmp_external_test.running:
            log.debug('Attempted to start ICMP external test while previous test is running')
            return

        if self._cb_on_icmp_external_test_start is not None:
            self._cb_on_icmp_external_test_start(external_server_address)

        self._icmp_external_test = ICMPTest(
            external_server_address,
            self._icmp_interval,
            self._icmp_timeout,
            self._on_icmp_external_test_update,
            self._on_icmp_external_test_error
        )
        self._icmp_external_test_thread = threading.Thread(
            target=self._icmp_external_test.run,
            name='icmp-test',
            daemon=True
        )
        self._icmp_external_test_thread.start()

        log.info(f'ICMP external test started to: {external_server_address}')
        log_icmp.info(f'ICMP external test started to: {external_server_address}')

    def _stop_icmp_external_test(self):
        log.debug('Stop ICMP external test')

        if self._icmp_external_test:
            self._icmp_external_test.stop()
            self._icmp_external_test_thread.join()

            log_icmp.info(f'External RTT data:\n{self._icmp_external_test.rtt_data}')

    def _on_icmp_external_test_update(self, rtt_data, lost):
        if self._cb_on_icmp_external_test_update is not None:
            self._cb_on_icmp_external_test_update(rtt_data)

        if lost:
            log_icmp.info(f'External ({rtt_data.host}): Timeout')
            self._icmp_external_test_seq_loss += 1
        else:
            if rtt_data.last >= self._icmp_late_time:
                log_icmp.info(f'External ({rtt_data.host}): {rtt_data.last:.3f} (late)')
            else:
                log_icmp.info(f'External ({rtt_data.host}): {rtt_data.last:.3f}')
            self._internet_connectivity = True
            self._icmp_external_test_seq_loss = 0

        if self._icmp_external_test_seq_loss >= self._icmp_max_seq_loss:
            self._internet_connectivity = False

    def _on_icmp_external_test_error(self, exception):
        log.debug('ICMP external test error', exc_info=True)
        log.error(f'ICMP external test error: {exception}')
        log_icmp.error('ICMP external test error', exc_info=True)
        log_icmp.info(f'External RTT data (on error):\n{self._icmp_external_test.rtt_data}')

        self._internet_connectivity = False

        time.sleep(self._icmp_interval)
        self._start_icmp_external_test(self._icmp_external_test_server)

    def _start_icmp_cs2_test(self, cs2_server_address):
        log.debug('Start ICMP CS2 test')

        if self._icmp_cs2_test and self._icmp_cs2_test.running:
            log.debug('Attempted to start ICMP CS2 test while previous test is running')
            return

        if self._cb_on_icmp_cs2_test_start is not None:
            self._cb_on_icmp_cs2_test_start(cs2_server_address)

        self._icmp_cs2_test = ICMPTest(
            cs2_server_address,
            self._icmp_interval,
            self._icmp_timeout,
            self._on_icmp_cs2_test_update,
            self._on_icmp_cs2_test_error
        )
        self._icmp_cs2_test_thread = threading.Thread(
            target=self._icmp_cs2_test.run,
            name='icmp-cs2-test',
            daemon=True
        )
        self._icmp_cs2_test_thread.start()

        log.info(f'ICMP CS2 test started to: {cs2_server_address}')
        log_icmp.info(f'ICMP CS2 test started to: {cs2_server_address}')

    def _stop_icmp_cs2_test(self):
        log.debug('Stop ICMP CS2 test')

        if self._icmp_cs2_test:
            self._icmp_cs2_test.stop()
            self._icmp_cs2_test_thread.join()

            log_icmp.info(f'CS2 RTT data:\n{self._icmp_cs2_test.rtt_data}')

    def _on_icmp_cs2_test_update(self, rtt_data, lost):
        if lost:
            log_icmp.info(f'CS2 ({rtt_data.host}): Timeout')
        elif rtt_data.last >= self._icmp_late_time:
            log_icmp.info(f'CS2 ({rtt_data.host}): {rtt_data.last} (late)')
        else:
            log_icmp.info(f'CS2 ({rtt_data.host}): {rtt_data.last}')

        if self._cb_on_icmp_cs2_test_update is not None:
            self._cb_on_icmp_cs2_test_update(rtt_data)

    def _on_icmp_cs2_test_error(self, exception):
        log.debug('ICMP CS2 test error', exc_info=True)
        log.error(f'ICMP CS2 test error: {exception}')
        log_icmp.error('ICMP CS2 test error', exc_info=True)
        log_icmp.info(f'CS2 RTT data (on error):\n{self._icmp_cs2_test.rtt_data}')

        time.sleep(self._icmp_interval)

        cs2_server = self._icmp_cs2_test_server
        if cs2_server is not None:
            self._start_icmp_cs2_test(cs2_server)

    def _run_traceroute(self, host, *args, **kwargs):
        log.debug('Run traceroute')

        if self._traceroute_thread and self._traceroute_thread.is_alive():
            log.debug('Attempted to run traceroute while thread is alive')
            return

        def wrapper():
            log.info(f"Traceroute to destination '{host}' started")
            log_tests.info(f"Traceroute to destination '{host}' started")
            self._traceroute.run()

            if self._traceroute.hops is not None:
                log.info(f"Traceroute to destination '{host}' finished")
                log_tests.info(f"Traceroute to destination '{host}' finished")
                self._on_traceroute_finish()

        self._traceroute = Traceroute(host, *args, on_error=self._on_traceroute_error, **kwargs)
        self._traceroute_thread = threading.Thread(target=wrapper, name='traceroute')
        self._traceroute_thread.start()

    def _wait_traceroute(self):
        log.debug('Wait traceroute')

        if self._traceroute_thread:
            self._traceroute_thread.join()

    def _on_traceroute_finish(self):
        log_tests.info(self._traceroute.format())

    def _on_traceroute_error(self, exception):
        log.debug(f"Traceroute to destination '{self._traceroute.host}' error", exc_info=True)
        log.error(f"Traceroute to destination '{self._traceroute.host}' error: {exception}")
        log_tests.error(f"Traceroute to destination '{self._traceroute.host}' error", exc_info=True)

    def _run_speed_test(self, *args, **kwargs):
        log.debug('Run speed test')

        if self._speed_test_thread and self._speed_test_thread.is_alive():
            log.debug('Attempted to run speed test while thread is alive')
            return

        def wrapper():
            log.info('Speed test started')
            log_tests.info('Speed test started')
            self._speed_test.run()

            if self._speed_test.results is not None:
                log.info('Speed test finished')
                log_tests.info('Speed test finished')
                self._on_speed_test_finish()

        self._speed_test = SpeedTest(*args, on_error=self._on_speed_test_error, **kwargs)
        self._speed_test_thread = threading.Thread(target=wrapper, name='speed-test')
        self._speed_test_thread.start()

    def _wait_speed_test(self):
        log.debug('Wait speed test')

        if self._speed_test_thread:
            self._speed_test_thread.join()

    def _on_speed_test_finish(self):
        log_tests.debug(self._speed_test.format())

    def _on_speed_test_error(self, exception):
        log.debug('Speed test error', exc_info=True)
        log.error(f'Speed test error: {exception}')
        log_tests.error('Speed test error', exc_info=True)

    def _start_event_log_listener(self):
        log.debug('Start event log listener')

        if self._event_log_listener_thread and self._event_log_listener_thread.is_alive():
            log.debug('Attempted to start event listener while thread is alive')
            return

        self._event_log_listener = EventLogListener(
            on_event=self._on_event_log_record,
            on_error=self._on_event_log_error
        )
        self._event_log_listener_thread = threading.Thread(
            target=self._event_log_listener.listen,
            name='event-log',
            daemon=True
        )
        self._event_log_listener_thread.start()

        log.info('Event log listener started')
        log_event_log.info('Event log listener started')

    def _stop_event_log_listener(self):
        log.debug('Stop event log listener')

        if self._event_log_listener:
            self._event_log_listener.stop()
            self._event_log_listener_thread.join()

    def _on_event_log_record(self, record):
        log_event_log.info(EventLogListener.format_record(record))

    def _on_event_log_error(self, exception):
        log.debug('Event log listener error', exc_info=True)
        log.error(f'Event log listener error: {exception}')
        log_event_log.error('Event log listener error', exc_info=True)


class NetworkInterface:

    def __init__(self, id_, name, mac_address, _ipv4_address, gateway_ipv4_address):
        self._id = id_
        self._name = name
        self._mac_address = mac_address
        self._ipv4_address = _ipv4_address
        self._gateway_ipv4_address = gateway_ipv4_address

    @property
    def id(self):
        return self._id

    @property
    def name(self):
        return self._name

    @property
    def mac_address(self):
        return self._mac_address

    @property
    def ipv4_address(self):
        return self._ipv4_address

    @property
    def gateway_ipv4_address(self):
        return self._gateway_ipv4_address

    @classmethod
    def from_default_gateway(cls):
        gateways = netifaces.gateways()
        default_gateway = gateways['default'].get(netifaces.AF_INET)
        if not default_gateway:
            raise ValueError('No default gateway present')
        gateway_ipv4_address = default_gateway[0]
        interface_id = default_gateway[1]

        addresses = netifaces.ifaddresses(interface_id)
        if not addresses:
            raise ValueError(f"No addresses found for interface '{interface_id}'")

        mac_addresses = addresses.get(netifaces.AF_LINK)
        if not mac_addresses:
            raise ValueError(f"No MAC addresses found for interface '{interface_id}'")
        interface_mac_address = mac_addresses[0]['addr'].replace(':', '-').upper()

        ipv4_addresses = addresses.get(netifaces.AF_INET)
        if not ipv4_addresses:
            raise ValueError(f"No IPv4 addresses found for interface '{interface_id}'")
        interface_ipv4_address = ipv4_addresses[0]['addr']

        interface_name = None
        for name, addresses in psutil.net_if_addrs().items():
            for address in addresses:
                if address.family == psutil.AF_LINK and interface_mac_address == address.address:
                    interface_name = name
                    break
            else:
                continue
            break
        if not interface_name:
            raise ValueError(f"No network names found by MAC address matching for interface '{interface_id}'")

        return cls(interface_id, interface_name, interface_mac_address, interface_ipv4_address, gateway_ipv4_address)

    def get_stats(self):
        interface_stats = psutil.net_if_stats().get(self._name)
        if not interface_stats:
            raise ValueError(f"No stats found for interface '{self._id}'")

        return {
            'up': interface_stats.isup,
            'speed': interface_stats.speed,
            'duplex': interface_stats.duplex,
            'duplex_str': (
                'full' if interface_stats.duplex == psutil.NIC_DUPLEX_FULL else
                'half' if interface_stats.duplex == psutil.NIC_DUPLEX_HALF else
                'unknown'
            ),
            'mtu': interface_stats.mtu
        }


class RTTData:

    def __init__(self, host):
        self._host = host

        self._initial_data = False
        self._last = None
        self._average = None
        self._minimum = None
        self._maximum = None
        self._jitter = None
        self._sent = 0
        self._received = 0

    def __str__(self):
        if not self._initial_data:
            return (
                f'Host: {self._host}\n'
                'Avg     Min     Max     Jitter  Sent Recv Loss\n'
                f'-       -       -       -       {self._sent:<4} {self._received:<4} {self.loss * 100:>5.1f}%'
            )

        return (
            f'Host: {self._host}\n'
            'Avg     Min     Max     Jitter  Sent Recv Loss\n'
            f'{self._average:<7.3f} {self._minimum:<7.3f} {self._maximum:<7.3f} {self._jitter:<7.3f} '
            f'{self._sent:<4} {self._received:<4} {self.loss * 100:>5.1f}%'
        )

    @property
    def host(self):
        return self._host

    @property
    def initial_data(self):
        return self._initial_data

    @property
    def last(self):
        return self._last

    @property
    def average(self):
        return self._average

    @property
    def minimum(self):
        return self._minimum

    @property
    def maximum(self):
        return self._maximum

    @property
    def jitter(self):
        return self._jitter

    @property
    def sent(self):
        return self._sent

    @property
    def received(self):
        return self._received

    @property
    def lost(self):
        return self._sent - self._received

    @property
    def loss(self):
        if self._sent == 0:
            return 0.0

        return 1 - self._received / self._sent

    def update(self, latency):
        if latency is None:
            self._sent += 1
            return

        latency = float(latency)
        self._sent += 1
        self._received += 1

        if not self._initial_data:
            self._last = latency
            self._average = latency
            self._minimum = latency
            self._maximum = latency
            self._jitter = 0.0
            self._initial_data = True
            return

        self._average = (self._sent * self._average + latency) / (self._sent + 1)

        if latency < self._minimum:
            self._minimum = latency

        if latency > self._maximum:
            self._maximum = latency

        last_jitter = abs(latency - self._last)
        self._jitter = (self._sent * self._jitter + last_jitter) / (self._sent + 1)

        self._last = latency


class ICMPTest:

    def __init__(self, host, interval=1.0, timeout=2.0, on_update=None, on_error=None):
        self._host = host
        self._rtt_data = None

        self._interval = interval
        self._timeout = timeout
        self._on_update = on_update
        self._on_error = on_error
        self._running = False

    @property
    def host(self):
        return self._host

    @property
    def rtt_data(self):
        return self._rtt_data

    @property
    def running(self):
        return self._running

    def run(self):
        self._running = True

        if icmplib.utils.is_hostname(self._host):
            try:
                address = icmplib.utils.resolve(self._host)[0]
            except icmplib.NameLookupError as e:
                self._error(e)
                return
        else:
            address = self._host

        self._rtt_data = RTTData(address)

        if icmplib.utils.is_ipv6_address(address):
            icmp_socket = icmplib.ICMPv6Socket
        else:
            icmp_socket = icmplib.ICMPv4Socket

        id_ = icmplib.utils.unique_identifier()
        count = 0

        try:
            with icmp_socket(None, True) as sock:

                while self._running:
                    if count > 0 and self._interval > 0:
                        time.sleep(self._interval)

                    request = icmplib.ICMPRequest(address, id_, count)

                    try:
                        sock.send(request)
                        reply = sock.receive(request, self._timeout)
                        reply.raise_for_status()
                        rtt = (reply.time - request.time) * 1000
                        self._update(rtt)

                    except icmplib.TimeoutExceeded:
                        self._update(None)

                    count += 1

        except icmplib.ICMPLibError as e:
            self._error(e)

    def stop(self):
        self._running = False

    def _update(self, current_rtt):
        self._rtt_data.update(current_rtt)

        if self._on_update is not None:
            self._on_update(self._rtt_data, lost=current_rtt is None)

    def _error(self, exception):
        self._running = False

        if self._on_error is not None:
            self._on_error(exception)
        raise


class Traceroute:

    def __init__(self, host, count=2, interval=0.0, timeout=2.0, max_hops=30, fast=False, on_error=None):
        self._host = host
        self._count = count
        self._interval = interval
        self._timeout = timeout
        self._max_hops = max_hops
        self._fast = fast
        self._on_error = on_error
        self._hops = None

    @property
    def host(self):
        return self._host

    @property
    def hops(self):
        return self._hops

    def run(self):
        self._hops = None

        try:
            self._hops = icmplib.traceroute(
                address=self._host,
                count=self._count,
                interval=self._interval,
                timeout=self._timeout,
                max_hops=self._max_hops,
                fast=self._fast
            )
        except icmplib.ICMPLibError as e:
            self._error(e)

    def format(self):
        if self._hops is None:
            raise ValueError('No traceroute data')

        lines = []
        lines.append(f'Host: {self._host}')
        lines.append('Hop Host            Avg     Min     Max     Jitter  Sent Recv Loss')

        last_distance = 0

        for hop in self._hops:
            for i in range(last_distance + 1, hop.distance):
                line = f'{i:<2}  *               -       -       -       -       -    -    100.0%'
                lines.append(line)

            last_distance = hop.distance

            loss_percent = (1 - hop.packets_received / hop.packets_sent) * 100
            line = (
                f'{hop.distance:<2}  {hop.address:<15} {hop.avg_rtt:<7.3f} {hop.min_rtt:<7.3f} {hop.max_rtt:<7.3f} '
                f'{hop.jitter:<7.3f} {hop.packets_sent:<4} {hop.packets_received:<4} {loss_percent:>5.1f}%'
            )
            lines.append(line)

        return '\n'.join(lines)

    def _error(self, exception):
        if self._on_error is not None:
            self._on_error(exception)
        raise


class SpeedTest:

    LATENCY_TEST = cfspeedtest.TestSpec(1, 20, 'latency', cfspeedtest.TestType.Down)
    SHORT_TESTS = (
        LATENCY_TEST,
        cfspeedtest.TestSpec(5_000_000, 4, '20MB', cfspeedtest.TestType.Down),
        cfspeedtest.TestSpec(2_500_000, 4, '10MB', cfspeedtest.TestType.Up)
    )
    LONG_TESTS = (
        LATENCY_TEST,
        cfspeedtest.TestSpec(25_000_000, 4, '100MB', cfspeedtest.TestType.Down),
        cfspeedtest.TestSpec(5_000_000, 4, '20MB', cfspeedtest.TestType.Up)
    )

    def __init__(self, tests=SHORT_TESTS, connection_timeout=10.0, read_timeout=10.0, on_error=None):
        self._tests = tests
        self._timeout = (connection_timeout, read_timeout)
        self._results = None
        self._on_error = on_error

    @property
    def results(self):
        return self._results

    def run(self):
        self._results = None

        try:
            speedtest = cfspeedtest.CloudflareSpeedtest(None, self._tests, self._timeout)
            raw_results = speedtest.run_all(megabits=True)
            self._results = {
                'tests': {
                    k: v.value for k, v in raw_results['tests'].items() if k != 'isp'
                },
                'meta': {
                    'colo': raw_results['meta']['location_code'].value
                }
            }
        except Exception as e:
            self._error(e)

    def format(self):
        if not self._results:
            raise ValueError('No speed test data')

        lines = [
            f'{k2}: {v2}'
            for k1, v1 in self._results.items()
            for k2, v2 in v1.items()
        ]
        return '\n'.join(lines)

    def _error(self, exception):
        if self._on_error is not None:
            self._on_error(exception)
        raise


class EventLogListener:

    LEVELS = {
        0: 'LogAlways',
        1: 'Critical',
        2: 'Error',
        3: 'Warning',
        4: 'Informational',
        5: 'Verbose'
    }

    def __init__(self, log='System', on_event=None, on_error=None):
        self._log = log
        self._on_event = on_event
        self._on_error = on_error
        self._subscription = None

    def listen(self):
        if self._subscription is not None:
            return

        try:
            self._subscription = win32evtlog.EvtSubscribe(
                self._log,
                win32evtlog.EvtSubscribeToFutureEvents,
                None,
                self._handler
            )
        except Exception as e:
            self._error(e)

    def stop(self):
        if self._subscription is not None:
            self._subscription.Close()
            self._subscription = None

    @staticmethod
    def format_record(record):
        timestamp = (
            datetime.datetime.fromisoformat(record['timestamp'])
            .astimezone()
            .strftime('%Y-%m-%dT%H:%M:%S.%fZ%z')
        )

        return (
            f"Timestamp (UTC): {timestamp}\n"
            f"Channel: {record['channel']}\n"
            f"Provider: {record['provider']}\n"
            f"Source: {record['source']}\n"
            f"Level: {EventLogListener.LEVELS[record['level']]}\n"
            f"OpCode: {record['opcode']}\n"

            f"Message:\n{record['message']}"
            if record['message'] is not None else

            f"Event Data:\n{'\n'.join(record['event_data_strings'])}"
            if record['event_data_strings'] is not None else

            f"User Data:\n{record['user_data']}"
            if record['user_data'] is not None else

            f"Raw:\n{record['raw']}"
        )

    def _error(self, exception):
        if self._on_error is not None:
            self._on_error(exception)
        raise

    def _handler(self, action, context, event):
        if action == win32evtlog.EvtSubscribeActionError:
            return

        try:
            event_data_xml = win32evtlog.EvtFormatMessage(None, event, win32evtlog.EvtFormatMessageXml)

            with xml.dom.minidom.parseString(event_data_xml) as dom:
                root = dom.getElementsByTagName('Event')[0]

                system = root.getElementsByTagName('System')[0]

                provider = system.getElementsByTagName('Provider')[0]
                provider_name = provider.getAttribute('Name')
                source_name = provider.getAttribute('EventSourceName') or None

                level = system.getElementsByTagName('Level')[0].firstChild.nodeValue
                opcode = system.getElementsByTagName('Opcode')[0].firstChild.nodeValue
                time_created = system.getElementsByTagName('TimeCreated')[0].getAttribute('SystemTime')
                channel = system.getElementsByTagName('Channel')[0].firstChild.nodeValue

                event_data_nodes = root.getElementsByTagName('EventData')
                event_data_strings = None
                if event_data_nodes:
                    event_data = event_data_nodes[0]
                    event_data_strings = [
                        (node.getAttribute('Name') or None, node.firstChild.nodeValue)
                        for node in event_data.childNodes
                        if node.firstChild.nodeType == node.TEXT_NODE
                    ]

                user_data_nodes = root.getElementsByTagName('UserData')
                user_data = None
                if user_data_nodes:
                    user_data = user_data_nodes[0].toprettyxml(' ' * 2)

                message = root.getElementsByTagName('Message')[0].nodeValue or None

        except Exception as e:
            self._error(e)
            return

        record = {
            'timestamp': time_created,
            'channel': channel,
            'provider': provider_name,
            'source': source_name,
            'level': level,
            'opcode': opcode,
            'message': message,
            'event_data_strings': event_data_strings,
            'user_data': user_data,
            'raw': root.toprettyxml(' ' * 2)
        }
        self._on_event(record)


class Netstat:

    @staticmethod
    def get_statistics():
        try:
            return subprocess.check_output(
                args=('netstat.exe', '-s'),
                stderr=subprocess.STDOUT,
                text=True,
                timeout=10.0
            )
        except (subprocess.TimeoutExpired, subprocess.CalledProcessError) as e:
            raise OSError(f'Failed to get netstat statistics: {e}') from e
