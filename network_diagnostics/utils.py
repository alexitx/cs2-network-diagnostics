import ctypes
import re
import subprocess
import tempfile
from pathlib import Path

import filelock
import platformdirs


def _check_firewall_icmp_rule():
    args = (
        'C:\\Windows\\System32\\netsh.exe',
        'advfirewall',
        'firewall',
        'show',
        'rule',
        'name=Network Diagnostics - Allow ICMPv4 Time Exceeded'
    )
    try:
        subprocess.check_output(
            args=args,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=10
        )
        return True
    except (subprocess.TimeoutExpired, subprocess.CalledProcessError):
        return False


def create_firewall_rules():
    if _check_firewall_icmp_rule():
        return

    args = (
        'C:\\Windows\\System32\\netsh.exe',
        'advfirewall',
        'firewall',
        'add',
        'rule',
        'name=Network Diagnostics - Allow ICMPv4 Time Exceeded',
        'dir=in',
        'action=allow',
        'protocol=icmpv4:11,any'
    )
    try:
        subprocess.check_output(
            args=args,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=10
        )
    except (subprocess.TimeoutExpired, subprocess.CalledProcessError) as e:
        raise OSError(f"Failed to create firewall rule. Exit code: {e.returncode}, Output: '{e.output.strip()}'") from e


def remove_firewall_rules():
    if not _check_firewall_icmp_rule():
        return

    args = (
        'C:\\Windows\\System32\\netsh.exe',
        'advfirewall',
        'firewall',
        'delete',
        'rule',
        'name=Network Diagnostics - Allow ICMPv4 Time Exceeded'
    )
    try:
        subprocess.check_output(
            args=args,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=10
        )
    except (subprocess.TimeoutExpired, subprocess.CalledProcessError) as e:
        raise OSError(f"Failed to remove firewall rule. Exit code: {e.returncode}, Output: '{e.output.strip()}'") from e


def is_running_as_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() == 1
    except Exception:
        return False


_data_path = platformdirs.user_data_path('network-diagnostics', False)
_logs_path = _data_path / 'logs'
_diagnostics_path = _data_path / 'diagnostics'


def create_app_dirs():
    try:
        _data_path.mkdir(parents=True, exist_ok=True)
        _logs_path.mkdir(parents=True, exist_ok=True)
        _diagnostics_path.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        raise OSError(e.errno, 'Failed to create application directory', str(_data_path)) from e


def get_data_dir():
    return _data_path


def get_logs_dir():
    return _logs_path


def get_diagnostics_dir():
    return _diagnostics_path


def get_file_lock():
    temp_dir = Path(tempfile.gettempdir())
    lock_file = temp_dir / 'network-diagnostics.lock'
    lock = filelock.FileLock(lock_file, blocking=False)
    return lock


_re_cs2_server = re.compile(
    r"^\[Networking\]\s+Remote host is in data center '(?P<location>\w+)'$",
    re.MULTILINE
)
_re_cs2_primary_router = re.compile(
    r"^\[Networking\]\s+Primary router: (?P<location>\w+#\d+) \((?P<host>(?:\d{1,3}\.){3}\d{1,3}):(?P<port>\d+)\)\s+Ping = (?P<latency_front>-?\d+)\+(?P<latency_back>-?\d+)=-?\d+ \(front\+back=total\)$",
    re.MULTILINE
)
_re_cs2_backup_router = re.compile(
    r"^\[Networking\]\s+Backup router: (?P<location>\w+#\d+) \((?P<host>(?:\d{1,3}\.){3}\d{1,3}):(?P<port>\d+)\)\s+Ping = (?P<latency_front>-?\d+)\+(?P<latency_back>-?\d+)=-?\d+ \(front\+back=total\)$",
    re.MULTILINE
)


def parse_cs2_ncs(text):
    cs2_server_match = _re_cs2_server.search(text)
    if not cs2_server_match:
        raise ValueError('Could not find CS2 server')

    server_location = cs2_server_match.groupdict()['location']

    cs2_primary_router_match = _re_cs2_primary_router.search(text)
    if not cs2_primary_router_match:
        raise ValueError('Could not find CS2 primary relay')

    pr_location, pr_address, pr_port, pr_latency_front, pr_latency_back = cs2_primary_router_match.groupdict().values()

    cs2_backup_router_match = _re_cs2_backup_router.search(text)
    if not cs2_backup_router_match:
        raise ValueError('Could not find CS2 backup relay')

    br_location, br_address, br_port, br_latency_front, br_latency_back = cs2_backup_router_match.groupdict().values()

    return {
        'server_location': server_location,
        'primary_relay_location': pr_location,
        'primary_relay_address': pr_address,
        'primary_relay_port': int(pr_port),
        'primary_relay_latency_front': int(pr_latency_front),
        'primary_relay_latency_back': int(pr_latency_back),
        'backup_relay_location': br_location,
        'backup_relay_address': br_address,
        'backup_relay_port': int(br_port),
        'backup_relay_latency_front': int(br_latency_front),
        'backup_relay_latency_back': int(br_latency_back)
    }


def open_path_in_explorer(path):
    try:
        subprocess.run(
            args=('explorer.exe', path),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=10
        )
    except (subprocess.TimeoutExpired, subprocess.CalledProcessError) as e:
        raise OSError(f'Failed to open path in explorer: {e}') from e
