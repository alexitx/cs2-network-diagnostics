<div align="center">
  <img src="https://raw.githubusercontent.com/alexitx/cs2-network-diagnostics/master/docs/assets/icon.svg" height="80px"/>
  <h1>Network Diagnostics</h1>
</div>


## About

This is an application indented to help with diagnosing the source of internet connection issues that I originally made
for a friend, but can be used on any Windows system. It optionally supports parsing Counter-Strike 2 matchmaking
connection info and monitoring the connection to the relay, part of Valve's Steam Datagram Relay network.

Administrator privileges are required when running the application because it adds firewall rules to allow ICMPv4 Time
Exceeded packets for traceroute tests to be performed.

Latest Windows 10 x64 is supported. Windows 11 should work, but is not tested.


## Features

- Monitoring the current active network interface
- Detecting internet connectivity interruptions
- Monitoring latency and packet loss
- Performing traceroute and speed tests to Cloudflare's network
- Monitoring Windows event log
- Parsing and monitoring CS2 matchmaking network connection
- Collection of per-protocol statistics


## Installation

Download the [latest release][releases] and extract the archive to a location of your choice.


## Building

Requirements:

- [Git Bash][git-scm]
- [Python][python] 3.12 x64

1. Create a virtual environment

    ```sh
    python -m virtualenv venv
    source ./venv/Scripts/activate
    ```

2. Install dependencies

    ```sh
    pip install -r requirements.txt -r requirements-dev.txt
    ```

3. Compile resources and build

    ```sh
    ./scripts/compile-ui.sh
    ./scripts/build.sh
    ```

Final build is located in `pyinstaller-dist`.


## License

GNU General Public License v3.0. See [LICENSE][license] for more information.


[releases]: https://github.com/alexitx/cs2-network-diagnostics/releases
[git-scm]: https://git-scm.com/downloads
[python]: https://www.python.org/downloads
[license]: https://raw.githubusercontent.com/alexitx/cs2-network-diagnostics/master/LICENSE
