#!/bin/bash

set -e

pyside6-uic --from-imports -o network_diagnostics/gui/generated/main_window.py resources/gui/main-window.ui
pyside6-uic --from-imports -o network_diagnostics/gui/generated/history_window.py resources/gui/history-window.ui
pyside6-rcc -o network_diagnostics/gui/generated/resources_rc.py resources/gui/resources.qrc
