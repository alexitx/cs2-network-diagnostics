#!/bin/bash

set -e

app_version="$(python -c 'from network_diagnostics.version import __version__; print(__version__)')"
app_dir='Network Diagnostics'
archive="Network-Diagnostics_v${app_version}_win-x64.7z"

rm -rf dist/
mkdir -p dist/
cd dist/

mkdir -p "$app_dir/"
cp -r ../pyinstaller-dist/* ../LICENSE ../README.md "$app_dir/"
7z a -t7z -mx9 -mtm=off -mtc=off -mta=off "$archive" "$app_dir"
