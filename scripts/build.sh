#!/bin/bash

set -e

rm -rf pyinstaller-dist/ pyinstaller-build/
pyinstaller --distpath pyinstaller-dist --workpath pyinstaller-build -y --clean build.spec
