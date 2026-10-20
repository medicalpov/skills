#!/bin/sh
# Run the phi-scrub-slides test suite (Python standard library only).
set -e
cd "$(dirname "$0")/.."
python3 -m unittest discover -s tests -v
