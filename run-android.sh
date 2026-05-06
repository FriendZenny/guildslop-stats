#!/data/data/com.termux/files/usr/bin/sh
cd "$(dirname "$0")"
if [ ! -d ".venv" ]; then
    python -m venv .venv
    .venv/bin/pip install -r requirements.txt
fi
.venv/bin/python scrape.py
