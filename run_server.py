#!/usr/bin/env python
"""Startet das Dashboard (http://localhost:5100/). Port: config/server.yaml."""
from src.api.app import start

if __name__ == "__main__":
    start()
