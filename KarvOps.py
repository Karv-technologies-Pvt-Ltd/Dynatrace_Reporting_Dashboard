import sys
import webbrowser
import time
import os
from pathlib import Path
import django
from django.core.management import execute_from_command_line
from threading import Thread

BASE_DIR = Path(__file__).resolve().parent

def wait_for_server_and_open():
    import socket
    while True:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.connect(("127.0.0.1", 8000))
            s.close()
            break
        except OSError:
            time.sleep(0.5)
    webbrowser.open("http://127.0.0.1:8000/")

def start_django():
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "dynatrace_tracker.settings")

    # Detect EXE mode (needed for DB extraction)
    os.environ["PYINSTALLER_RUNNING"] = "true"

    django.setup()

    execute_from_command_line([
        "",
        "runserver",
        "127.0.0.1:8000",
        "--noreload",
        "--nothreading"
    ])


if __name__ == "__main__":
    print("Starting Django…")

    # Start browser watcher thread
    Thread(target=wait_for_server_and_open).start()

    # Start Django server
    start_django()
    print("Django server has stopped.")
