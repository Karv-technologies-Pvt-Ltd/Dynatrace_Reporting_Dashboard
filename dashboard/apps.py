# dashboard/apps.py
from django.apps import AppConfig
import os

class DashboardConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "dashboard"

    def ready(self):
        """
        Start APScheduler:
        - In normal runserver child process (RUN_MAIN == "true")
        - In PyInstaller EXE (PYINSTALLER_RUNNING == "true")
        """
        if os.environ.get("RUN_MAIN") == "true" or os.environ.get("PYINSTALLER_RUNNING") == "true":
            from .scheduler import start_scheduler
            start_scheduler()
