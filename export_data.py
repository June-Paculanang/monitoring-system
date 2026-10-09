import os

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "TeacherMonitoringSystem.settings")

import django
django.setup()

from django.core.management import call_command

with open("local_data.json", "w", encoding="utf-8", newline="\n") as f:
    call_command(
        "dumpdata",
        "monitoring",
        "auth.user",
        "auth.group",
        natural_foreign=True,
        natural_primary=True,
        exclude=[
            "contenttypes",
            "auth.permission",
            "sessions",
        ],
        indent=2,
        stdout=f,
    )

print("SUCCESS: local_data.json created as UTF-8")