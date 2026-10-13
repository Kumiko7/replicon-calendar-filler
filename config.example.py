# Personal configuration for calendar_server.py.
# config.py is gitignored - edit it freely. config.example.py is the committed template.
from datetime import time

PROJECT_MAPPING = {
    # Outlook Event Titles -> Project Names
    # An exact title match is used first, otherwise the first entry whose title is contained in the event title.
    "Event Name": "Project Name",

    # Custom Rule Titles -> Project Names
    "Personal administrative work": "Project Name",
    "Replicon timesheets, personal administrative work": "Project Name",
    # Add more title-to-project mappings here.
    # If a title is not found here, the project will be an empty string "".
}

CUSTOM_RULES = [
    {
        "title": "Break Time",
        "start_time": time(12, 0),
        "end_time": time(13, 0),
        "weekdays": [0, 1, 2, 3, 4]
    },
    {
        "title": "Personal administrative work",
        "start_time": time(9, 0),
        "end_time": time(10, 30),
        "weekdays": [0]  # Monday
    },
    {
        "title": "Replicon timesheets, personal administrative work",
        "start_time": time(17, 0),
        "end_time": time(18, 0),
        "weekdays": [4]  # Friday
    }
]

WORK_START_TIME = time(9, 0)
WORK_END_TIME = time(18, 0)
WORK_WEEKDAYS = [0, 1, 2, 3, 4]

# If True, timed events are trimmed to WORK_START_TIME-WORK_END_TIME on work days,
# and anything entirely outside working hours (or on a weekend/day off) is dropped.
# All-day events are always trimmed to working hours.
CLAMP_TO_WORK_HOURS = True

# Events marked "Free" are always skipped. Set this to True to also skip "Tentative"
# events (note: meeting invites you haven't answered yet also show as Tentative).
SKIP_TENTATIVE_EVENTS = False
