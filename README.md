# replicon-calendar-filler
Python and Tampermonkey scripts for filling Replicon timesheets from Outlook calendar data.

1. Add Replicon Calendar Importer.user.js as a Tampermonkey script.
2. Run calendar_server.bat
3. In Replicon press the button to pull the data from the Outlook calendar automatically.

## Configuration
Personal settings live in `config.py`, which is gitignored. On first run it is created from `config.example.py`.

- `PROJECT_MAPPING`: Outlook event titles (or custom rule titles) -> Replicon project names.
  An exact title match is used first, otherwise the first entry contained in the event title.
- `CUSTOM_RULES`: recurring blocks added every week (lunch break, admin time, ...).
- `WORK_START_TIME` / `WORK_END_TIME` / `WORK_WEEKDAYS`: working hours. Gaps are filled with blank rows.
- `CLAMP_TO_WORK_HOURS`: trim events to working hours and drop ones outside them.
- `SKIP_TENTATIVE_EVENTS`: also skip events marked Tentative.

## What gets skipped
- Public holidays and leave in Outlook: all-day events marked Free or Out of Office.
- Days with time off already booked in Replicon (e.g. Annual Leave).
- Declined and cancelled meetings, and events marked Free.
- If a project can't be found in Replicon it is left blank and the row is outlined in red.
