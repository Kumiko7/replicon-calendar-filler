import win32com.client
import json
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timedelta, time
import os
import shutil
import sys

# --- Configuration ---
# Settings live in config.py (gitignored). It is created from config.example.py on first run.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if not os.path.exists(os.path.join(SCRIPT_DIR, "config.py")):
    shutil.copy(os.path.join(SCRIPT_DIR, "config.example.py"), os.path.join(SCRIPT_DIR, "config.py"))
    print("Created config.py from config.example.py - edit it to add your project mappings.")
sys.path.insert(0, SCRIPT_DIR)
from config import (
    PROJECT_MAPPING, CUSTOM_RULES, WORK_START_TIME, WORK_END_TIME, WORK_WEEKDAYS,
    CLAMP_TO_WORK_HOURS, SKIP_TENTATIVE_EVENTS,
)

# Outlook constants
OL_FREE, OL_TENTATIVE, OL_BUSY, OL_OUT_OF_OFFICE = 0, 1, 2, 3
OL_RESPONSE_DECLINED = 4
OL_MEETING_CANCELED, OL_MEETING_RECEIVED_AND_CANCELED = 5, 7

def get_public_holidays(start_date_str, end_date_str):
    """
    Fetches days off from the Outlook calendar.
    Days off are all-day events with a "Free" (public holiday) or
    "Out of Office" (vacation/leave) busy status. Multi-day events mark every day they cover.
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        calendar = outlook.GetDefaultFolder(9)
        appointments = calendar.Items
        appointments.Sort("[Start]")
        appointments.IncludeRecurrences = True
    except Exception as e:
        print(f"Error connecting to Outlook: {e}")
        return []

    start_dt = datetime.strptime(start_date_str, '%Y-%m-%d')
    end_dt = datetime.strptime(end_date_str, '%Y-%m-%d').replace(hour=23, minute=59, second=59)

    start_date_filter = start_dt.strftime('%m/%d/%Y %I:%M %p')
    end_date_filter = end_dt.strftime('%m/%d/%Y %I:%M %p')

    # Any overlap with the period, so multi-day leave that started earlier is still caught
    restriction = f"[Start] <= '{end_date_filter}' AND [End] >= '{start_date_filter}'"
    restricted_appointments = appointments.Restrict(restriction)

    holidays = set()
    for appointment in restricted_appointments:
        if appointment.AllDayEvent and appointment.BusyStatus in (OL_FREE, OL_OUT_OF_OFFICE):
            day = appointment.Start.date()
            last_day = max(appointment.End.date() - timedelta(days=1), day)  # End is exclusive (midnight)
            while day <= last_day:
                if start_dt.date() <= day <= end_dt.date():
                    holidays.add(day)
                    print(f"[DAY OFF] {day}: {appointment.Subject}")
                day += timedelta(days=1)
    return holidays

def get_outlook_calendar(start_date_str, end_date_str):
    """
    Fetches calendar events from Outlook and assigns a project based on the title.
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        calendar = outlook.GetDefaultFolder(9)
        appointments = calendar.Items
        appointments.Sort("[Start]")
        appointments.IncludeRecurrences = True
    except Exception as e:
        print(f"Error connecting to Outlook: {e}")
        return []

    start_dt = datetime.strptime(start_date_str, '%Y-%m-%d')
    end_dt = datetime.strptime(end_date_str, '%Y-%m-%d').replace(hour=23, minute=59, second=59)

    start_date_filter = start_dt.strftime('%m/%d/%Y %I:%M %p')
    end_date_filter = end_dt.strftime('%m/%d/%Y %I:%M %p')

    restriction = f"[Start] >= '{start_date_filter}' AND [End] <= '{end_date_filter}'"
    restricted_appointments = appointments.Restrict(restriction)

    events = []
    for appointment in restricted_appointments:
        subject = appointment.Subject or ""

        # Free events (incl. public holidays) and all-day Out of Office (handled as days off)
        if appointment.BusyStatus == OL_FREE:
            continue
        if appointment.AllDayEvent and appointment.BusyStatus == OL_OUT_OF_OFFICE:
            continue
        if SKIP_TENTATIVE_EVENTS and appointment.BusyStatus == OL_TENTATIVE:
            continue

        if subject.startswith("Canceled: ") or appointment.MeetingStatus in (OL_MEETING_CANCELED, OL_MEETING_RECEIVED_AND_CANCELED):
            continue

        if appointment.ResponseStatus == OL_RESPONSE_DECLINED:
            print(f"[SKIPPED] Declined: \"{subject}\"")
            continue

        project = PROJECT_MAPPING.get(subject, "")
        if not project:
            for key, cproject in PROJECT_MAPPING.items():
                if key in subject:
                    project = cproject
                    break

        required_attendees = [name.strip() for name in str(appointment.RequiredAttendees).split(';') if name.strip()]
        optional_attendees = [name.strip() for name in str(appointment.OptionalAttendees).split(';') if name.strip()]

        events.append({
            "subject": subject,
            "project": project,
            "start": appointment.Start.replace(tzinfo=None).isoformat(),
            "end": appointment.End.replace(tzinfo=None).isoformat(),
            "organizer": str(appointment.Organizer),
            "required_attendees": required_attendees,
            "optional_attendees": optional_attendees,
            "all_day": bool(appointment.AllDayEvent),
        })
    return events

def fit_events_to_workdays(events, public_holidays):
    """
    Splits events that span multiple days into one piece per day, and trims them to
    working hours (always for all-day events, for timed events if CLAMP_TO_WORK_HOURS).
    Pieces that fall entirely outside working hours, on non-work days or on days off are dropped.
    """
    fitted = []
    for event in events:
        event_start = datetime.fromisoformat(event['start'])
        event_end = datetime.fromisoformat(event['end'])
        clamp = CLAMP_TO_WORK_HOURS or event.get('all_day')

        day = event_start.date()
        while datetime.combine(day, time(0, 0)) < event_end:
            piece_start = max(event_start, datetime.combine(day, time(0, 0)))
            piece_end = min(event_end, datetime.combine(day + timedelta(days=1), time(0, 0)))

            if clamp:
                if day.weekday() not in WORK_WEEKDAYS or day in public_holidays:
                    piece_end = piece_start  # drop
                else:
                    piece_start = max(piece_start, datetime.combine(day, WORK_START_TIME))
                    piece_end = min(piece_end, datetime.combine(day, WORK_END_TIME))

            if piece_end > piece_start:
                piece = dict(event)
                piece['start'] = piece_start.isoformat()
                piece['end'] = piece_end.isoformat()
                fitted.append(piece)
            elif not event.get('all_day'):
                print(f"[SKIPPED] Outside working hours / not a work day: \"{event['subject']}\" [{event['start']} -> {event['end']}]")

            day += timedelta(days=1)
    return fitted

def add_custom_rules(start_date_str, end_date_str, public_holidays):
    """
    Generates recurring events and assigns a project based on the rule title,
    skipping public holidays.
    """
    start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
    end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
    custom_events = []
    
    current_date = start_date
    while current_date <= end_date:
        if current_date in public_holidays:
            current_date += timedelta(days=1)
            continue
            
        for rule in CUSTOM_RULES:
            if current_date.weekday() in rule["weekdays"]:
                title = rule["title"]
                project = PROJECT_MAPPING.get(title, "")

                custom_events.append({
                    "subject": title,
                    "project": project,
                    "start": datetime.combine(current_date, rule["start_time"]).isoformat(),
                    "end": datetime.combine(current_date, rule["end_time"]).isoformat(),
                    "organizer": "System Rule",
                    "required_attendees": [],
                    "optional_attendees": [],
                })
        current_date += timedelta(days=1)
    return custom_events

def split_events_around_contained_events(events):
    """
    Splits any event that completely surrounds another event (calendar event or custom rule)
    into a part before and a part after it, so the surrounding event resumes afterwards.
    e.g. Workshop 9:00-17:00 containing Meeting 10:00-11:00
         -> Workshop 9:00-10:00, Meeting 10:00-11:00, Workshop 11:00-17:00
    """
    processed_events = [dict(e) for e in events]

    changed = True
    while changed:
        changed = False
        for outer in processed_events:
            outer_start = datetime.fromisoformat(outer['start'])
            outer_end = datetime.fromisoformat(outer['end'])
            for inner in processed_events:
                if inner is outer:
                    continue
                inner_start = datetime.fromisoformat(inner['start'])
                inner_end = datetime.fromisoformat(inner['end'])

                if outer_start < inner_start and outer_end > inner_end:
                    part1 = dict(outer, end=inner_start.isoformat())
                    part2 = dict(outer, start=inner_end.isoformat())
                    processed_events.remove(outer)
                    processed_events.extend([part1, part2])
                    print(f"[SPLIT] \"{outer['subject'] or '(No Subject)'}\" around \"{inner['subject'] or '(No Subject)'}\" "
                          f"[{inner_start.strftime('%Y-%m-%d %H:%M')} -> {inner_end.strftime('%H:%M')}]")
                    changed = True
                    break
            if changed:
                break

    return processed_events

def resolve_overlaps(events):
    """
    Detects overlapping events, prints details to console, and truncates
    events so they do not overlap while allowing both to have time if possible.
    """
    if not events:
        return []

    event_list = [dict(e) for e in events]
    overlap_found = False

    while True:
        # Sort events chronologically by start time, then end time
        event_list.sort(key=lambda x: (datetime.fromisoformat(x['start']), datetime.fromisoformat(x['end'])))
        
        # Remove zero or negative duration events
        event_list = [
            e for e in event_list 
            if datetime.fromisoformat(e['end']) > datetime.fromisoformat(e['start'])
        ]

        overlap_detected_in_pass = False

        for i in range(len(event_list) - 1):
            e1 = event_list[i]
            e2 = event_list[i + 1]

            s1 = datetime.fromisoformat(e1['start'])
            e1_end = datetime.fromisoformat(e1['end'])
            s2 = datetime.fromisoformat(e2['start'])
            e2_end = datetime.fromisoformat(e2['end'])

            if s2 < e1_end:
                overlap_found = True
                overlap_detected_in_pass = True

                sub1 = e1.get('subject') or '(No Subject)'
                sub2 = e2.get('subject') or '(No Subject)'

                print("\n" + "=" * 60)
                print("[OVERLAP DETECTED]")
                print(f"  Event 1: \"{sub1}\" [{e1['start']} -> {e1['end']}]")
                print(f"  Event 2: \"{sub2}\" [{e2['start']} -> {e2['end']}]")

                if s1 == s2:
                    # Both events start at the same time: divide the shared duration
                    common_end = min(e1_end, e2_end)
                    midpoint = s1 + (common_end - s1) / 2
                    e1['end'] = midpoint.isoformat()
                    e2['start'] = midpoint.isoformat()
                    print(f"  -> Same start time: Divided time at {midpoint.strftime('%H:%M:%S')}")
                    print(f"     Adjusted Event 1: [{e1['start']} -> {e1['end']}]")
                    print(f"     Adjusted Event 2: [{e2['start']} -> {e2['end']}]")
                else:
                    # Event 1 started first: end Event 1 early at Event 2's start
                    e1['end'] = s2.isoformat()
                    print(f"  -> Truncated Event 1 to end early at {s2.isoformat()}")
                    print(f"     Adjusted Event 1: [{e1['start']} -> {e1['end']}]")
                print("=" * 60)

                # Restart the loop to re-sort and verify remaining overlaps
                break

        if not overlap_detected_in_pass:
            break

    if not overlap_found:
        print("\n[INFO] No overlapping events found.")

    return event_list

def fill_workday_gaps(sorted_events, start_date_str, end_date_str, public_holidays):
    """
    Creates placeholder events (with a blank project) for gaps in the workday,
    skipping public holidays.
    """
    placeholder_events = []
    start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
    end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()

    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() not in WORK_WEEKDAYS or current_date in public_holidays:
            current_date += timedelta(days=1)
            continue

        workday_start = datetime.combine(current_date, WORK_START_TIME)
        workday_end = datetime.combine(current_date, WORK_END_TIME)
        
        events_for_today = [
            e for e in sorted_events 
            if datetime.fromisoformat(e['start']).date() == current_date
        ]

        cursor = workday_start

        for event in events_for_today:
            event_start = datetime.fromisoformat(event['start'])
            event_end = datetime.fromisoformat(event['end'])
            
            effective_start = max(event_start, workday_start)
            effective_end = min(event_end, workday_end)
            
            if effective_end <= cursor:
                continue

            if effective_start > cursor:
                placeholder_events.append({
                    "subject": "",
                    "project": "",
                    "start": cursor.isoformat(),
                    "end": effective_start.isoformat(),
                    "organizer": "Placeholder",
                    "required_attendees": [],
                    "optional_attendees": [],
                })
            
            cursor = max(cursor, effective_end)

        if cursor < workday_end:
            placeholder_events.append({
                "subject": "",
                "project": "",
                "start": cursor.isoformat(),
                "end": workday_end.isoformat(),
                "organizer": "Placeholder",
                "required_attendees": [],
                "optional_attendees": [],
            })

        current_date += timedelta(days=1)
        
    return placeholder_events

class CalendarAPIHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed_path = urlparse(self.path)
        if parsed_path.path == '/calendar':
            query_params = parse_qs(parsed_path.query)
            
            today = datetime.now().date()
            from_date = query_params.get('from', [str(today)])[0]
            to_date = query_params.get('to', [str(today)])[0]

            try:
                public_holidays = get_public_holidays(from_date, to_date)
                outlook_events = get_outlook_calendar(from_date, to_date)
                outlook_events = fit_events_to_workdays(outlook_events, public_holidays)
                custom_events = add_custom_rules(from_date, to_date, public_holidays)

                all_events = split_events_around_contained_events(outlook_events + custom_events)

                # Resolve all overlaps between Outlook events & Custom Rules
                all_events = resolve_overlaps(all_events)

                placeholder_events = fill_workday_gaps(all_events, from_date, to_date, public_holidays)

                all_events.extend(placeholder_events)
                all_events.sort(key=lambda event: event['start'])

                self.send_response(200)
                self.send_header('Content-type', 'application/json; charset=utf-8')
                self.end_headers()
                
                json_response = json.dumps(all_events, indent=4, ensure_ascii=False)
                self.wfile.write(json_response.encode('utf-8'))

            except Exception as e:
                self.send_response(500)
                self.send_header('Content-type', 'application/json; charset=utf-8')
                self.end_headers()
                error_response = json.dumps({"error": str(e)}, ensure_ascii=False)
                self.wfile.write(error_response.encode('utf-8'))
        else:
            self.send_response(404)
            self.send_header('Content-type', 'application/json; charset=utf-8')
            self.end_headers()
            not_found_response = json.dumps({"error": "Not Found"}, ensure_ascii=False)
            self.wfile.write(not_found_response.encode('utf-8'))

def run_server(server_class=HTTPServer, handler_class=CalendarAPIHandler, port=8000):
    server_address = ('localhost', port)
    httpd = server_class(server_address, handler_class)
    print(f"Starting server on http://localhost:{port}")
    httpd.serve_forever()

if __name__ == '__main__':
    run_server()