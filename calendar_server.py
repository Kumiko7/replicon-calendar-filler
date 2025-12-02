import win32com.client
import json
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timedelta, time

# --- Configuration ---
# --- New Project Mapping Configuration ---
PROJECT_MAPPING = {
    # Outlook Event Titles -> Project Names
    "Event Name": "Project Name",

    # Custom Rule Titles -> Project Names
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
        "end_time": time(10, 0),
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

def get_public_holidays(start_date_str, end_date_str):
    """
    Fetches public holidays from the Outlook calendar.
    Holidays are identified as all-day events with a "Free" busy status.
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

    holidays = []
    for appointment in restricted_appointments:
        if appointment.AllDayEvent and appointment.BusyStatus == 0: # 0 represents "Free"
            holidays.append(appointment.Start.date())
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
        # Ignore all-day events that are holidays
        if appointment.AllDayEvent and appointment.BusyStatus == 0:
            continue
            
        subject = appointment.Subject
        
        if subject.startswith("Canceled: "):
            continue
        
        project = PROJECT_MAPPING.get(subject, "")
        
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
        })
    return events

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
                custom_events = add_custom_rules(from_date, to_date, public_holidays)
                
                all_events = outlook_events + custom_events
                all_events.sort(key=lambda event: event['start'])

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