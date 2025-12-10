#By Matheus Caella Santis
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import calendar
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
import json
import tkinter as tk
from tkinter import simpledialog, messagebox
import os
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


SAVE_FILE = "calendar_data.json"

# ---------- Google Calendar integration ---------- #
SCOPES = ['https://www.googleapis.com/auth/calendar.events', 'https://www.googleapis.com/auth/calendar']

CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.pickle'
APP_CALENDAR_SUMMARY = "Work Calendar (LocalApp)"
APP_CALENDAR_ID_FILE = 'app_calendar_id.txt'

def _is_weekday(date_obj: datetime) -> bool:
    return date_obj.weekday() < 5

def _coerce_to_weekday(date_obj: datetime) -> datetime:
    while date_obj.weekday() >= 5:
        date_obj += timedelta(days=1)
    return date_obj

def _start_of_week(date_obj: datetime) -> datetime:
    return date_obj - timedelta(days=date_obj.weekday())

def _is_holiday(date_obj: datetime) -> bool:
    """Check if a date is a US holiday. Works for any year."""
    year = date_obj.year
    month = date_obj.month
    day = date_obj.day
    
    if (month == 1 and day == 1):  # New Year's Day
        return True
    if (month == 6 and day == 19):  # Juneteenth
        return True
    if (month == 7 and day == 4):  # Independence Day
        return True
    if (month == 11 and day == 11):  # Veterans Day
        return True
    if (month == 12 and day == 24):  # Christmas Eve
        return True
    if (month == 12 and day == 25):  # Christmas Day
        return True
    
    if month == 1:
        first_weekday = datetime(year, 1, 1).weekday()  # 0=Mon, 6=Sun
        days_to_first_monday = (7 - first_weekday) % 7
        if days_to_first_monday == 7:
            days_to_first_monday = 0
        first_monday = 1 + days_to_first_monday
        mlk_day = first_monday + 14  # 3rd Monday
        if day == mlk_day:
            return True
    
    # Presidents Day - 3rd Monday of February
    if month == 2:
        first_weekday = datetime(year, 2, 1).weekday()
        days_to_first_monday = (7 - first_weekday) % 7
        if days_to_first_monday == 7:
            days_to_first_monday = 0
        first_monday = 1 + days_to_first_monday
        presidents_day = first_monday + 14  # 3rd Monday
        if day == presidents_day:
            return True
    
    # Good Friday - Friday before Easter
    if month == 4:
        # Easter calculation
        a = year % 19
        b = year // 100
        c = year % 100
        d = b // 4
        e = b % 4
        f = (b + 8) // 25
        g = (b - f + 1) // 3
        h = (19 * a + b - d - g + 15) % 30
        i = c // 4
        k = c % 4
        l = (32 + 2 * e + 2 * i - h - k) % 7
        m = (a + 11 * h + 22 * l) // 451
        easter_month = (h + l - 7 * m + 114) // 31
        easter_day = ((h + l - 7 * m + 114) % 31) + 1
        if easter_month == 4:
            good_friday = datetime(year, easter_month, easter_day) - timedelta(days=2)
            if date_obj.date() == good_friday.date():
                return True
    
    # Memorial Day - Last Monday of May
    if month == 5:
        last_day = 31
        last_weekday = datetime(year, 5, last_day).weekday()
        days_back = last_weekday
        last_monday = last_day - days_back
        if day == last_monday:
            return True
    
    # Labor Day - First Monday of September
    if month == 9:
        first_weekday = datetime(year, 9, 1).weekday()
        days_to_first_monday = (7 - first_weekday) % 7
        if days_to_first_monday == 7:
            days_to_first_monday = 0
        first_monday = 1 + days_to_first_monday
        if day == first_monday:
            return True
    
    # Thanksgiving - 4th Thursday of November
    if month == 11:
        first_weekday = datetime(year, 11, 1).weekday()  # 0=Mon, 3=Thu, 6=Sun
        if first_weekday <= 3:
            days_to_first_thursday = 3 - first_weekday
        else:
            days_to_first_thursday = 7 - first_weekday + 3
        first_thursday = 1 + days_to_first_thursday
        thanksgiving = first_thursday + 21  # 4th Thursday
        if day == thanksgiving:
            return True
    
    return False

def _skip_holidays(date_obj: datetime) -> datetime:
    """Move a date forward to skip holidays and weekends. Returns next valid weekday."""
    result = date_obj
    max_iterations = 14
    iterations = 0
    
    while iterations < max_iterations:
        # Skip weekends first
        while result.weekday() >= 5:
            result += timedelta(days=1)
            iterations += 1
            if iterations >= max_iterations:
                break
        
        # Check if it's a holiday - if so, move forward and check again
        if _is_holiday(result):
            result += timedelta(days=1)
            iterations += 1
            continue  # Go back to check if this new date is a weekend or holiday
        
        # If we get here, it's a weekday and not a holiday
        return result
    
    return _coerce_to_weekday(result)

def ensure_google_credentials():
    # Check if credentials.json is missing required fields, prompt user if so
    creds_path = CREDENTIALS_FILE
    needs_input = False
    creds_data = None
    if not os.path.exists(creds_path):
        needs_input = True
    else:
        with open(creds_path, 'r') as f:
            try:
                creds_data = json.load(f)
                installed = creds_data.get('installed', {})
                if not installed.get('client_id') or not installed.get('client_secret') or not installed.get('project_id'):
                    needs_input = True
            except Exception:
                needs_input = True
    if needs_input:
        # Use Tkinter to prompt for credentials
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Google Credentials Required", "Please enter your Google Calendar API credentials. You can get these from the Google Cloud Console.")
        client_id = simpledialog.askstring("Google Credentials", "Enter your client_id:")
        if not client_id:
            root.destroy()
            raise SystemExit(1)
        client_secret = simpledialog.askstring("Google Credentials", "Enter your client_secret:")
        if not client_secret:
            root.destroy()
            raise SystemExit(1)
        project_id = simpledialog.askstring("Google Credentials", "Enter your project_id:")
        if not project_id:
            root.destroy()
            raise SystemExit(1)
        creds_data = {
            "installed": {
                "client_id": client_id,
                "project_id": project_id,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_secret": client_secret,
                "redirect_uris": ["http://localhost"]
            }
        }
        with open(creds_path, 'w') as f:
            json.dump(creds_data, f, indent=2)
        root.destroy()
    return creds_data

def get_calendar_service():
    ensure_google_credentials()
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(f"Missing {CREDENTIALS_FILE}. Create OAuth credentials in Google Cloud Console and download as {CREDENTIALS_FILE}")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
    service = build('calendar', 'v3', credentials=creds)
    return service


def style_combobox_listbox(cb: ttk.Combobox, *, background="#2a2a2a", foreground="#ffffff", selectbackground="#505050", selectforeground="#ffffff"):
    try:
        popdown = cb.tk.call('ttk::combobox::PopdownWindow', cb)
        lb_path = popdown + '.f.l'
        lb = cb.nametowidget(lb_path)
        lb.configure(background=background, foreground=foreground,
                     selectbackground=selectbackground, selectforeground=selectforeground,
                     activestyle='none')

        # Make list start at top
        def _on_map(e, lb=lb):
            try:
                lb.yview_moveto(0)
                lb.see(0)
            except Exception:
                pass
        try:
            lb.bind('<Map>', _on_map)
        except Exception:
            pass
    except Exception:
        try:
            for child in cb.winfo_toplevel().winfo_children():
                try:
                    if isinstance(child, tk.Listbox):
                        child.configure(background=background, foreground=foreground,
                                        selectbackground=selectbackground, selectforeground=selectforeground,
                                        activestyle='none')
                        try:
                            child.bind('<Map>', lambda e, lb=child: (lb.yview_moveto(0), lb.see(0)))
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

def get_or_create_app_calendar(service):
    # If we already have a calendar id saved, check it first
    if os.path.exists(APP_CALENDAR_ID_FILE):
        with open(APP_CALENDAR_ID_FILE, 'r', encoding='utf-8') as f:
            cal_id = f.read().strip()
            try:
                service.calendars().get(calendarId=cal_id).execute()
                return cal_id
            except Exception:
                pass

    # Create a new calendar dedicated to this app
    calendar_body = {
        'summary': APP_CALENDAR_SUMMARY,
        'timeZone': 'America/New_York'
    }
    created = service.calendars().insert(body=calendar_body).execute()
    cal_id = created.get('id')
    with open(APP_CALENDAR_ID_FILE, 'w', encoding='utf-8') as f:
        f.write(cal_id)
    return cal_id

def slot_to_times(date_obj, time_slot):
    tz = '-04:00'  # (EDT)
    date_str = date_obj.strftime("%Y-%m-%d")
    if time_slot == "7-12":
        start = f"{date_str}T07:00:00{tz}"
        end = f"{date_str}T12:00:00{tz}"
    elif time_slot == "12-3pm":
        start = f"{date_str}T12:00:00{tz}"
        end = f"{date_str}T15:00:00{tz}"
    else:
        start = f"{date_str}T09:00:00{tz}"
        end = f"{date_str}T10:00:00{tz}"
    return start, end

def task_to_event_body(task, event_date, calendar_id):
    start_iso, end_iso = slot_to_times(event_date, task.time_slot)

    body = {
        'summary': task.description,
        'description': ", ".join(task.who) if task.who else "",
        'start': {'dateTime': start_iso, 'timeZone': 'America/New_York'},
        'end': {'dateTime': end_iso, 'timeZone': 'America/New_York'},
    }

    if task.recurrence in ("Daily", "Weekly", "Bi-Weekly", "Monthly", "Annually"):
        freq = None
        interval = None
        byday = None
        bymonthday = None
        until = None
        start_ref = task.start_date
        if isinstance(start_ref, datetime):
            start_ref = _coerce_to_weekday(start_ref)

        if task.recurrence == "Daily":
            freq = "DAILY"
            # Skip Weekends
            byday = "MO,TU,WE,TH,FR"
        elif task.recurrence == "Weekly":
            freq = "WEEKLY"
            # If user picked days, use them, otherwise default to weekdays
            weekday_map = {0: 'SU', 1: 'MO', 2: 'TU', 3: 'WE', 4: 'TH', 5: 'FR', 6: 'SA'}
            if task.weekly_days:
                byday = ",".join(weekday_map[d] for d in task.weekly_days if d not in (0, 6))  # skip Sat(6), Sun(0)
            else:
                byday = "MO,TU,WE,TH,FR"
        elif task.recurrence == "Bi-Weekly":
            freq = "WEEKLY"
            interval = 2
            byday = "MO,TU,WE,TH,FR"
        elif task.recurrence == "Monthly":
            freq = "MONTHLY"
            weekday_map = {0: 'MO', 1: 'TU', 2: 'WE', 3: 'TH', 4: 'FR'}
            weekday_code = weekday_map.get(start_ref.weekday() if isinstance(start_ref, datetime) else 0)
            if weekday_code:
                week_num = task.monthly_day or ((start_ref.day - 1) // 7 + 1) if isinstance(start_ref, datetime) else 1
                week_num = max(1, min(4, week_num))
                byday = f"{week_num}{weekday_code}"
            else:
                byday = "1MO"
        elif task.recurrence == "Annually":
            freq = "YEARLY"

        if task.end_date:
            until = task.end_date.strftime("%Y%m%dT235959Z")

        if freq:
            rrule = f"FREQ={freq}"
            if interval:
                rrule += f";INTERVAL={interval}"
            if byday:
                rrule += f";BYDAY={byday}"
            if bymonthday:
                rrule += f";BYMONTHDAY={bymonthday}"
            if until:
                rrule += f";UNTIL={until}"
            
            recurrence_list = [f"RRULE:{rrule}"]
            
            if getattr(task, 'skip_holidays', False):
                exdates = []
                check_date = start_ref if isinstance(start_ref, datetime) else datetime.now()
                end_check = task.end_date if task.end_date else (check_date + timedelta(days=365*2))  # Check up to 2 years ahead
                
                current = check_date
                max_holidays = 50
                holiday_count = 0
                
                while current <= end_check and holiday_count < max_holidays:
                    if _is_holiday(current) and current.weekday() < 5:  # Only weekdays
                        exdates.append(current.strftime("%Y%m%d"))
                        holiday_count += 1
                    current += timedelta(days=1)
                    # Limit search to reasonable range
                    if (current - check_date).days > 730:  # 2 years max
                        break
                
                if exdates:
                    # Group consecutive dates if possible, otherwise list individually
                    exdate_str = ",".join(exdates)
                    recurrence_list.append(f"EXDATE;VALUE=DATE:{exdate_str}")
            
            body['recurrence'] = recurrence_list

    return body

def create_or_update_event(service, calendar_id, task, event_date, existing_event_id=None):
    try:
        if isinstance(event_date, datetime):
            # Apply skip_holidays if enabled
            if getattr(task, 'skip_holidays', False):
                # Skip holidays and weekends
                event_date = _skip_holidays(event_date)
            else:
                # Just handle weekends and monthly adjustment
                if getattr(task, "recurrence", None) == "Monthly":
                    event_date = _start_of_week(event_date)
                event_date = _coerce_to_weekday(event_date)
            
            # Update task.start_date to match the adjusted event_date
            if isinstance(task.start_date, datetime) and task.start_date != event_date:
                task.start_date = event_date
        body = task_to_event_body(task, event_date, calendar_id)
        if existing_event_id:
            # update existing event for this specific calendar
            try:
                service.events().get(calendarId=calendar_id, eventId=existing_event_id).execute()
                updated = service.events().patch(calendarId=calendar_id, eventId=existing_event_id, body=body).execute()
                return updated.get('id')
            except Exception:
                # if the event id is invalid/missing, fall through to create
                pass
        created = service.events().insert(calendarId=calendar_id, body=body).execute()
        return created.get('id')
    except Exception as e:
        print("Google Calendar sync error:", e)
        try:
            messagebox.showerror("Google Calendar Error", f"Could not create/update the event:\n{e}")
        except Exception:
            pass
        return None

def delete_event(service, calendar_id, event_id):
    if not event_id:
        return
    try:
        service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
    except Exception as e:
        print("Failed to delete event:", e)

def get_google_calendar_colors(service):
    """Get available colors from Google Calendar API"""
    try:
        colors = service.colors().get().execute()
        return colors.get('calendar', {})
    except Exception as e:
        print(f"Failed to get Google Calendar colors: {e}")
        return {}

def create_profile_calendar(service, profile_name):
    """Create a Google Calendar for a profile"""
    calendar_body = {
        'summary': f"{profile_name} - Work Calendar",
        'timeZone': 'America/New_York'
    }
    try:
        created = service.calendars().insert(body=calendar_body).execute()
        calendar_id = created.get('id')
        print(f"Successfully created calendar for {profile_name}")
        return calendar_id
    except Exception as e:
        print(f"Failed to create calendar for {profile_name}:", e)
        return None

def delete_profile_calendar(service, calendar_id):
    """Delete a Google Calendar"""
    try:
        service.calendars().delete(calendarId=calendar_id).execute()
        print(f"Successfully deleted calendar {calendar_id}")
        return True
    except Exception as e:
        print(f"Failed to delete calendar {calendar_id}:", e)
        return False

def update_calendar_name(service, calendar_id, new_name):
    """Update the name of a Google Calendar"""
    try:
        calendar_body = {
            'summary': f"{new_name} - Work Calendar"
        }
        updated = service.calendars().patch(calendarId=calendar_id, body=calendar_body).execute()
        print(f"Successfully updated calendar name to {new_name}")
        return True
    except Exception as e:
        print(f"Failed to update calendar name: {e}")
        return False

def share_calendar_with_email(service, calendar_id, email):
    """Share a calendar with an email address (view only)"""
    try:
        rule = {
            'role': 'reader',
            'scope': {
                'type': 'user',
                'value': email
            }
        }
        service.acl().insert(calendarId=calendar_id, body=rule).execute()
        print(f"Successfully shared calendar with {email}")
        return True
    except Exception as e:
        print(f"Failed to share calendar with {email}: {e}")
        return False

def cleanup_tasks_for_deleted_profile(tasks, deleted_profile_name, service, profiles):
    """Clean up tasks when a profile is deleted.
    - If the deleted profile was the only assignee, delete the task (and its events)
    - If other assignees remain, remove only the deleted profile and keep the task
    - Only remove the deleted profile's event; keep and ensure events for remaining assignees
    """
    tasks_to_remove = []
    tasks_to_update = []

    for task in tasks:
        if deleted_profile_name in task.who:
            remaining_assignments = [name for name in task.who if name != deleted_profile_name]
            remaining_real_assignees = [name for name in remaining_assignments if name != "Unassigned"]

            if len(remaining_real_assignees) == 0:
                tasks_to_remove.append(task)
            else:
                task.who = remaining_real_assignees
                if not hasattr(task, 'event_ids') or not isinstance(task.event_ids, dict):
                    task.event_ids = {}
                # drop the deleted profile's event id
                task.event_ids.pop(deleted_profile_name, None)
                tasks_to_update.append(task)

    # Delete tasks that are assigned only to the deleted profile
    for task in tasks_to_remove:
        try:
            delete_task_from_calendars(service, task, profiles)
            print(f"Deleted task '{task.description}' (was assigned only to {deleted_profile_name})")
        except Exception as e:
            print(f"Error deleting task '{task.description}': {e}")

    # Ensure remaining assignees' events are preserved/created
    for task in tasks_to_update:
        try:
            for assignee in task.who:
                # Find calendar for this assignee
                cal_id = None
                if assignee == "Unassigned":
                    try:
                        cal_id = get_or_create_app_calendar(service)
                    except Exception:
                        cal_id = None
                else:
                    for p in profiles:
                        if isinstance(p, Profile) and p.name == assignee:
                            cal_id = p.calendar_id
                            break
                if not cal_id:
                    print(f"Skipping event sync for assignee '{assignee}' due to missing calendar id")
                    continue
                existing = task.event_ids.get(assignee)
                eid = create_or_update_event(service, cal_id, task, task.start_date, existing_event_id=existing)
                if eid:
                    task.event_ids[assignee] = eid
            # Maintain legacy single event_id for compatibility
            if task.event_ids:
                task.event_id = list(task.event_ids.values())[0]
            print(f"Updated task '{task.description}' (removed {deleted_profile_name} from assignment; preserved others)")
        except Exception as e:
            print(f"Error updating task '{task.description}': {e}")

    return tasks_to_remove

def _map_byday_to_weekday_indices(byday_list):
    code_to_idx = {'MO': 0, 'TU': 1, 'WE': 2, 'TH': 3, 'FR': 4, 'SA': 5, 'SU': 6}
    return [code_to_idx.get(code) for code in byday_list if code in code_to_idx]

def _parse_rrule_to_app_recurrence(rrule: str):
    # returns (recurrence, weekly_days, monthly_day, end_date)
    if not rrule:
        return ("None", [], None, None)
    parts = {}
    for seg in rrule.split(';'):
        if '=' in seg:
            k, v = seg.split('=', 1)
            parts[k.upper()] = v
    freq = parts.get('FREQ')
    interval = int(parts.get('INTERVAL', '1'))
    until = parts.get('UNTIL')
    end_date = None
    if until:
        try:
            end_date = datetime.strptime(until[:8], "%Y%m%d")
        except Exception:
            end_date = None
    if freq == 'DAILY':
        return ("Daily", [], None, end_date)
    if freq == 'WEEKLY':
        byday = parts.get('BYDAY', '')
        bydays = [d for d in byday.split(',') if d]
        weekday_indices = [i for i in _map_byday_to_weekday_indices(bydays) if i is not None]
        if interval == 2:
            return ("Bi-Weekly", weekday_indices, None, end_date)
        return ("Weekly", weekday_indices, None, end_date)
    if freq == 'MONTHLY':
        week_num = None
        byday = parts.get('BYDAY')
        if byday:
            entries = [d for d in byday.split(',') if d]
            for entry in entries:
                prefix = ''.join(ch for ch in entry if ch in '+-' or ch.isdigit())
                suffix = ''.join(ch for ch in entry if ch.isalpha())
                if prefix:
                    try:
                        pos = int(prefix)
                        if pos > 0:
                            week_num = max(1, min(4, pos))
                            break
                    except ValueError:
                        continue
        if not week_num:
            bymonthday = parts.get('BYMONTHDAY')
            if bymonthday:
                try:
                    day_num = int(bymonthday)
                    week_num = max(1, min(4, ((day_num - 1) // 7) + 1))
                except ValueError:
                    week_num = None
        return ("Monthly", [], week_num, end_date)
    if freq == 'YEARLY':
        return ("Annually", [], None, end_date)
    return ("None", [], None, end_date)

def _infer_time_slot(start_dt: datetime, end_dt: datetime):
    try:
        sh = start_dt.hour
        eh = end_dt.hour if end_dt else None
        if sh == 7 and (eh is None or eh <= 12):
            return "7-12"
        if sh in (12,) and (eh is None or eh <= 15):
            return "12-3pm"
        return "9-10"
    except Exception:
        return "9-10"

def _parse_google_event_to_task_fields(event: dict):
    desc = event.get('summary', '') or ''
    details = event.get('description', '') or ''
    who = [w.strip() for w in details.split(',') if w.strip()] if details else []
    # start/end
    def _parse_dt(d):
        if 'dateTime' in d:
            s = d['dateTime']
            try:
                return datetime.fromisoformat(s.replace('Z', '+00:00'))
            except Exception:
                return None
        elif 'date' in d:
            try:
                return datetime.strptime(d['date'], "%Y-%m-%d")
            except Exception:
                return None
        return None
    start_dt = _parse_dt(event.get('start', {})) or datetime.now()
    end_dt = _parse_dt(event.get('end', {})) or None
    time_slot = _infer_time_slot(start_dt, end_dt)
    # recurrence
    rrules = event.get('recurrence', []) or []
    rec, weekly_days, monthly_day, end_date = ("None", [], None, None)
    if rrules:
        for r in rrules:
            if r.upper().startswith('RRULE:'):
                rec, weekly_days, monthly_day, end_date = _parse_rrule_to_app_recurrence(r[6:])
                break
    # Adjust monthly events to start on Monday of that week
    if rec == "Monthly" and start_dt:
        start_dt = _start_of_week(start_dt)
    return {
        'description': desc,
        'who': who,
        'start_date': start_dt,
        'end_date': end_date,
        'recurrence': rec,
        'weekly_days': weekly_days,
        'monthly_day': monthly_day,
        'time_slot': time_slot,
    }

def sync_from_google_calendar(service, tasks, profiles):
    """Sync changes from Google Calendar back to the app.
    - Imports new events as tasks, mapped to recurrence/time slots/assignees (using recurring master when applicable)
    - Updates modified events
    - Removes only missing per-profile events from tasks; deletes task only if no assignees remain
    """
    print("Syncing from Google Calendar...")

    # Build calendar id -> profile name mapping
    cal_to_profile = {}
    for p in profiles:
        if isinstance(p, Profile) and p.calendar_id:
            cal_to_profile[p.calendar_id] = p.name
    try:
        main_cal_id = get_or_create_app_calendar(service)
    except Exception:
        main_cal_id = None

    # Track local mapping of event ids to tasks and which profile
    event_id_to_task = {}
    local_event_ids = set()
    for task in tasks:
        if getattr(task, 'event_ids', None):
            for prof, eid in task.event_ids.items():
                if eid:
                    event_id_to_task[eid] = (task, prof)
                    local_event_ids.add(eid)
        elif task.event_id:
            event_id_to_task[task.event_id] = (task, 'Unassigned')
            local_event_ids.add(task.event_id)

    # Collect events from all calendars
    cal_ids = set(cal_to_profile.keys())
    if main_cal_id:
        cal_ids.add(main_cal_id)

    # Time window
    now = datetime.now()
    time_min = (now - timedelta(days=180)).isoformat() + 'Z'
    time_max = (now + timedelta(days=365)).isoformat() + 'Z'

    # First pass: gather instances to discover recurringEventIds
    recurring_ids_needed = {}
    single_instances = {}  # eid -> (event, calendar_id) for non-recurring single events
    for calendar_id in cal_ids:
        try:
            page_token = None
            while True:
                inst = service.events().list(
                    calendarId=calendar_id,
                    timeMin=time_min,
                    timeMax=time_max,
                    singleEvents=True,
                    orderBy='startTime',
                    pageToken=page_token
                ).execute()
                for event in inst.get('items', []):
                    eid = event.get('id')
                    rid = event.get('recurringEventId')
                    if rid:
                        recurring_ids_needed.setdefault(calendar_id, set()).add(rid)
                    else:
                        if not event.get('recurrence'):
                            single_instances[eid] = (event, calendar_id)
                page_token = inst.get('nextPageToken')
                if not page_token:
                    break
        except Exception as e:
            print(f"Error fetching instances from calendar {calendar_id}: {e}")

    masters = {}
    for calendar_id in cal_ids:
        try:
            page_token = None
            while True:
                res = service.events().list(
                    calendarId=calendar_id,
                    timeMin=time_min,
                    timeMax=time_max,
                    singleEvents=False,
                    pageToken=page_token
                ).execute()
                for event in res.get('items', []):
                    if event.get('recurrence'):
                        masters[event.get('id')] = (event, calendar_id)
                page_token = res.get('nextPageToken')
                if not page_token:
                    break
        except Exception as e:
            print(f"Error fetching masters from calendar {calendar_id}: {e}")

    for calendar_id, rid_set in recurring_ids_needed.items():
        for rid in rid_set:
            if rid not in masters:
                try:
                    ev = service.events().get(calendarId=calendar_id, eventId=rid).execute()
                    if ev and ev.get('recurrence'):
                        masters[rid] = (ev, calendar_id)
                except Exception as e:
                    print(f"Failed to fetch master {rid} from {calendar_id}: {e}")

    google_event_index = {}
    for mid, tup in masters.items():
        google_event_index[mid] = {'event': tup[0], 'calendar_id': tup[1]}
    for eid, tup in single_instances.items():
        if eid not in google_event_index:
            google_event_index[eid] = {'event': tup[0], 'calendar_id': tup[1]}

    new_event_ids = set(google_event_index.keys()) - set(event_id_to_task.keys())

    # Import new events
    for eid in new_event_ids:
        data = google_event_index[eid]
        event = data['event']
        calendar_id = data['calendar_id']
        fields = _parse_google_event_to_task_fields(event)
        # Determine who based on calendar
        if calendar_id == main_cal_id:
            assignees = fields['who'] or ["Unassigned"]
        else:
            prof_name = cal_to_profile.get(calendar_id)
            assignees = fields['who'] or ([prof_name] if prof_name else ["Unassigned"])
        new_task = Task(
            fields['description'],
            fields['recurrence'],
            fields['start_date'],
            fields['end_date'],
            weekly_days=fields['weekly_days'],
            monthly_day=fields['monthly_day'],
            who=assignees,
            time_slot=fields['time_slot']
        )
        new_task.event_ids = {}
        for who_name in assignees:
            if fields['recurrence'] != "None":
                target_match = (calendar_id == main_cal_id and who_name == "Unassigned") or \
                               (calendar_id in cal_to_profile and cal_to_profile[calendar_id] == who_name)
                if target_match:
                    new_task.event_ids[who_name] = eid
            else:
                target_match = (calendar_id == main_cal_id and who_name == "Unassigned") or \
                               (calendar_id in cal_to_profile and cal_to_profile[calendar_id] == who_name)
                if target_match:
                    new_task.event_ids[who_name] = eid
        if not new_task.event_ids and fields['recurrence'] == "None":
            new_task.event_id = eid
        tasks.append(new_task)
        print(f"Imported new {'recurring' if fields['recurrence']!='None' else 'single'} event as task: {new_task.description}")

    # Handle per-profile deletions using google_event_index only (masters/singles)
    google_known_ids = set(google_event_index.keys())
    missing_event_ids = set(event_id_to_task.keys()) - google_known_ids
    for missing_eid in list(missing_event_ids):
        task_info = event_id_to_task.get(missing_eid)
        if not task_info:
            continue
        task, prof = task_info
        # Remove only this profile's event
        if getattr(task, 'event_ids', None) and prof in task.event_ids and task.event_ids.get(prof) == missing_eid:
            task.event_ids.pop(prof, None)
            if prof in task.who:
                task.who = [w for w in task.who if w != prof]
        elif task.event_id == missing_eid:
            task.event_id = None
            task.who = [w for w in task.who if w != "Unassigned"]
        # If no one remains assigned, remove the task
        if not task.who:
            try:
                tasks.remove(task)
                print(f"Removed task '{task.description}' (no remaining assignees after deletion)")
            except ValueError:
                pass

    # Handle modified events (masters/singles)
    modified_count = 0
    for eid, data in google_event_index.items():
        if eid not in event_id_to_task:
            continue
        task, prof = event_id_to_task[eid]
        event = data['event']
        fields = _parse_google_event_to_task_fields(event)
        changed = False
        if fields['description'] and fields['description'] != task.description:
            task.description = fields['description']
            changed = True
        # Update dates only if start date changed (keep recurrence model)
        if fields['start_date'] and fields['start_date'].date() != task.start_date.date():
            task.start_date = fields['start_date']
            changed = True
        if fields['end_date'] != task.end_date:
            task.end_date = fields['end_date']
            changed = True
        # Update recurrence mapping if RRULE changed
        if fields['recurrence'] != task.recurrence or fields['weekly_days'] != task.weekly_days or fields['monthly_day'] != task.monthly_day:
            task.recurrence = fields['recurrence']
            task.weekly_days = fields['weekly_days']
            task.monthly_day = fields['monthly_day']
            changed = True
        # Update time slot if needed
        if fields['time_slot'] != task.time_slot:
            task.time_slot = fields['time_slot']
            changed = True
        if changed:
            modified_count += 1

    print(f"Sync complete. Imported {len(new_event_ids)} new, updated {modified_count}, removed {len(missing_event_ids)} per-profile events.")
    return (len(missing_event_ids), modified_count)

def get_calendar_ids_for_task(task, profiles, service):
    """Get all appropriate calendar IDs for a task based on assigned profiles"""
    calendar_ids = {}
    
    if not task.who or task.who == ["Unassigned"]:
        # Use the main app calendar for unassigned tasks
        main_cal_id = get_or_create_app_calendar(service)
        calendar_ids["Unassigned"] = main_cal_id
        return calendar_ids
    
    # For tasks assigned to profiles, get all their calendars
    for profile_name in task.who:
        if profile_name == "Unassigned":
            continue
            
        for profile in profiles:
            if isinstance(profile, Profile) and profile.name == profile_name:
                if profile.calendar_id:
                    calendar_ids[profile_name] = profile.calendar_id
                else:
                    # Profile exists but no calendar - this shouldn't happen with auto-creation
                    print(f"Warning: Profile {profile.name} has no calendar ID")
                break
    
    # If no profile calendars found, fallback to main app calendar
    if not calendar_ids:
        main_cal_id = get_or_create_app_calendar(service)
        calendar_ids["Unassigned"] = main_cal_id
    
    return calendar_ids

def sync_task_to_calendars(service, task, profiles):
    """Sync a task to all appropriate calendars, preserving existing per-profile events."""
    calendar_ids = get_calendar_ids_for_task(task, profiles, service)
    existing_map = dict(task.event_ids) if getattr(task, 'event_ids', None) else {}
    new_event_ids = dict(existing_map)

    for profile_name, calendar_id in calendar_ids.items():
        try:
            existing = existing_map.get(profile_name)
            event_id = create_or_update_event(service, calendar_id, task, task.start_date, existing_event_id=existing)
            if event_id:
                new_event_ids[profile_name] = event_id
                print(f"Created/updated event for {profile_name} in calendar {calendar_id}")
            else:
                print(f"Failed to create/update event for {profile_name}")
        except Exception as e:
            print(f"Error syncing task to {profile_name}'s calendar: {e}")

    return new_event_ids

def delete_task_from_calendars(service, task, profiles):
    """Delete a task from all calendars it was synced to"""
    if not task.event_ids:
        # Legacy: try to delete using single event_id
        if task.event_id:
            calendar_ids = get_calendar_ids_for_task(task, profiles, service)
            for profile_name, calendar_id in calendar_ids.items():
                try:
                    delete_event(service, calendar_id, task.event_id)
                    print(f"Deleted legacy event for {profile_name}")
                except Exception as e:
                    print(f"Error deleting legacy event for {profile_name}: {e}")
        return
    
    # Delete from all calendars using stored event_ids
    for profile_name, event_id in task.event_ids.items():
        if not event_id:
            continue
            
        # Find the calendar ID for this profile
        calendar_id = None
        if profile_name == "Unassigned":
            calendar_id = get_or_create_app_calendar(service)
        else:
            for profile in profiles:
                if isinstance(profile, Profile) and profile.name == profile_name:
                    calendar_id = profile.calendar_id
                    break
        
        if calendar_id:
            try:
                delete_event(service, calendar_id, event_id)
                print(f"Deleted event for {profile_name}")
            except Exception as e:
                print(f"Error deleting event for {profile_name}: {e}")

# Legacy function for backward compatibility
def get_calendar_id_for_task(task, profiles, service):
    """Get the appropriate calendar ID for a task based on assigned profiles (legacy)"""
    calendar_ids = get_calendar_ids_for_task(task, profiles, service)
    # Return the first calendar ID for backward compatibility
    return list(calendar_ids.values())[0] if calendar_ids else get_or_create_app_calendar(service)


class Profile:
    def __init__(self, name, email=None, calendar_id=None, is_admin=False):
        self.name = name
        self.email = email  # Optional email for sharing
        self.calendar_id = calendar_id  # Google Calendar ID
        self.is_admin = is_admin  # Administrator can see all tasks

class Task:
    def __init__(self, description, recurrence, start_date, end_date,
                 weekly_days=None, monthly_day=None, who=None, time_slot=None, event_id=None, event_ids=None, skip_holidays=False):
        self.description = description
        self.recurrence = recurrence
        if isinstance(start_date, datetime):
            if recurrence == "Monthly":
                start_date = _start_of_week(start_date)
            if start_date.weekday() >= 5:
                start_date = _coerce_to_weekday(start_date)
            if skip_holidays:
                # Skip holidays and weekends, moving to next valid workday
                start_date = _skip_holidays(start_date)
        self.start_date = start_date
        if isinstance(end_date, datetime):
            if end_date.weekday() >= 5:
                end_date = _coerce_to_weekday(end_date)
            if skip_holidays:
                # Skip holidays and weekends, moving to next valid workday
                end_date = _skip_holidays(end_date)
        self.end_date = end_date  # None means indefinite
        self.weekly_days = weekly_days or []
        self.monthly_day = monthly_day
        self.who = who or []
        self.time_slot = time_slot or ""  # "7-12" or "12-3pm"
        self.event_id = event_id  # store google event id so edits/deletes sync (legacy)
        self.event_ids = event_ids or {}  # store multiple event ids: {profile_name: event_id}
        self.skip_holidays = skip_holidays

    def occurs_on(self, date: datetime):
        d = date.date()
        if d < self.start_date.date():
            return False
        if self.end_date and d > self.end_date.date():
            return False
        
        # Skip holidays if enabled
        if self.skip_holidays and _is_holiday(date):
            return False

        if self.recurrence == "None":
            return d == self.start_date.date()
        if self.recurrence == "Daily":
            # Skip weekends
            return (self.start_date.date() <= d
                    and (not self.end_date or d <= self.end_date.date())
                    and date.weekday() < 5)  # 0=Mon, 1=Tue, ..., 4=Fri
        if self.recurrence == "Weekly":
            return date.weekday() in self.weekly_days
        if self.recurrence == "Bi-Weekly":
            delta_days = (d - self.start_date.date()).days
            weeks = delta_days // 7
            return (weeks % 2 == 0) and (date.weekday() in self.weekly_days)
        if self.recurrence == "Monthly":
            if date.weekday() >= 5:
                return False
            # monthly_day stores week number (1-4). For months with 5th occurrence, treat as 4th week.
            week_num = (date.day - 1) // 7 + 1
            target_week = self.monthly_day or ((self.start_date.day - 1) // 7 + 1)
            if target_week == 4 and week_num >= 4:
                week_match = True
            else:
                week_match = week_num == target_week
            return week_match and date.weekday() == self.start_date.weekday()
        if self.recurrence == "Annually":
            return (
                        date.month == self.start_date.month and date.day == self.start_date.day and d >= self.start_date.date())
        return False

def _prune_removed_assignee_events(service, task, profiles):
    """Delete Google Calendar events for assignees removed from the task."""
    if not getattr(task, 'event_ids', None):
        return
    current_assignees = set(task.who or [])
    to_remove = []
    for assignee, event_id in list(task.event_ids.items()):
        if assignee not in current_assignees and event_id:
            to_remove.append((assignee, event_id))
    for assignee, event_id in to_remove:
        cal_id = None
        if assignee == "Unassigned":
            try:
                cal_id = get_or_create_app_calendar(service)
            except Exception:
                cal_id = None
        else:
            for p in profiles:
                if isinstance(p, Profile) and p.name == assignee:
                    cal_id = p.calendar_id
                    break
        if cal_id:
            try:
                delete_event(service, cal_id, event_id)
                print(f"Pruned event for removed assignee '{assignee}' on task '{task.description}'")
            except Exception as e:
                print(f"Failed to prune event for '{assignee}': {e}")
        # drop from mapping regardless
        task.event_ids.pop(assignee, None)


class DateRangePicker(tk.Toplevel):
    """
    Date range picker that highlights start/end and range.
    If allow_end is False, a single click selects the start (used for 'indefinitely' mode).
    """

    def __init__(self, parent, start_var: tk.StringVar, end_var: tk.StringVar, allow_end=True, reset_start=False):
        super().__init__(parent)
        self.title("Select Date Range")
        self.start_var = start_var
        self.end_var = end_var
        self.allow_end = allow_end

        # *** Important: initialize before draw_calendar ***
        self.selected_start = None
        self.selected_end = None
        
        # If start is already set, use it unless caller asked to reset (so user can re-pick start)
        if not reset_start:
            try:
                start_str = start_var.get()
                if start_str:
                    self.selected_start = datetime.strptime(start_str, "%Y-%m-%d")
            except Exception:
                pass

        self.transient(parent)
        self.grab_set()

        # Show month of start date if set, otherwise current month
        if self.selected_start:
            now = self.selected_start
        else:
            now = datetime.now()
        self.year = now.year
        self.month = now.month

        nav = ttk.Frame(self)
        nav.pack(padx=8, pady=6, fill="x")
        ttk.Button(nav, text="<", width=3, command=self.prev_month).pack(side="left")
        ttk.Button(nav, text=">", width=3, command=self.next_month).pack(side="right")

        self.calendar_frame = ttk.Frame(self, padding=6)
        self.calendar_frame.pack()
        self.draw_calendar()

    def draw_calendar(self):
        for w in self.calendar_frame.winfo_children():
            w.destroy()

        cal = calendar.Calendar(firstweekday=6)
        ttk.Label(self.calendar_frame, text=f"{calendar.month_name[self.month]} {self.year}",
                  font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=7)

        days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
        for i, d in enumerate(days):
            ttk.Label(self.calendar_frame, text=d).grid(row=1, column=i)

        month_days = cal.monthdayscalendar(self.year, self.month)
        for r, week in enumerate(month_days, start=2):
            for c, day in enumerate(week):
                if day == 0:
                    ttk.Label(self.calendar_frame, text="", width=4).grid(row=r, column=c, padx=2, pady=2)
                    continue
                date_obj = datetime(self.year, self.month, day)
                is_weekend = date_obj.weekday() >= 5
                state = "disabled" if is_weekend else "normal"
                btn = tk.Button(
                    self.calendar_frame,
                    text=str(day),
                    width=4,
                    command=(lambda d=date_obj: self.pick_date(d)) if not is_weekend else None,
                    state=state,
                    disabledforeground="#a0a0a0"
                )
                if is_weekend:
                    btn.configure(bg="#e6e6e6", relief="flat")

                # highlight selected range / start / end
                if self.selected_start and self.selected_end and self.selected_start.date() <= date_obj.date() <= self.selected_end.date():
                    btn.configure(bg="lightblue")
                elif self.selected_start and date_obj.date() == self.selected_start.date():
                    btn.configure(bg="lightgreen")
                elif self.selected_end and date_obj.date() == self.selected_end.date():
                    btn.configure(bg="lightcoral")

                btn.grid(row=r, column=c, padx=2, pady=2)

    def pick_date(self, date_obj: datetime):
        if date_obj.weekday() >= 5:
            return
        if not self.selected_start:
            # first click: set start
            self.selected_start = date_obj
            if not self.allow_end:
                # single-selection mode: set start immediately and finish (no end date needed)
                self.start_var.set(self.selected_start.strftime("%Y-%m-%d"))
                # Clear end date for non-recurring tasks
                self.end_var.set("")
                self.destroy()
                return
            else:
                self.selected_end = None
        elif not self.selected_end:
            # second click: set end (swap if earlier)
            if date_obj < self.selected_start:
                self.selected_end, self.selected_start = self.selected_start, date_obj
            else:
                self.selected_end = date_obj
            # save and close
            self.start_var.set(self.selected_start.strftime("%Y-%m-%d"))
            self.end_var.set(self.selected_end.strftime("%Y-%m-%d"))
            self.destroy()
            return
        else:
            # reset start to new selection (start over)
            self.selected_start = date_obj
            self.selected_end = None

        self.draw_calendar()

    def prev_month(self):
        self.month -= 1
        if self.month == 0:
            self.month = 12
            self.year -= 1
        self.draw_calendar()

    def next_month(self):
        self.month += 1
        if self.month == 13:
            self.month = 1
            self.year += 1
        self.draw_calendar()


class CalendarApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Work Calendar")

        # top state
        self.year_var = tk.IntVar()
        self.month_var = tk.IntVar()
        now = datetime.now()
        self.year_var.set(now.year)
        # Set month state to current month so app opens on current month
        self.month_var.set(now.month)

        self.tasks = []
        self.profiles = []  # Will store profile objects with name, color, calendar_id
        self.selected_day = None
        self.selected_date = None
        self.current_edit_task = None

        # track where we came from
        self.came_from_all_tasks = False

        # main frames
        self.main_frame = ttk.Frame(root, padding=6)
        self.main_frame.pack(fill="both", expand=True)

        # left side contains topbar + calendar/day/editor
        self.left_frame = ttk.Frame(self.main_frame)
        self.left_frame.pack(side="left", fill="both", expand=True)

        # full-page all tasks frame (shown instead of left_frame)
        self.all_tasks_frame = ttk.Frame(self.main_frame, padding=10)

        # profiles full page frame (left_frame area)
        self.profiles_page = ttk.Frame(self.left_frame, padding=10)

        # Topbar (inside left_frame)
        self.topbar = ttk.Frame(self.left_frame)
        self.topbar.pack(fill="x")
        self.create_topbar()

        # pages inside left_frame
        self.calendar_frame = ttk.Frame(self.left_frame, padding=10)
        self.day_view_frame = ttk.Frame(self.left_frame, padding=10)
        self.task_create_frame = ttk.Frame(self.left_frame, padding=10)

        # recurrence var used by task editor
        self.recurrence_var = tk.StringVar(value="None")

        # Date stringvars for task editor
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()

        # show calendar and load data
        self.show_calendar()
        self.load_data()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_topbar(self):
        # Topbar no longer contains month/year controls - title now provides dropdowns.

        ttk.Button(self.topbar, text="All Tasks", command=self.show_all_tasks).pack(side="left", padx=8)
        ttk.Button(self.topbar, text="Sync from Google", command=self.manual_sync).pack(side="left", padx=8)
        ttk.Button(self.topbar, text="Profiles", command=self.show_profiles_page).pack(side="right", padx=6)
        ttk.Button(self.topbar, text="Export PDF", command=self.export_pdf).pack(side="right", padx=6)

    # ---------- Profiles full page ----------
    def show_profiles_page(self):
        # Hide all other main frames so only profiles page is visible
        for frame in [self.calendar_frame, self.day_view_frame, self.task_create_frame, self.all_tasks_frame]:
            try:
                frame.pack_forget()
            except Exception:
                pass
        if not self.left_frame.winfo_ismapped():
            self.left_frame.pack(side="left", fill="both", expand=True)
        self.topbar.pack_forget()

        for w in self.profiles_page.winfo_children():
            w.destroy()
        self.profiles_page.pack(fill="both", expand=True)

        ttk.Label(self.profiles_page, text="Profiles", font=("Arial", 14, "bold")).pack(pady=10)

        # Profiles list with Treeview for better display
        columns = ("Name", "Email", "Admin", "Calendar", "Actions")
        # Restore original Treeview setup: no custom style, no heading color change, and ensure profiles display
        self.profiles_tree = ttk.Treeview(self.profiles_page, columns=columns, show="headings", height=8)
        self.profiles_tree.heading("Name", text="Name")
        self.profiles_tree.heading("Email", text="Email")
        self.profiles_tree.heading("Admin", text="Admin")
        self.profiles_tree.heading("Calendar", text="Calendar Status")
        self.profiles_tree.heading("Actions", text="Actions")
        self.profiles_tree.column("Name", width=120)
        self.profiles_tree.column("Email", width=180)
        self.profiles_tree.column("Admin", width=60)
        self.profiles_tree.column("Calendar", width=100)
        self.profiles_tree.column("Actions", width=100)
        self.profiles_tree.pack(pady=5, fill="x")
        self.profiles_tree.bind("<Button-1>", self.on_profile_click)
        self._refresh_profiles_tree()

        # Add profile section
        add_frame = ttk.LabelFrame(self.profiles_page, text="Add New Profile", padding=10)
        add_frame.pack(pady=10, fill="x")
        
        name_frame = ttk.Frame(add_frame)
        name_frame.pack(fill="x", pady=2)
        ttk.Label(name_frame, text="Name:").pack(side="left")
        self.new_profile_entry = ttk.Entry(name_frame)
        self.new_profile_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        email_frame = ttk.Frame(add_frame)
        email_frame.pack(fill="x", pady=2)
        ttk.Label(email_frame, text="Email (optional):").pack(side="left")
        self.new_profile_email_entry = ttk.Entry(email_frame)
        self.new_profile_email_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        admin_frame = ttk.Frame(add_frame)
        admin_frame.pack(fill="x", pady=2)
        self.new_profile_admin_var = tk.BooleanVar()
        ttk.Checkbutton(admin_frame, text="Administrator (can see all tasks)", variable=self.new_profile_admin_var).pack(side="left")
        
        ttk.Button(add_frame, text="Add Profile", command=self.add_profile).pack(pady=5)

        # Action buttons
        button_frame = ttk.Frame(self.profiles_page)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Delete Selected", command=self.delete_profiles).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Reset Credentials", command=self.reset_google_credentials).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Back", command=self.back_from_profiles).pack(side="right", padx=5)

    def on_profile_click(self, event):
        """Handle clicks on profile treeview, specifically the Edit button"""
        item = self.profiles_tree.identify_row(event.y)
        column = self.profiles_tree.identify_column(event.x)
        
        if item and column == "#5":  # Actions column
            # Get the profile name from the first column
            profile_name = self.profiles_tree.item(item, "values")[0]
            self.edit_profile(profile_name)
    
    def back_from_profiles(self):
        self.profiles_page.pack_forget()
        self.topbar.pack(fill="x")
        self.show_calendar()

    # ---------- Calendar ----------
    def show_calendar(self):
        # ensure all_tasks frame is hidden and left side visible
        self.all_tasks_frame.pack_forget()
        if not self.left_frame.winfo_ismapped():
            self.left_frame.pack(side="left", fill="both", expand=True)
        self.topbar.pack(fill="x")

        self.day_view_frame.pack_forget()
        self.task_create_frame.pack_forget()

        for w in self.calendar_frame.winfo_children():
            w.destroy()
        self.calendar_frame.pack(fill="both", expand=True)

        year, month = self.year_var.get(), self.month_var.get()
        cal = calendar.Calendar(firstweekday=6)

        title_frame = ttk.Frame(self.calendar_frame)
        title_frame.grid(row=0, column=0, columnspan=7, sticky="ew", pady=1)
        for i in range(7):
            self.calendar_frame.grid_columnconfigure(i, weight=1)
        title_frame.grid_columnconfigure(0, minsize=48)
        title_frame.grid_columnconfigure(1, weight=1)
        title_frame.grid_columnconfigure(2, minsize=48)

        ttk.Button(title_frame, text="◀", width=3, command=self.prev_month).grid(row=0, column=0, sticky="w")
        # Replace static title with Month + Year dropdowns
        months = list(calendar.month_name)[1:]
        years = list(range(1900, 2101))
        title_select = ttk.Frame(title_frame)
        title_select.grid(row=0, column=1)
        # Month combobox (shows full month name)
        mbox = ttk.Combobox(title_select, values=months, width=12, state="readonly")
        # Show the current month as the selected value, keep list ordered from January
        try:
            mbox.set(calendar.month_name[month])
        except Exception:
            mbox.set(months[0])
        mbox.pack(side="left", padx=(0, 6))
        try:
            style_combobox_listbox(mbox, background="#2a2a2a", foreground="#ffffff", selectbackground="#505050", selectforeground="#ffffff")
        except Exception:
            pass
        def _on_title_month(e):
            try:
                sel = mbox.get()
                idx = months.index(sel) + 1
                self.month_var.set(idx)
                self.show_calendar()
            except Exception:
                pass
        mbox.bind("<<ComboboxSelected>>", _on_title_month)

        # Year combobox
        ybox = ttk.Combobox(title_select, values=years, width=6, state="readonly")
        try:
            ybox.set(str(year))
        except Exception:
            ybox.set(str(years[0]))
        ybox.pack(side="left")
        try:
            style_combobox_listbox(ybox, background="#2a2a2a", foreground="#ffffff", selectbackground="#505050", selectforeground="#ffffff")
        except Exception:
            pass
        def _on_title_year(e):
            try:
                self.year_var.set(int(ybox.get()))
                self.show_calendar()
            except Exception:
                pass
        ybox.bind("<<ComboboxSelected>>", _on_title_year)

        ttk.Button(title_frame, text="▶", width=3, command=self.next_month).grid(row=0, column=2, sticky="e")

        days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
        for i, d in enumerate(days):
            ttk.Label(self.calendar_frame, text=d, font=("Arial", 10, "bold")).grid(row=1, column=i, padx=4)

        month_days = cal.monthdayscalendar(year, month)
        for r, week in enumerate(month_days, start=2):
            for c, day in enumerate(week):
                if day == 0:
                    ttk.Label(self.calendar_frame, text="").grid(row=r, column=c, padx=2, pady=2)
                    continue
                date_obj = datetime(year, month, day)
                padx = 3 if c in (0, 6) else 2
                pady = 3 if c in (0, 6) else 2
                if c in (0, 6):
                    # Greyed-out weekend cells: slightly darker background and more muted text
                    lbl = tk.Label(
                        self.calendar_frame,
                        text=str(day),
                        width=4,
                        relief="flat",
                        bd=0,
                        highlightthickness=0,
                        bg="#2e2e2e",
                        fg="#bfbfbf",
                    )
                    lbl.grid(row=r, column=c, padx=padx, pady=pady)
                else:
                    btn = tk.Button(self.calendar_frame, text=str(day), width=4, command=lambda d=day: self.show_day_view(d))
                    btn.grid(row=r, column=c, padx=padx, pady=pady)

    # ---------- Day view ----------
    def show_day_view(self, day):
        self.selected_day = day

        # ensure left frame visible and all_tasks hidden
        self.all_tasks_frame.pack_forget()
        if not self.left_frame.winfo_ismapped():
            self.left_frame.pack(side="left", fill="both", expand=True)
        self.topbar.pack(fill="x")

        self.calendar_frame.pack_forget()
        self.task_create_frame.pack_forget()

        for w in self.day_view_frame.winfo_children():
            w.destroy()
        self.day_view_frame.pack(fill="both", expand=True)

        year, month = self.year_var.get(), self.month_var.get()
        date_obj = datetime(year, month, day)
        self.selected_date = date_obj
        date_str = date_obj.strftime("%A, %B %d, %Y")

        ttk.Label(self.day_view_frame, text=f"Tasks for {date_str}", font=("Arial", 14)).pack(pady=10)

        tasks_today = [t for t in self.tasks if t.occurs_on(date_obj)]
        if not tasks_today:
            ttk.Label(self.day_view_frame, text="No tasks.").pack(pady=5)
        else:
            for t in tasks_today:
                assigned = "Unassigned"
                if t.who and t.who != ["Unassigned"]:
                    assigned = ", ".join(t.who) if len(t.who) <= 2 else f"{len(t.who)} people"

                row = ttk.Frame(self.day_view_frame)
                row.pack(fill="x", pady=2)
                # left-aligned text with time slot
                ttk.Label(row, text=f"- [{t.time_slot}] {t.description}  | {assigned}", anchor="w", justify="left").pack(fill="x", side="left")
                ttk.Button(row, text="Edit", command=lambda task=t: self.edit_task(task)).pack(side="right", padx=2)
                ttk.Button(row, text="Delete", command=lambda task=t: self.delete_task(task)).pack(side="right", padx=2)

        ttk.Button(self.day_view_frame, text="Add Task", command=lambda: self.show_task_create(editing=False)).pack(pady=10)
        ttk.Button(self.day_view_frame, text="Back", command=self.show_calendar).pack(pady=5)

    # ---------- Create / edit task ----------
    def show_task_create(self, editing=False):
        # ensure left_frame visible (editor is inside left frame)
        self.all_tasks_frame.pack_forget()
        if not self.left_frame.winfo_ismapped():
            self.left_frame.pack(side="left", fill="both", expand=True)
        self.topbar.pack(fill="x")

        self.calendar_frame.pack_forget()
        self.day_view_frame.pack_forget()
        for w in self.task_create_frame.winfo_children():
            w.destroy()
        self.task_create_frame.pack(fill="both", expand=True)

        ttk.Label(self.task_create_frame, text="Edit Task" if editing else "New Task", font=("Arial", 14)).pack(pady=10)

        ttk.Label(self.task_create_frame, text="Description:").pack(anchor="w")
        self.task_desc_entry = ttk.Entry(self.task_create_frame, width=40)
        self.task_desc_entry.pack(pady=5)

        ttk.Label(self.task_create_frame, text="Assign To:").pack(anchor="w")
        # Use checkboxes instead of listbox to avoid selection issues
        self.profile_checkboxes = []
        profile_frame = ttk.Frame(self.task_create_frame)
        profile_frame.pack(pady=5, fill="x")
        self.profile_vars = {}  # profile_name -> tk.BooleanVar
        for p in self.profiles:
            if isinstance(p, Profile):
                var = tk.BooleanVar()
                self.profile_vars[p.name] = var
                cb = ttk.Checkbutton(profile_frame, text=p.name, variable=var)
                cb.pack(side="left", padx=10)
                self.profile_checkboxes.append(cb)

        ttk.Label(self.task_create_frame, text="Time slot:").pack(anchor="w")
        self.time_slot_var = tk.StringVar()
        self.time_slot_menu = ttk.Combobox(self.task_create_frame, values=["7-12", "12-3pm"], textvariable=self.time_slot_var, state="readonly")
        self.time_slot_menu.pack(pady=5)
        try:
            style_combobox_listbox(self.time_slot_menu, background="#2a2a2a", foreground="#ffffff", selectbackground="#505050", selectforeground="#ffffff")
        except Exception:
            pass

        ttk.Label(self.task_create_frame, text="Recurrence:").pack(anchor="w", pady=(6, 0))
        recs = ["None", "Daily", "Weekly", "Bi-Weekly", "Monthly", "Annually"]
        self.recurrence_menu = ttk.Combobox(self.task_create_frame, values=recs, textvariable=self.recurrence_var,state="readonly")
        self.recurrence_menu.pack(pady=5)
        try:
            style_combobox_listbox(self.recurrence_menu, background="#2a2a2a", foreground="#ffffff", selectbackground="#505050", selectforeground="#ffffff")
        except Exception:
            pass
        self.recurrence_menu.bind("<<ComboboxSelected>>", lambda e: (self.update_recurrence_options(e), self._update_end_date_visibility()))

        # Date pickers / display
        ttk.Label(self.task_create_frame, text="Date Range:").pack(anchor="w")
        # These StringVars are per-editor; we'll use the instance ones defined on app so save/load matches
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()
        # Auto-fill start date if a day was selected
        selected_date = getattr(self, "selected_date", None)
        if selected_date:
            selected_date = _coerce_to_weekday(selected_date)
            self.selected_date = selected_date
            self.start_date_var.set(selected_date.strftime("%Y-%m-%d"))
        else:
            default_start = _coerce_to_weekday(datetime.now())
            self.start_date_var.set(default_start.strftime("%Y-%m-%d"))
            self.selected_date = default_start
        self.end_date_var.set("")
        # Button to open date picker. Open with reset_start=True so the
        # user can re-pick the start date using the same button (no separate button).
        def open_date_picker():
            rec = self.recurrence_var.get()
            allow_end = (rec != "None") and not self.repeat_indef_var.get()
            # reset_start=True ensures the picker won't auto-select the existing start
            # and the first click will set the start (so single button covers both flows).
            DateRangePicker(self.root, self.start_date_var, self.end_date_var, allow_end=allow_end, reset_start=True)

        ttk.Button(self.task_create_frame, text="Pick Start/End Dates", command=open_date_picker).pack(pady=6)
        ttk.Label(self.task_create_frame, text="Start:").pack(anchor="w")
        ttk.Label(self.task_create_frame, textvariable=self.start_date_var).pack(anchor="w")
        # Only show End label if recurrence is not "None"
        self.end_label = ttk.Label(self.task_create_frame, text="End:")
        self.end_value_label = ttk.Label(self.task_create_frame, textvariable=self.end_date_var)
        self._update_end_date_visibility()

        # Repeat Indefinitely checkbox ABOVE date pickers
        self.repeat_indef_var = tk.BooleanVar(value=False)
        repeat_chk = ttk.Checkbutton(self.task_create_frame, text="Repeat Indefinitely", variable=self.repeat_indef_var,command=self._on_repeat_toggle)
        repeat_chk.pack(anchor="w", pady=(6, 4))
        
        # Skip Holidays checkbox
        self.skip_holidays_var = tk.BooleanVar(value=False)
        skip_holidays_chk = ttk.Checkbutton(self.task_create_frame, text="Skip Holidays", variable=self.skip_holidays_var)
        skip_holidays_chk.pack(anchor="w", pady=(4, 0))

        self.recur_options_frame = ttk.Frame(self.task_create_frame)
        self.recur_options_frame.pack(pady=4, fill="x")

        self.indef_var = tk.BooleanVar(value=False)  # kept for backward compatibility in code paths
        # Save / Cancel buttons
        btns = ttk.Frame(self.task_create_frame)
        btns.pack(pady=12)
        ttk.Button(btns, text="Save", command=lambda: self.save_task(editing)).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="Cancel", command=self.show_day_view_cancel).grid(row=0, column=1, padx=6)

        # prefill when editing
        if editing and self.current_edit_task:
            t = self.current_edit_task
            self.task_desc_entry.insert(0, t.description)
            self.time_slot_var.set(t.time_slot)
            if t.who:
                for p in t.who:
                    if p in self.profile_vars:
                        self.profile_vars[p].set(True)
            # recurrence/options
            self.recurrence_var.set(t.recurrence)
            self.recurrence_menu.set(t.recurrence)
            self.update_recurrence_options()
            # dates
            self.start_date_var.set(t.start_date.strftime("%Y-%m-%d"))
            if t.recurrence == "None":
                # Non-recurring tasks don't need end date
                self.end_date_var.set("")
                self.repeat_indef_var.set(False)
            elif t.end_date:
                self.end_date_var.set(t.end_date.strftime("%Y-%m-%d"))
                self.repeat_indef_var.set(False)
            else:
                self.end_date_var.set("Indefinite")
                self.repeat_indef_var.set(True)
            # skip holidays
            if hasattr(t, 'skip_holidays'):
                self.skip_holidays_var.set(t.skip_holidays)
            else:
                self.skip_holidays_var.set(False)
            # Update end date visibility based on recurrence
            self._update_end_date_visibility()

    def _update_end_date_visibility(self):
        """Show or hide end date labels based on recurrence type"""
        rec = self.recurrence_var.get()
        if rec == "None":
            # Hide end date labels for non-recurring tasks
            self.end_label.pack_forget()
            self.end_value_label.pack_forget()
            self.end_date_var.set("")  # Clear end date
        else:
            # Show end date labels for recurring tasks
            self.end_label.pack(anchor="w")
            self.end_value_label.pack(anchor="w")
    
    def _on_repeat_toggle(self):
        # When toggling repeat indefinitely, update end label string and behaviour of picker
        if self.repeat_indef_var.get():
            # set end display to Indefinite and don't require end in picker
            self.end_date_var.set("Indefinite")
        else:
            # clear end so user knows to pick one
            if self.end_date_var.get() == "Indefinite":
                self.end_date_var.set("")

    def update_recurrence_options(self, event=None):
        # clear options frame
        for w in self.recur_options_frame.winfo_children():
            w.destroy()
        rec = self.recurrence_var.get()
        if rec in ("Weekly", "Bi-Weekly"):
            self.weekday_vars = []
            days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
            for i, d in enumerate(days):
                var = tk.IntVar()
                cb = ttk.Checkbutton(self.recur_options_frame, text=d, variable=var)
                cb.pack(anchor="w")
                self.weekday_vars.append((i, var))
        elif rec == "Monthly":
            # Monthly now uses week-based selection
            ttk.Label(self.recur_options_frame, text="Week:").pack(anchor="w")
            self.monthly_week_var = tk.IntVar(value=1)
            week_frame = ttk.Frame(self.recur_options_frame)
            week_frame.pack(anchor="w")
            for i in range(1, 5):
                rb = ttk.Radiobutton(week_frame, text=f"{i}{'st' if i == 1 else 'nd' if i == 2 else 'rd' if i == 3 else 'th'}", 
                                     variable=self.monthly_week_var, value=i)
                rb.pack(side="left", padx=5)
        else:
            # nothing to show for None, Daily, Annually
            pass

    def save_task(self, editing=False):
        desc = self.task_desc_entry.get().strip()
        if not desc:
            messagebox.showerror("Error", "Description required")
            return

        time_slot = self.time_slot_var.get()
        if not time_slot:
            messagebox.showerror("Error", "Select a time slot")
            return

        # Read from checkboxes instead of listbox
        who = [name for name, var in getattr(self, 'profile_vars', {}).items() if var.get()]
        if not who:
            who = ["Unassigned"]

        rec = self.recurrence_var.get()

        # Get skip_holidays flag
        skip_holidays = getattr(self, 'skip_holidays_var', tk.BooleanVar(value=False)).get()
        
        # parse start date
        try:
            start_date = datetime.strptime(self.start_date_var.get(), "%Y-%m-%d")
        except Exception:
            messagebox.showerror("Error", "Bad start date (use Pick Start/End Dates)")
            return
        if rec == "Monthly":
            start_date = _start_of_week(start_date)
        if start_date.weekday() >= 5:
            start_date = _coerce_to_weekday(start_date)
        if skip_holidays:
            start_date = _skip_holidays(start_date)
            self.start_date_var.set(start_date.strftime("%Y-%m-%d"))
        if start_date.weekday() >= 5:
            messagebox.showerror("Error", "Start date must be a weekday (Monday-Friday)")
            return

        # parse end date unless indefinite or non-recurring
        if rec == "None":
            # Non-recurring tasks don't need end date
            end_date = None
        elif self.repeat_indef_var.get() or self.end_date_var.get() == "Indefinite":
            end_date = None
        else:
            try:
                end_date = datetime.strptime(self.end_date_var.get(), "%Y-%m-%d")
            except Exception:
                messagebox.showerror("Error", "Bad end date (use Pick Start/End Dates or check Repeat Indefinitely)")
                return
            if end_date < start_date:
                messagebox.showerror("Error", "End date before start")
                return
            if rec == "Monthly":
                end_date = _start_of_week(end_date)
            if end_date.weekday() >= 5:
                end_date = _coerce_to_weekday(end_date)
            if skip_holidays:
                end_date = _skip_holidays(end_date)
                self.end_date_var.set(end_date.strftime("%Y-%m-%d"))
            if end_date.weekday() >= 5:
                messagebox.showerror("Error", "End date must be a weekday (Monday-Friday)")
                return

        weekly_days = []
        monthly_day = None
        if rec in ("Weekly", "Bi-Weekly"):
            weekly_days = [d for d, v in getattr(self, "weekday_vars", []) if v.get() == 1]
            if not weekly_days:
                messagebox.showerror("Error", "Select weekday(s)")
                return
        elif rec == "Monthly":
            # Monthly uses week number (1-4) stored as monthly_day
            # This represents which week of the month to repeat (1st, 2nd, 3rd, or 4th)
            monthly_day = getattr(self, "monthly_week_var", tk.IntVar(value=1)).get()

        sync_target = None
        if editing and self.current_edit_task:
            t = self.current_edit_task
            t.description = desc
            t.recurrence = rec
            t.start_date = start_date
            t.end_date = end_date
            t.weekly_days = weekly_days
            t.monthly_day = monthly_day
            t.who = who
            t.time_slot = time_slot
            t.skip_holidays = skip_holidays
            sync_target = t  # Store reference for sync
        else:
            new_task = Task(desc, rec, start_date, end_date, weekly_days, monthly_day, who, time_slot=time_slot, skip_holidays=skip_holidays)
            self.tasks.append(new_task)
            sync_target = new_task

        messagebox.showinfo("Saved", "Task saved")

        # ---------- Google sync: create/update event ----------
        try:
            service = get_calendar_service()
            
            if sync_target:
                # Remove events only for assignees that were removed, do NOT delete all
                _prune_removed_assignee_events(service, sync_target, self.profiles)
                
                # Sync to all appropriate calendars (create/update per assignee)
                event_ids = sync_task_to_calendars(service, sync_target, self.profiles)
                if event_ids:
                    sync_target.event_ids = event_ids
                    # Keep legacy event_id for backward compatibility (use first one)
                    if event_ids:
                        sync_target.event_id = list(event_ids.values())[0]
                    self.save_data()

            if editing:
                self.current_edit_task = None  # Clear after sync

        except Exception as e:
            # don't block user flow if calendar fails — log to console
            print("Auto-sync failed:", e)
            try:
                messagebox.showerror("Google Sync Error", f"Failed to sync to Google Calendar:\n{e}")
            except Exception:
                pass

        # return to All Tasks if we came from there, otherwise day or calendar
        if self.came_from_all_tasks:
            self.came_from_all_tasks = False
            self.show_all_tasks()
        else:
            if self.selected_day:
                try:
                    self.show_day_view(self.selected_day)
                except Exception:
                    self.show_calendar()
            else:
                self.show_calendar()

    def edit_task(self, task):
        # called from day view
        self.current_edit_task = task
        self.came_from_all_tasks = False
        self.show_task_create(editing=True)

    def edit_task_from_all(self, task):
        # called from All Tasks full page
        self.current_edit_task = task
        self.came_from_all_tasks = True
        # hide all_tasks_frame and ensure left_frame visible for editor
        self.all_tasks_frame.pack_forget()
        if not self.left_frame.winfo_ismapped():
            self.left_frame.pack(side="left", fill="both", expand=True)
        self.topbar.pack(fill="x")
        self.show_task_create(editing=True)

    def delete_task(self, task):
        if messagebox.askyesno("Confirm", "Delete this task?"):
            try:
                # First try deleting from Google Calendar
                try:
                    service = get_calendar_service()
                    delete_task_from_calendars(service, task, self.profiles)
                except Exception as e:
                    print("Error deleting google events:", e)

                # Then delete from local list
                self.tasks.remove(task)

            except ValueError:
                pass

            # refresh current visible view
            if self.all_tasks_frame.winfo_ismapped():
                self.show_all_tasks()
            elif self.selected_day:
                self.show_day_view(self.selected_day)
            else:
                self.show_calendar()

    def delete_task_from_all(self, task):
        if messagebox.askyesno("Confirm", "Delete this task?"):
            try:
                # First try deleting from Google Calendar
                try:
                    service = get_calendar_service()
                    delete_task_from_calendars(service, task, self.profiles)
                except Exception as e:
                    print("Error deleting google events:", e)

                # Then delete from local list
                self.tasks.remove(task)

            except ValueError:
                pass

            # refresh current visible view...
            self.show_all_tasks()

    def show_day_view_cancel(self):
        # used by Cancel inside task editor
        self.task_create_frame.pack_forget()
        if self.came_from_all_tasks:
            self.came_from_all_tasks = False
            self.show_all_tasks()
        elif self.selected_day:
            self.show_day_view(self.selected_day)
        else:
            self.show_calendar()

    # ---------- profiles helpers ----------
    def _populate_assigned_listbox(self):
        try:
            self.assigned_listbox.delete(0, tk.END)
            for p in self.profiles:
                if isinstance(p, Profile):
                    self.assigned_listbox.insert(tk.END, p.name)
                else:
                    self.assigned_listbox.insert(tk.END, p)
        except Exception:
            pass

    def _refresh_profiles_listboxes(self):
        # Legacy method for backward compatibility
        if hasattr(self, "profiles_page_listbox"):
            self.profiles_page_listbox.delete(0, tk.END)
            for p in self.profiles:
                if isinstance(p, Profile):
                    self.profiles_page_listbox.insert(tk.END, p.name)
                else:
                    self.profiles_page_listbox.insert(tk.END, p)
        # update assigned_listbox if it exists
        try:
            self._populate_assigned_listbox()
        except Exception:
            pass

    def _refresh_profiles_tree(self):
        if hasattr(self, "profiles_tree"):
            for item in self.profiles_tree.get_children():
                self.profiles_tree.delete(item)
            for p in self.profiles:
                if isinstance(p, Profile):
                    calendar_status = "Created" if p.calendar_id else "Not Created"
                    email_display = p.email if p.email else "No email"
                    admin_status = "Yes" if p.is_admin else "No"
                    self.profiles_tree.insert("", "end", values=(p.name, email_display, admin_status, calendar_status, "Edit"))
                else:
                    # Handle legacy string profiles
                    self.profiles_tree.insert("", "end", values=(p, "No email", "No", "Not Created", "Edit"))

        # Force Treeview heading background to remain unchanged on hover
        style = ttk.Style()
        orig_bg = style.lookup('Treeview.Heading', 'background')
        style.map('Treeview.Heading', background=[('active', orig_bg), ('!active', orig_bg)])

    def add_profile(self):
        name = getattr(self, "new_profile_entry", None).get().strip() if hasattr(self, "new_profile_entry") else ""
        if not name:
            messagebox.showerror("Error", "Profile name required")
            return
        
        # Check if profile name already exists
        existing_names = [p.name if isinstance(p, Profile) else p for p in self.profiles]
        if name in existing_names:
            messagebox.showerror("Error", "Profile already exists")
            return
        
        email = getattr(self, "new_profile_email_entry", None).get().strip() if hasattr(self, "new_profile_email_entry") else ""
        is_admin = getattr(self, "new_profile_admin_var", None).get() if hasattr(self, "new_profile_admin_var") else False
        new_profile = Profile(name, email if email else None, is_admin=is_admin)
        
        # Automatically create Google Calendar for the profile
        try:
            service = get_calendar_service()
            calendar_id = create_profile_calendar(service, name)
            if calendar_id:
                new_profile.calendar_id = calendar_id
                
                # Share calendar with email if provided
                if email:
                    if share_calendar_with_email(service, calendar_id, email):
                        messagebox.showinfo("Success", f"Profile '{name}' created with Google Calendar and shared with {email}")
                    else:
                        messagebox.showwarning("Warning", f"Profile '{name}' created but failed to share calendar with {email}")
                else:
                    messagebox.showinfo("Success", f"Profile '{name}' created with Google Calendar")
            else:
                messagebox.showwarning("Warning", f"Profile '{name}' created but failed to create Google Calendar")
        except Exception as e:
            messagebox.showwarning("Warning", f"Profile '{name}' created but failed to create Google Calendar: {e}")
        
        self.profiles.append(new_profile)
        
        # Clear the entry fields
        if hasattr(self, "new_profile_entry"):
            self.new_profile_entry.delete(0, tk.END)
        if hasattr(self, "new_profile_email_entry"):
            self.new_profile_email_entry.delete(0, tk.END)
        
        self._refresh_profiles_listboxes()
        if hasattr(self, "profiles_tree"):
            self._refresh_profiles_tree()
        
        # Save the data to persist the new profile and calendar ID
        self.save_data()

    def delete_profiles(self):
        if hasattr(self, "profiles_tree"):
            # New treeview-based deletion
            selected_items = self.profiles_tree.selection()
            if not selected_items:
                messagebox.showerror("Error", "No profile selected")
                return
            
            if messagebox.askyesno("Confirm", "Delete selected profile(s) and their calendars? This will also delete tasks assigned only to these profiles."):
                try:
                    service = get_calendar_service()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to connect to Google Calendar: {e}")
                    return
                
                deleted_tasks_count = 0
                for item in selected_items:
                    values = self.profiles_tree.item(item, "values")
                    profile_name = values[0]
                    
                    # Clean up tasks for this profile
                    tasks_to_remove = cleanup_tasks_for_deleted_profile(self.tasks, profile_name, service, self.profiles)
                    deleted_tasks_count += len(tasks_to_remove)
                    
                    # Remove the tasks from the tasks list
                    for task in tasks_to_remove:
                        if task in self.tasks:
                            self.tasks.remove(task)
                    
                    # Find and remove the profile
                    for i, p in enumerate(self.profiles):
                        if isinstance(p, Profile) and p.name == profile_name:
                            # Delete the Google Calendar if it exists
                            if p.calendar_id:
                                delete_profile_calendar(service, p.calendar_id)
                            del self.profiles[i]
                            break
                        elif p == profile_name:  # Legacy string profile
                            del self.profiles[i]
                            break
                
                self._refresh_profiles_tree()
                self._refresh_profiles_listboxes()
                self.save_data()
                
                # Show summary of what was deleted
                if deleted_tasks_count > 0:
                    messagebox.showinfo("Profile Deleted", f"Profile(s) deleted successfully.\n{deleted_tasks_count} task(s) were also deleted as they were assigned only to the deleted profile(s).")
                else:
                    messagebox.showinfo("Profile Deleted", "Profile(s) deleted successfully.")
        else:
            # Legacy listbox-based deletion
            if not hasattr(self, "profiles_page_listbox"):
                messagebox.showerror("Error", "No profiles list to delete from")
                return
            sel = list(self.profiles_page_listbox.curselection())
            if not sel:
                messagebox.showerror("Error", "No profile selected")
                return
            
            if messagebox.askyesno("Confirm", "Delete selected profile(s) and their calendars? This will also delete tasks assigned only to these profiles."):
                try:
                    service = get_calendar_service()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to connect to Google Calendar: {e}")
                    return
                
                deleted_tasks_count = 0
                for idx in reversed(sel):
                    name = self.profiles_page_listbox.get(idx)
                    
                    # Clean up tasks for this profile
                    tasks_to_remove = cleanup_tasks_for_deleted_profile(self.tasks, name, service, self.profiles)
                    deleted_tasks_count += len(tasks_to_remove)
                    
                    # Remove the tasks from the tasks list
                    for task in tasks_to_remove:
                        if task in self.tasks:
                            self.tasks.remove(task)
                    
                    for i, p in enumerate(self.profiles):
                        if isinstance(p, Profile) and p.name == name:
                            # Delete the Google Calendar if it exists
                            if p.calendar_id:
                                delete_profile_calendar(service, p.calendar_id)
                            del self.profiles[i]
                            break
                        elif p == name:  # Legacy string profile
                            del self.profiles[i]
                            break
                
                self._refresh_profiles_listboxes()
                self.save_data()
                
                # Show summary of what was deleted
                if deleted_tasks_count > 0:
                    messagebox.showinfo("Profile Deleted", f"Profile(s) deleted successfully.\n{deleted_tasks_count} task(s) were also deleted as they were assigned only to the deleted profile(s).")
                else:
                    messagebox.showinfo("Profile Deleted", "Profile(s) deleted successfully.")

    def reset_google_credentials(self):
        """Reset Google Calendar credentials to blank"""
        if messagebox.askyesno("Confirm Reset", "Reset Google Calendar credentials? You will be prompted to enter them again on next startup."):
            try:
                # Blank out the credentials file
                blank_creds = {
                    "installed": {
                        "client_id": "",
                        "project_id": "",
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                        "client_secret": "",
                        "redirect_uris": ["http://localhost"]
                    }
                }
                with open(CREDENTIALS_FILE, 'w') as f:
                    json.dump(blank_creds, f, indent=2)
                
                # Also delete the token file if it exists
                if os.path.exists(TOKEN_FILE):
                    try:
                        os.remove(TOKEN_FILE)
                    except Exception:
                        pass
                
                messagebox.showinfo("Success", "Google Calendar credentials have been reset. Restart the app to enter new credentials.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to reset credentials: {e}")

    def share_calendar(self):
        """Share a calendar with an email address"""
        if not hasattr(self, "profiles_tree"):
            messagebox.showerror("Error", "No profiles tree available")
            return
        
        selected_items = self.profiles_tree.selection()
        if not selected_items:
            messagebox.showerror("Error", "No profile selected")
            return
        
        if len(selected_items) > 1:
            messagebox.showerror("Error", "Please select only one profile")
            return
        
        item = selected_items[0]
        values = self.profiles_tree.item(item, "values")
        profile_name = values[0]
        
        # Find the profile object
        profile = None
        for p in self.profiles:
            if isinstance(p, Profile) and p.name == profile_name:
                profile = p
                break
        
        if not profile:
            messagebox.showerror("Error", "Profile not found")
            return
        
        if not profile.calendar_id:
            messagebox.showerror("Error", "Profile has no calendar to share")
            return
        
        # Get email from user
        email = tk.simpledialog.askstring("Share Calendar", f"Enter email address to share {profile_name}'s calendar with:")
        if not email:
            return
        
        try:
            service = get_calendar_service()
            if share_calendar_with_email(service, profile.calendar_id, email):
                messagebox.showinfo("Success", f"Calendar shared with {email}")
            else:
                messagebox.showerror("Error", f"Failed to share calendar with {email}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to share calendar: {e}")

    def edit_profile(self, profile_name):
        """Edit a profile's name and email"""
        # Find the profile object
        profile = None
        for p in self.profiles:
            if isinstance(p, Profile) and p.name == profile_name:
                profile = p
                break
        
        if not profile:
            messagebox.showerror("Error", "Profile not found")
            return
        
        # Create edit dialog
        edit_window = tk.Toplevel(self.root)
        edit_window.title(f"Edit Profile: {profile_name}")
        edit_window.geometry("450x300")
        edit_window.transient(self.root)
        edit_window.grab_set()
        
        # Name field
        ttk.Label(edit_window, text="Name:").pack(pady=5)
        name_var = tk.StringVar(value=profile.name)
        name_entry = ttk.Entry(edit_window, textvariable=name_var, width=40)
        name_entry.pack(pady=5)
        
        # Email field
        ttk.Label(edit_window, text="Email (optional):").pack(pady=5)
        email_var = tk.StringVar(value=profile.email or "")
        email_entry = ttk.Entry(edit_window, textvariable=email_var, width=40)
        email_entry.pack(pady=5)
        
        # Admin checkbox
        admin_var = tk.BooleanVar(value=profile.is_admin)
        admin_checkbox = ttk.Checkbutton(edit_window, text="Administrator (can see all tasks)", variable=admin_var)
        admin_checkbox.pack(pady=5)
        
        # Buttons
        button_frame = ttk.Frame(edit_window)
        button_frame.pack(pady=20, fill="x", padx=20)
        
        def share_calendar_from_edit():
            if not profile.calendar_id:
                messagebox.showerror("Error", "Profile has no calendar to share")
                return
            
            # Get email from user
            email = tk.simpledialog.askstring("Share Calendar", f"Enter email address to share {profile.name}'s calendar with:")
            if not email:
                return
            
            try:
                service = get_calendar_service()
                if share_calendar_with_email(service, profile.calendar_id, email):
                    messagebox.showinfo("Success", f"Calendar shared with {email}")
                else:
                    messagebox.showerror("Error", f"Failed to share calendar with {email}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to share calendar: {e}")
        
        def save_changes():
            new_name = name_var.get().strip()
            new_email = email_var.get().strip()
            new_admin = admin_var.get()
            
            if not new_name:
                messagebox.showerror("Error", "Name is required")
                return
            
            # Check if new name conflicts with existing profiles
            if new_name != profile_name:
                existing_names = [p.name if isinstance(p, Profile) else p for p in self.profiles if p != profile]
                if new_name in existing_names:
                    messagebox.showerror("Error", "Profile name already exists")
                    return
            
            # Update profile
            old_name = profile.name
            profile.name = new_name
            profile.email = new_email if new_email else None
            profile.is_admin = new_admin
            
            # Update Google Calendar name if it exists
            if profile.calendar_id:
                try:
                    service = get_calendar_service()
                    update_calendar_name(service, profile.calendar_id, new_name)
                except Exception as e:
                    messagebox.showwarning("Warning", f"Profile updated but failed to update Google Calendar name: {e}")
            
            # Update tasks that reference this profile
            for task in self.tasks:
                if old_name in task.who:
                    # Replace old name with new name in task assignments
                    task.who = [new_name if name == old_name else name for name in task.who]
            
            self._refresh_profiles_tree()
            self._refresh_profiles_listboxes()
            self.save_data()
            edit_window.destroy()
            messagebox.showinfo("Success", f"Profile updated successfully")
        
        ttk.Button(button_frame, text="Save", command=save_changes).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Share Calendar", command=share_calendar_from_edit).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Cancel", command=edit_window.destroy).pack(side="right", padx=5)

    def on_profile_double_click(self, event):
        """Handle double-click on profile to edit it"""
        item = self.profiles_tree.selection()[0] if self.profiles_tree.selection() else None
        if item:
            values = self.profiles_tree.item(item, "values")
            profile_name = values[0]
            self.edit_profile(profile_name)

    def manual_sync(self):
        """Manually sync from Google Calendar"""
        try:
            service = get_calendar_service()
            deleted_count, modified_count = sync_from_google_calendar(service, self.tasks, self.profiles)
            
            # Refresh the current view
            if self.all_tasks_frame.winfo_ismapped():
                self.show_all_tasks()
            elif self.selected_day:
                self.show_day_view(self.selected_day)
            else:
                self.show_calendar()
            
            # Save changes
            self.save_data()
            
            # Show sync results
            if deleted_count > 0 or modified_count > 0:
                messagebox.showinfo("Sync Complete", f"Sync completed successfully!\n{deleted_count} task(s) removed\n{modified_count} task(s) updated")
            else:
                messagebox.showinfo("Sync Complete", "Sync completed successfully!\nNo changes found.")
                
        except Exception as e:
            messagebox.showerror("Sync Error", f"Failed to sync from Google Calendar: {e}")

    def sync_on_startup(self):
        """Sync from Google Calendar on app startup (runs in background)"""
        def background_sync():
            try:
                service = get_calendar_service()
                deleted_count, modified_count = sync_from_google_calendar(service, self.tasks, self.profiles)
                
                if deleted_count > 0 or modified_count > 0:
                    # Save changes and refresh UI
                    self.save_data()
                    # Schedule UI refresh on main thread
                    self.root.after(0, self._refresh_ui_after_sync)
                    print(f"Startup sync: {deleted_count} tasks removed, {modified_count} tasks updated")
                else:
                    print("Startup sync: No changes found")
                    
            except Exception as e:
                print(f"Startup sync failed: {e}")
        
        # Run sync in background thread to avoid blocking UI
        import threading
        sync_thread = threading.Thread(target=background_sync, daemon=True)
        sync_thread.start()

    def _refresh_ui_after_sync(self):
        """Refresh UI after background sync completes"""
        try:
            if self.all_tasks_frame.winfo_ismapped():
                self.show_all_tasks()
            elif self.selected_day:
                self.show_day_view(self.selected_day)
            else:
                self.show_calendar()
        except:
            pass  # Ignore errors during UI refresh

    # ---------- All Tasks (full page) ----------
    def show_all_tasks(self):
        # make the all_tasks_frame full width by hiding left_frame
        # Hide other left-side pages so only All Tasks is visible
        self.calendar_frame.pack_forget()
        self.day_view_frame.pack_forget()
        self.task_create_frame.pack_forget()
        # Ensure Profiles page is hidden when showing All Tasks
        try:
            self.profiles_page.pack_forget()
        except Exception:
            pass

        if self.left_frame.winfo_ismapped():
            self.left_frame.pack_forget()

        for w in self.all_tasks_frame.winfo_children():
            w.destroy()
        self.all_tasks_frame.pack(fill="both", expand=True)

        ttk.Label(self.all_tasks_frame, text="All Tasks", font=("Arial", 16, "bold")).pack(pady=10)

        # Ensure tasks are being added correctly
        print("Debug: Verifying task addition")
        for task in self.tasks:
            print(f"Task: {task.description}, Recurrence: {task.recurrence}, Assigned: {task.who}")

        # Fix categorization logic
        categories = {
            "None": [],
            "Daily": [],
            "Weekly": [],
            "Bi-Weekly": [],
            "Monthly": [],
            "Annually": []
        }
        for t in self.tasks:
            rec = getattr(t, 'recurrence', "None")
            if rec not in categories:
                print(f"Warning: Task with unexpected recurrence '{rec}'")
                categories.setdefault("Uncategorized", []).append(t)
            else:
                categories[rec].append(t)

        # Render tasks in the UI
        for rec, tasks_in_cat in categories.items():
            section = ttk.Frame(self.all_tasks_frame, padding=(6, 6))
            section.pack(fill="x", expand=False, padx=8, pady=6)

            header = tk.Label(section, text=rec, font=("Arial", 11, "bold"), anchor="w", fg="#FFFFFF", bg=self.root.cget('bg'))
            header.pack(fill="x", pady=(0, 6))

            inner = ttk.Frame(section)
            inner.pack(fill="x", expand=True)

            if not tasks_in_cat:
                placeholder = tk.Label(inner, text="No tasks", anchor="w", fg="#cfcfcf", bg=self.root.cget('bg'), width=80)
                placeholder.pack(anchor="w", pady=4)
            else:
                for idx, t in enumerate(tasks_in_cat):
                    assigned = "Unassigned" if not t.who else ", ".join(t.who)
                    row = ttk.Frame(inner, padding=(4, 6))
                    row.pack(fill="x")
                    txt = tk.Label(row, text=f"- [{t.time_slot}] {t.description}  | {assigned}", anchor="w", fg="#e6e6e6", bg=self.root.cget('bg'))
                    txt.pack(side="left", fill="x", expand=True)
                    btn_del = ttk.Button(row, text="Delete", command=lambda task=t: self.delete_task_from_all(task))
                    btn_edit = ttk.Button(row, text="Edit", command=lambda task=t: self.edit_task_from_all(task))
                    btn_del.pack(side="right", padx=2)
                    btn_edit.pack(side="right", padx=2)

                    if idx != len(tasks_in_cat) - 1:
                        sep = ttk.Separator(inner, orient='horizontal')
                        sep.pack(fill='x', pady=(6, 6))

        # Place Back button in the parent frame so it's always visible
        back_btn = ttk.Button(self.all_tasks_frame, text="Back", command=self._close_all_tasks)
        back_btn.pack(side="bottom", pady=10)

    def _close_all_tasks(self):
        # hide all_tasks_frame and restore left_frame + topbar
        self.all_tasks_frame.pack_forget()
        if not self.left_frame.winfo_ismapped():
            self.left_frame.pack(side="left", fill="both", expand=True)
        self.topbar.pack(fill="x")
        self.show_calendar()

    # ---------- Export PDF ----------
    def export_pdf(self):
        start_var, end_var = tk.StringVar(), tk.StringVar()
        picker = DateRangePicker(self.root, start_var, end_var)  # always allow end here
        self.root.wait_window(picker)
        if not start_var.get() or not end_var.get():
            return
        try:
            start_date = datetime.strptime(start_var.get(), "%Y-%m-%d")
            end_date = datetime.strptime(end_var.get(), "%Y-%m-%d")
        except Exception:
            messagebox.showerror("Error", "Invalid date format from picker")
            return
        if end_date < start_date:
            messagebox.showerror("Error", "End must be on/after start")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialfile="Work_Schedule.pdf")
        if not file_path:
            return

        c = canvas.Canvas(file_path, pagesize=LETTER)
        width, height = LETTER

        current = start_date
        while current <= end_date:
            # weekdays only (Mon-Fri)
            if current.weekday() < 5:
                y = height - 60
                c.setFont("Helvetica-Bold", 14)
                c.drawString(50, y, current.strftime("%A, %B %d, %Y"))
                y -= 30
                tasks_today = [t for t in self.tasks if t.occurs_on(current)]
                c.setFont("Helvetica", 11)
                if not tasks_today:
                    c.drawString(60, y, "No tasks scheduled.")
                else:
                    for t in tasks_today:
                        assigned = "Unassigned"
                        if t.who and t.who != ["Unassigned"]:
                            assigned = ", ".join(t.who) if len(t.who) <= 2 else f"{len(t.who)} people"
                        text = f"- [{t.time_slot}] {t.description} ({assigned})"
                        c.drawString(60, y, text)
                        y -= 18
                        if y < 50:
                            c.showPage()
                            y = height - 60
                            c.setFont("Helvetica", 11)
                c.showPage()
            current += timedelta(days=1)

        c.save()
        messagebox.showinfo("PDF Exported", f"Saved as {os.path.basename(file_path)}")

            
    # ---------- persistence ----------
    def save_data(self):
        data = {
            "tasks": [
                {
                    "description": t.description,
                    "recurrence": t.recurrence,
                    "start_date": t.start_date.strftime("%Y-%m-%d"),
                    "end_date": t.end_date.strftime("%Y-%m-%d") if t.end_date else None,
                    "weekly_days": t.weekly_days,
                    "monthly_day": t.monthly_day,
                    "who": t.who,
                    "time_slot": t.time_slot,
                    "event_id": t.event_id,  # Legacy field
                    "event_ids": t.event_ids,  # New field for multiple events
                    "skip_holidays": getattr(t, 'skip_holidays', False)
                } for t in self.tasks
            ],
            "profiles": [
                {
                    "name": p.name,
                    "email": p.email,
                    "calendar_id": p.calendar_id
                } if isinstance(p, Profile) else {"name": p, "email": None, "calendar_id": None}
                for p in self.profiles
            ]
        }
        with open(SAVE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    def load_data(self):
        if not os.path.exists(SAVE_FILE):
            return
        with open(SAVE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # Load profiles
        self.profiles = []
        for pd in data.get("profiles", []):
            if isinstance(pd, dict):
                profile = Profile(
                    pd["name"],
                    pd.get("email"),
                    pd.get("calendar_id"),
                    pd.get("is_admin", False)
                )
                self.profiles.append(profile)
            else:
                # Handle legacy string profiles
                self.profiles.append(Profile(pd, None))
        
        # Load tasks
        self.tasks = []
        for td in data.get("tasks", []):
            start = datetime.strptime(td["start_date"], "%Y-%m-%d")
            end = datetime.strptime(td["end_date"], "%Y-%m-%d") if td.get("end_date") else None
            t = Task(td["description"],
                     td.get("recurrence", "None"),
                     start,
                     end,
                     td.get("weekly_days", []),
                     td.get("monthly_day"),
                     td.get("who", []),
                     td.get("time_slot"),
                     event_id=td.get("event_id"),
                     event_ids=td.get("event_ids", {}),
                     skip_holidays=td.get("skip_holidays", False))
            self.tasks.append(t)
        
        # Sync from Google Calendar on startup (in background)
        self.sync_on_startup()

    def on_close(self):
        self.save_data()
        self.root.destroy()

    # ---------- month navigation ----------
    def prev_month(self):
        month = self.month_var.get()
        year = self.year_var.get()
        if month == 1:
            month = 12
            year -= 1
        else:
            month -= 1
        self.month_var.set(month)
        self.year_var.set(year)
        self.show_calendar()

    def next_month(self):
        month = self.month_var.get()
        year = self.year_var.get()
        if month == 12:
            month = 1
            year += 1
        else:
            month += 1
        self.month_var.set(month)
        self.year_var.set(year)
        self.show_calendar()
