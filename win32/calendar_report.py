#!/usr/bin/env python3
"""
Calendar Category Usage Report

This script queries your Outlook calendar for meetings in the current workweek
and produces a report showing time spent in each category versus your budgets.
Uses the same Outlook COM interface technique as outlook_calendar_search.py.
"""

import win32com.client
import datetime
import sys
import subprocess
import time
import pytz
from collections import defaultdict
from typing import Dict, List

# ---------- Global Variables and Configuration ----------
DEBUG_MODE = "--verbose" in sys.argv  # pass --verbose to see extra details

# Display help/usage information
if "--help" in sys.argv or "-h" in sys.argv:
    print(
        "Calendar Category Usage Report\n\n"
        "Usage: python calendar_report.py [options]\n\n"
        "Options:\n"
        "  --lastweek   Analyze previous workweek (Mon-Fri)\n"
        "  --nextweek   Analyze next workweek\n"
        "  --start DATE Analyze starting from DATE (format: yyyy-MM-dd)\n"
        "  --end DATE   Analyze ending at DATE (format: yyyy-MM-dd)\n"
        "  --verbose    Show extra details in output\n"
        "  --help, -h   Show this help message and exit\n"
    )
    sys.exit(0)

# Parse --start parameter
START_DATE = None
if "--start" in sys.argv:
    try:
        start_idx = sys.argv.index("--start")
        if start_idx + 1 >= len(sys.argv):
            print("Error: --start requires a date argument (format: yyyy-MM-dd)")
            sys.exit(1)
        start_date_str = sys.argv[start_idx + 1]
        START_DATE = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
    except ValueError as e:
        print(f"Error: Invalid date format for --start. Expected yyyy-MM-dd, got: {start_date_str}")
        print(f"Example: --start 2024-01-15")
        sys.exit(1)

# Parse --end parameter
END_DATE = None
if "--end" in sys.argv:
    try:
        end_idx = sys.argv.index("--end")
        if end_idx + 1 >= len(sys.argv):
            print("Error: --end requires a date argument (format: yyyy-MM-dd)")
            sys.exit(1)
        end_date_str = sys.argv[end_idx + 1]
        END_DATE = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()
    except ValueError as e:
        print(f"Error: Invalid date format for --end. Expected yyyy-MM-dd, got: {end_date_str}")
        print(f"Example: --end 2024-01-20")
        sys.exit(1)

# Validate --end doesn't precede --start
if START_DATE is not None and END_DATE is not None:
    if END_DATE < START_DATE:
        print(f"Error: --end date ({END_DATE}) cannot precede --start date ({START_DATE})")
        sys.exit(1)

# Validate --end is not used without --start
if END_DATE is not None and START_DATE is None:
    print("Error: --end requires --start to be specified")
    sys.exit(1)

# Determine which week to analyze based on command-line flags
if START_DATE is not None and ("--lastweek" in sys.argv or "--nextweek" in sys.argv):
    print("Error: Cannot specify --start with --lastweek or --nextweek flags.")
    sys.exit(1)

if "--lastweek" in sys.argv and "--nextweek" in sys.argv:
    print("Error: Cannot specify both --lastweek and --nextweek flags.")
    sys.exit(1)
elif "--lastweek" in sys.argv:
    WEEK_OFFSET = -1
elif "--nextweek" in sys.argv:
    WEEK_OFFSET = 1
else:
    WEEK_OFFSET = 0

# Define work hours for each day of the week (Monday=0, Sunday=6)
DAILY_WORK_HOURS = {
    0: 8,  # Monday: 9am-5pm (8 hours)
    1: 8,  # Tuesday: 9am-5pm (8 hours)
    2: 8,  # Wednesday: 9am-5pm (8 hours)
    3: 9,  # Thursday: 8am-5pm (9 hours)
    4: 9,  # Friday: 8am-5pm (9 hours)
    5: 0,  # Saturday: 0 hours
    6: 0,  # Sunday: 0 hours
}

# Define work start times for each day (Monday=0, Sunday=6)
DAILY_START_TIMES = {
    0: 9,  # Monday: 9am
    1: 9,  # Tuesday: 9am
    2: 9,  # Wednesday: 9am
    3: 8,  # Thursday: 8am
    4: 8,  # Friday: 8am
    5: 9,  # Saturday: 9am (not used)
    6: 9,  # Sunday: 9am (not used)
}

# Calculate total work hours for the week
TOTAL_WORK_HOURS = sum(hours for day, hours in DAILY_WORK_HOURS.items() if day < 5)

TIMEZONE = "US/Pacific"

BUDGETS: Dict[str, Dict[str, int]] = {
    "Focus Time": {"min": 12, "max": 15, "warn": 14},
    "Collaboration": {"min": 4, "max": 8, "warn": 6},
    "Communication": {"min": 0, "max": 8, "warn": 6},
    "Work Meeting": {"min": 0, "max": 12, "warn": 10},
    "Unavailable": {"min": 0, "max": 6, "warn": 5},
}

KNOWN_CATEGORIES = {
    "Focus Time",
    "Communication",
    "Unavailable", 
    "Collaboration",
    "Holiday/Vacation",
}

class OutlookCalendarReporter:
    """Class to handle Outlook calendar reporting via COM."""
    
    def __init__(self):
        """Initialize the Outlook application connection."""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.calendar_folder = self.namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
            self.pacific_tz = pytz.timezone(TIMEZONE)
        except Exception as e:
            raise ConnectionError(f"Failed to connect to Outlook: {e}") from e

    def convert_to_pacific(self, start_time: datetime.datetime, end_time: datetime.datetime) -> tuple:
        """Convert datetime objects to Pacific Time, handling Outlook's timezone quirks."""
        # Convert to datetime objects if needed
        if hasattr(start_time, 'year'):
            start_time = datetime.datetime(
                start_time.year, start_time.month, start_time.day,
                start_time.hour, start_time.minute, start_time.second
            )
            end_time = datetime.datetime(
                end_time.year, end_time.month, end_time.day,
                end_time.hour, end_time.minute, end_time.second
            )
        
        # Handle timezone conversion
        if start_time.tzinfo is None:
            # For timezone-naive datetimes, assume they're already in Pacific Time
            start_time = self.pacific_tz.localize(start_time)
            end_time = self.pacific_tz.localize(end_time)
        else:
            # Check if it's UTC (Outlook incorrectly marking Pacific times as UTC)
            if (hasattr(start_time.tzinfo, 'utcoffset') and 
                start_time.tzinfo.utcoffset(start_time).total_seconds() == 0):
                # If it's UTC offset 0, the time values are actually Pacific times
                # with incorrect UTC timezone markers - treat them as Pacific directly
                start_time = start_time.replace(tzinfo=self.pacific_tz)
                end_time = end_time.replace(tzinfo=self.pacific_tz)
        
        return start_time, end_time

    def calculate_work_hours_overlap(self, start_time: datetime.datetime, end_time: datetime.datetime) -> float:
        """Calculate how many hours of an event overlap with work hours in Pacific Time.
        Assumes start_time and end_time are already in Pacific Time."""
        start_pt = start_time
        end_pt = end_time
        
        total_overlap = 0.0
        
        # Process each day the event spans
        current_date = start_pt.date()
        end_date = end_pt.date()
        
        while current_date <= end_date:
            # Get work hours for this day from configuration
            weekday = current_date.weekday()
            daily_hours = DAILY_WORK_HOURS.get(weekday, 0)
            start_hour = DAILY_START_TIMES.get(weekday, 9)
            
            # Skip days with no work hours
            if daily_hours == 0:
                current_date += datetime.timedelta(days=1)
                continue
                
            work_start = self.pacific_tz.localize(
                datetime.datetime.combine(current_date, datetime.time(start_hour, 0))
            )
            work_end = self.pacific_tz.localize(
                datetime.datetime.combine(current_date, datetime.time(start_hour + daily_hours, 0))
            )
            

            
            # Find overlap between event and work hours for this day
            event_start_today = max(start_pt, work_start)
            event_end_today = min(end_pt, work_end)
            
            if event_start_today < event_end_today:
                overlap_seconds = (event_end_today - event_start_today).total_seconds()
                daily_overlap = overlap_seconds / 3600.0
                total_overlap += daily_overlap
            # No debug output
            else:
                pass
            
            current_date += datetime.timedelta(days=1)
        
        return max(0.0, total_overlap)

    def is_all_day_holiday(self, event: dict) -> bool:
        """Check if an event is an all-day Holiday/Vacation event."""
        categories = event.get('categories', '')
        if not categories:
            return False
        
        # Check if Holiday/Vacation is in the categories
        category = categories.split(",")[0].strip()
        if category != "Holiday/Vacation":
            return False
        
        # Check if it's an all-day event
        start_time = event['start']
        end_time = event['end']
        
        # Check if the event starts at midnight and ends at midnight the next day
        # or spans multiple days at midnight boundaries
        start_is_midnight = start_time.hour == 0 and start_time.minute == 0 and start_time.second == 0
        end_is_midnight = end_time.hour == 0 and end_time.minute == 0 and end_time.second == 0
        
        # Duration should be at least 24 hours for an all-day event
        duration_hours = (end_time - start_time).total_seconds() / 3600.0
        is_full_day = duration_hours >= 24
        
        return start_is_midnight and end_is_midnight and is_full_day

    def _fetch_events_for_range(self, range_start: datetime.datetime, range_end: datetime.datetime) -> List:
        """Fetch calendar events for a specific date range. Internal helper method."""
        events = []
        
        try:
            # Get calendar items within date range
            items = self.calendar_folder.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True
            
            # Create filter for date range (include any event that overlaps the window)
            start_str = range_start.strftime('%m/%d/%Y %I:%M %p')
            end_str = range_end.strftime('%m/%d/%Y %I:%M %p')

            # Outlook filter: event overlaps range if it starts before range end AND ends after range start
            date_filter = (
                f"[Start] < '{end_str}' AND "
                f"[End]   > '{start_str}'"
            )
            filtered_items = items.Restrict(date_filter)
            
            # Process each calendar item
            for item in filtered_items:
                try:
                    # Skip cancelled events
                    if hasattr(item, 'IsCancelled') and item.IsCancelled:
                        continue
                        
                    # Skip events marked Free (BusyStatus = 0) *unless* the category is Holiday/Vacation
                    busy_status = getattr(item, 'BusyStatus', 2)
                    categories_str = getattr(item, 'Categories', '')
                    first_cat = categories_str.split(',')[0].strip() if categories_str else ''
                    if busy_status == 0 and first_cat != "Holiday/Vacation":
                        continue

                    if hasattr(item, 'Start') and hasattr(item, 'End'):
                        # Convert to Pacific Time first
                        start_time = item.Start
                        end_time = item.End
                        start_pt, end_pt = self.convert_to_pacific(start_time, end_time)
                        
                        # Calculate work hours overlap (9am-5pm Pacific Time only)
                        work_hours_duration = self.calculate_work_hours_overlap(start_pt, end_pt)
                        
                        # Decide whether to include this event:
                        #  1) Include if it overlaps work hours (>0)
                        #  2) Always include all-day Holiday/Vacation events so they can be processed later
                        include_event = work_hours_duration > 0 or first_cat == "Holiday/Vacation"
                        if include_event:
                            event_info = {
                                'subject': getattr(item, 'Subject', 'Untitled'),
                                'start': start_pt,
                                'end': end_pt,
                                'duration_hours': work_hours_duration,
                                'categories': categories_str
                            }
                            events.append(event_info)
                except (AttributeError, TypeError, ValueError) as e:
                    print(f"Error processing calendar item: {e}")
                    continue
                    
        except (AttributeError, TypeError, ValueError) as e:
            print(f"Error retrieving calendar events: {e}")
        
        return events

    def get_current_workweek_events(self, week_offset: int = 0, start_date: datetime.date = None, end_date: datetime.date = None) -> List:
        """Get all calendar events for a workweek offset relative to the current week or custom date range.
        week_offset = 0 (default) -> current workweek
        week_offset = -1 -> previous workweek
        week_offset = 1 -> next workweek
        start_date: If provided, use this date as the starting date for the analysis
        end_date: If provided with start_date, use this as the ending date (inclusive)
        """
        # Calculate target date range bounds
        if start_date is not None:
            # Use provided start date
            range_start = datetime.datetime.combine(start_date, datetime.time(0, 0, 0))
            if end_date is not None:
                # Use provided end date, set to end of day (23:59:59)
                range_end = datetime.datetime.combine(end_date, datetime.time(23, 59, 59))
            else:
                # If only start_date provided, find the Monday of that week and use Friday
                days_since_monday = start_date.weekday()
                monday_start_date = start_date - datetime.timedelta(days=days_since_monday)
                range_start = datetime.datetime.combine(monday_start_date, datetime.time(0, 0, 0))
                range_end = range_start + datetime.timedelta(days=4, hours=23, minutes=59, seconds=59)
        else:
            now = datetime.datetime.now()
            # Monday of the current week (00:00)
            monday_current = now - datetime.timedelta(days=now.weekday())
            monday_current = monday_current.replace(hour=0, minute=0, second=0, microsecond=0)
            # Shift to the desired week
            range_start = monday_current + datetime.timedelta(days=7 * week_offset)
            range_end = range_start + datetime.timedelta(days=4, hours=23, minutes=59, seconds=59)
        
        # Calculate the total days in the range
        total_days = (range_end.date() - range_start.date()).days + 1
        
        # For large date ranges, process in monthly chunks to avoid Outlook COM slowdown
        if total_days > 31:
            print(f"Processing {total_days} days in monthly chunks...")
            events = []
            seen_events = set()  # Track unique events by (subject, start, end)
            chunk_start = range_start
            
            while chunk_start < range_end:
                # Calculate chunk end (up to 31 days or until range_end)
                chunk_end = min(
                    chunk_start + datetime.timedelta(days=30, hours=23, minutes=59, seconds=59),
                    range_end
                )
                
                # Show progress
                print(f"  Fetching: {chunk_start.strftime('%Y-%m-%d')} to {chunk_end.strftime('%Y-%m-%d')}...")
                
                # Fetch events for this chunk
                chunk_events = self._fetch_events_for_range(chunk_start, chunk_end)
                
                # Deduplicate events that may span chunk boundaries
                for event in chunk_events:
                    event_key = (event['subject'], event['start'], event['end'])
                    if event_key not in seen_events:
                        seen_events.add(event_key)
                        events.append(event)
                
                # Move to next chunk
                chunk_start = chunk_end.replace(hour=0, minute=0, second=0, microsecond=0) + datetime.timedelta(days=1)
            
            return events
        else:
            # For small date ranges, fetch directly
            return self._fetch_events_for_range(range_start, range_end)


def categorize_event(categories_str: str) -> str:
    """Determine the category for an event based on its categories."""
    if not categories_str or not categories_str.strip():
        return "Work Meeting"
    
    # Take the first category if multiple exist
    category = categories_str.split(",")[0].strip()
    
    if category not in KNOWN_CATEGORIES:
        return "Work Meeting"
    
    return category


def build_report(events: List[dict], reporter: OutlookCalendarReporter, debug: bool = False,
                  monday_start: datetime.datetime = None, friday_end: datetime.datetime = None) -> None:
    """Generate the weekly calendar usage report."""
    if not events:
        print("No calendar events found for the current workweek.")
        return
    
    # Initialize reporter instance to access holiday checking method
    # reporter = OutlookCalendarReporter()
    
    # Track all-day holidays and their dates
    holiday_dates = set()
    holiday_hours_reduction = 0.0
    
    # Find all-day holidays and calculate work hours reduction
    for event in events:
        if reporter.is_all_day_holiday(event):
            start_date = event['start'].date()
            end_date = event['end'].date()
            
            # Add all dates covered by the holiday
            current_date = start_date
            while current_date < end_date:  # End date is exclusive for all-day events
                holiday_dates.add(current_date)
                
                # Reduce work hours for this day
                weekday = current_date.weekday()
                daily_hours = DAILY_WORK_HOURS.get(weekday, 0)
                holiday_hours_reduction += daily_hours
                
                current_date += datetime.timedelta(days=1)
    
    # Filter out meetings that occur during holidays
    non_holiday_events = []
    for event in events:
        event_date = event['start'].date()
        # Keep the event if it doesn't fall on a holiday date
        if event_date not in holiday_dates:
            non_holiday_events.append(event)
    
    durations: Dict[str, float] = defaultdict(float)
    
    # Categorize and sum durations (raw, may overlap) - only for non-holiday events
    for event in non_holiday_events:
        duration = event.get('duration_hours', 0.0)
        category = categorize_event(event.get('categories', ''))
        durations[category] += duration
    
    # Add Holiday/Vacation category work-hours reduction
    has_holiday = holiday_hours_reduction > 0
    if has_holiday:
        holiday_hours = holiday_hours_reduction
    
    total_meeting_time_raw = sum(durations.values())

    # --- Busy/Free calculation using union of intervals ---
    pacific = pytz.timezone(TIMEZONE)
    intervals_by_day: Dict[datetime.date, List[tuple]] = defaultdict(list)

    # Only process non-holiday events for busy time calculation
    for ev in non_holiday_events:
        start_pt: datetime.datetime = ev['start']
        end_pt: datetime.datetime = ev['end']

        current_date = start_pt.date()
        last_date = end_pt.date()

        while current_date <= last_date:
            # Skip if this date is a holiday
            if current_date in holiday_dates:
                current_date += datetime.timedelta(days=1)
                continue
                
            # Get work hours for this day from configuration
            weekday = current_date.weekday()
            daily_hours = DAILY_WORK_HOURS.get(weekday, 0)
            start_hour = DAILY_START_TIMES.get(weekday, 9)
            
            # Skip days with no work hours
            if daily_hours > 0:
                work_start = pacific.localize(datetime.datetime.combine(current_date, datetime.time(start_hour, 0)))
                work_end   = pacific.localize(datetime.datetime.combine(current_date, datetime.time(start_hour + daily_hours, 0)))

                interval_start = max(start_pt, work_start)
                interval_end   = min(end_pt,   work_end)

                if interval_start < interval_end:
                    intervals_by_day[current_date].append((interval_start, interval_end))
            current_date += datetime.timedelta(days=1)

    busy_hours_total = 0.0
    for day_intervals in intervals_by_day.values():
        # merge intervals
        day_intervals.sort(key=lambda iv: iv[0])
        merged: List[tuple] = []
        for iv in day_intervals:
            if not merged or iv[0] > merged[-1][1]:
                merged.append(list(iv))
            else:
                merged[-1][1] = max(merged[-1][1], iv[1])
        # sum durations
        for iv in merged:
            busy_hours_total += (iv[1] - iv[0]).total_seconds() / 3600.0

    # Calculate adjusted total work hours (subtract holiday hours)
    adjusted_total_work_hours = TOTAL_WORK_HOURS - holiday_hours_reduction
    free_time = max(0.0, adjusted_total_work_hours - busy_hours_total)
    # Build date range string for report headers
    if monday_start and friday_end:
        date_range_str = f"{monday_start.strftime('%Y-%m-%d')} to {friday_end.strftime('%Y-%m-%d')}"
    else:
        date_range_str = ""

    # Define target ranges for each category
    RANGES = {
        "Work Meeting": "<= 12 hours",
        "Focus Time": "12-15 hours", 
        "Collaboration": "4-8 hours",
        "Communication": "<= 8 hours",
        "Free time": "5-25 hours",
        "Unavailable": "<= 6 hours",
    }
    
    # Sort categories with Holiday/Vacation at the bottom (after Unavailable)
    def sort_categories(item):
        cat, _ = item
        if cat == "Holiday/Vacation":
            return ("zz", cat)  # Force Holiday/Vacation to sort very last
        elif cat == "Unavailable":
            return ("z", cat)  # Unavailable just before Holiday/Vacation
        return ("a", cat)
    
    # Print complete markdown version first
    print("\n===== Markdown Format =====\n")
    print("#### Weekly Calendar Usage Report" + (" (debug)" if debug else ""))
    print()
    if date_range_str:
        print(f"Date range: {date_range_str}")
    print(f"Events analyzed: {len(events)}")
    print(f"Total planned meeting time (raw): {total_meeting_time_raw:.2f} h")
    print(f"Busy time (union): {busy_hours_total:.2f} h")
    if holiday_hours_reduction > 0:
        print(f"Work hours reduced by holidays: {holiday_hours_reduction:.2f} h")
        print(f"Adjusted total work hours: {adjusted_total_work_hours:.2f} h")
    print(f"Free time remaining: {free_time:.2f} h")
    print()
    print("| Category | Range | Hours | Remaining | Warning |")
    print("|----------|-------|-------|-----------|---------|")
    
    # Print category rows in markdown (sorted with Holiday/Vacation and Unavailable at bottom)
    for cat, hrs in sorted(durations.items(), key=sort_categories):
        budget = BUDGETS.get(cat)
        
        # No Holiday/Vacation in main loop
        
        remaining = budget["max"] - hrs if budget else ""
        warn_msg = ""
        if budget:
            if hrs > budget["max"]:
                warn_msg = "Exceeded"
            elif hrs > budget["warn"]:
                warn_msg = "Warning"
            elif hrs < budget["min"]:
                warn_msg = "Below min"
        range_str = RANGES.get(cat, "")
        remaining_str = f"{remaining:.2f}" if remaining != "" else ""
        print(f"| {cat} | {range_str} | {hrs:.2f} | {remaining_str} | {warn_msg} |")
    
    # Show categories with zero hours in markdown (sorted with Unavailable at bottom)
    zero_categories = [(cat, 0.0) for cat in BUDGETS if cat not in durations]
    for cat, hrs in sorted(zero_categories, key=sort_categories):
        status = "Below min" if BUDGETS[cat]["min"] else ""
        range_str = RANGES.get(cat, "")
        print(f"| {cat} | {range_str} | 0.00 | {BUDGETS[cat]['max']:.2f} | {status} |")
    
    # Append Holiday/Vacation row at bottom if present
    if has_holiday:
        print(f"| Holiday/Vacation |  | {holiday_hours:.2f} |  |  |")
    
    # Show free time in markdown
    range_str = RANGES.get("Free time", "")
    print(f"| Free time | {range_str} | {free_time:.2f} |  |  |")
    
    # Print ASCII report header
    print("\n\n===== Weekly Calendar Usage Report =====\n")
    if date_range_str:
        print(f"Date range: {date_range_str}")
    print(f"Events analyzed: {len(events)}")
    print(f"Total planned meeting time (raw): {total_meeting_time_raw:.2f} h")
    print(f"Busy time (union): {busy_hours_total:.2f} h")
    if holiday_hours_reduction > 0:
        print(f"Work hours reduced by holidays: {holiday_hours_reduction:.2f} h")
        print(f"Adjusted total work hours: {adjusted_total_work_hours:.2f} h")
    print(f"Free time remaining:       {free_time:.2f} h\n")
    
    # Print category breakdown
    header = f"{'Category':<15}{'Range':>15}{'Hours':>10}{'Remaining':>12}{'Warning':>12}"
    print(header)
    print("-" * len(header))
    
    for cat, hrs in sorted(durations.items(), key=sort_categories):
        budget = BUDGETS.get(cat)
        
        # No Holiday/Vacation in main loop
        
        remaining = budget["max"] - hrs if budget else 0.0
        warn_msg = ""
        if budget:
            if hrs > budget["max"]:
                warn_msg = "Exceeded"
            elif hrs > budget["warn"]:
                warn_msg = "Warning"
            elif hrs < budget["min"]:
                warn_msg = "Below min"
        range_str = RANGES.get(cat, "")
        remaining_str = f"{remaining:.2f}" if budget else ""
        print(f"{cat:<15}{range_str:>15}{hrs:>10.2f}{remaining_str:>12}{warn_msg:>12}")
    
    # Show categories with zero hours but have budgets (sorted with Unavailable at bottom)
    for cat, hrs in sorted(zero_categories, key=sort_categories):
        status = "Below min" if BUDGETS[cat]["min"] else ""
        range_str = RANGES.get(cat, "")
        print(f"{cat:<15}{range_str:>15}{0.0:>10.2f}{BUDGETS[cat]['max']:>12.2f}{status:>12}")
    
    # Append Holiday/Vacation row at bottom (ASCII)
    if has_holiday:
        print(f"{'Holiday/Vacation':<15}{'':>15}{holiday_hours:>10.2f}{'':>12}{'':>12}")
    
    # Show free time at the end
    range_str = RANGES.get("Free time", "")
    print(f"{'Free time':<15}{range_str:>15}{free_time:>10.2f}{'':>12}{'':>12}")

    # ---------- Debug section ----------
    if debug:
        print("\n===== DEBUG: All calendar items (processed) =====")
        events_sorted = sorted(events, key=lambda x: x['start'])
        prev_date = None
        for ev in events_sorted:
            cat = categorize_event(ev.get('categories', ''))
            s_dt = ev['start']
            e_dt = ev['end']
            current_date = s_dt.date()
            date_str = s_dt.strftime('%Y-%m-%d')
            duration_str = f"{ev['duration_hours']:.2f}h"
            time_str = f"{s_dt.strftime('%H:%M')}-{e_dt.strftime('%H:%M')}"
            holiday_marker = " [ALL-DAY HOLIDAY/VACATION]" if reporter.is_all_day_holiday(ev) else ""
            filtered_marker = " [FILTERED]" if ev['start'].date() in holiday_dates and not reporter.is_all_day_holiday(ev) else ""
            if prev_date is not None and current_date != prev_date:
                print()
            print(f"{date_str} {duration_str:>6}  [{cat:<13}] {time_str} {ev['subject']}{holiday_marker}{filtered_marker}")
            prev_date = current_date


def get_workweek_bounds(week_offset: int = 0, start_date: datetime.date = None, end_date: datetime.date = None):
    """Return the start and end datetimes for the specified workweek offset or custom date range."""
    if start_date is not None:
        # Use provided start date
        monday_start = datetime.datetime.combine(start_date, datetime.time(0, 0, 0))
        if end_date is not None:
            # Use provided end date
            friday_end = datetime.datetime.combine(end_date, datetime.time(23, 59, 59))
        else:
            # If only start_date provided, find the Monday of that week and use Friday
            days_since_monday = start_date.weekday()
            monday_start_date = start_date - datetime.timedelta(days=days_since_monday)
            monday_start = datetime.datetime.combine(monday_start_date, datetime.time(0, 0, 0))
            friday_end = monday_start + datetime.timedelta(days=4, hours=23, minutes=59, seconds=59)
    else:
        now = datetime.datetime.now()
        monday_current = now - datetime.timedelta(days=now.weekday())
        monday_current = monday_current.replace(hour=0, minute=0, second=0, microsecond=0)
        monday_start = monday_current + datetime.timedelta(days=7 * week_offset)
        friday_end = monday_start + datetime.timedelta(days=4, hours=23, minutes=59, seconds=59)
    return monday_start, friday_end


def is_admin():
    """Return True only if the current process is running elevated."""
    import ctypes.wintypes as wt
    import ctypes
    advapi32 = ctypes.windll.advapi32
    token = wt.HANDLE()
    TOKEN_QUERY = 0x0008
    TOKEN_ELEVATION = 20
    class TOKEN_ELEVATION_STRUCT(ctypes.Structure):
        _fields_ = [("TokenIsElevated", wt.DWORD)]
    if not advapi32.OpenProcessToken(ctypes.windll.kernel32.GetCurrentProcess(),
                                     TOKEN_QUERY, ctypes.byref(token)):
        return False
    te = TOKEN_ELEVATION_STRUCT()
    te_size = wt.DWORD(ctypes.sizeof(te))
    advapi32.GetTokenInformation(token, TOKEN_ELEVATION,
                                 ctypes.byref(te), te_size, ctypes.byref(te_size))
    return bool(te.TokenIsElevated)


def is_outlook_running():
    """Check if Outlook is currently running."""
    try:
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq OUTLOOK.EXE'], 
                              capture_output=True, text=True, timeout=10, check=False)
        return 'OUTLOOK.EXE' in result.stdout
    except (subprocess.TimeoutExpired, OSError):
        return False


def kill_outlook():
    """Kill Outlook process if running."""
    try:
        print("Outlook is running. Terminating Outlook process...")
        subprocess.run(['taskkill', '/F', '/IM', 'OUTLOOK.EXE'], 
                      capture_output=True, timeout=10, check=False)
        print("Waiting 10 seconds for process cleanup...")
        time.sleep(10)
        return True
    except (subprocess.TimeoutExpired, OSError) as e:
        print(f"Warning: Failed to terminate Outlook: {e}")
        return False


def main():
    """Main function to generate the calendar report."""
    # Check if running as Administrator and exit if so
    if is_admin():
        print("Error: This script should not be run as Administrator.")
        print("Please run this script in a normal user context (without 'Run as Administrator').")
        print("Outlook COM automation works best when running with the same privileges as Outlook.")
        sys.exit(1)
    
    # Check if Outlook is running and kill it if necessary
    if is_outlook_running():
        if not kill_outlook():
            print("Warning: Could not terminate Outlook. COM automation may fail.")
            print("Consider manually closing Outlook and running the script again.")
    
    try:
        # Initialize the calendar reporter
        reporter = OutlookCalendarReporter()
        
        print("Connecting to Outlook calendar...")
        
        # Get date range and events based on requested parameters
        monday_start, friday_end = get_workweek_bounds(WEEK_OFFSET, START_DATE, END_DATE)
        events = reporter.get_current_workweek_events(week_offset=WEEK_OFFSET, start_date=START_DATE, end_date=END_DATE)
 
        # Generate and display report
        build_report(events, reporter, DEBUG_MODE, monday_start, friday_end)
        
    except ConnectionError as e:
        print(f"Connection Error: {e}")
        print("Make sure Outlook is installed and accessible.")
    except KeyboardInterrupt:
        print("\nReport generation cancelled by user.")
    except (OSError, ValueError, RuntimeError) as e:
        print(f"An unexpected error occurred: {e}")


if __name__ == "__main__":
    main() 