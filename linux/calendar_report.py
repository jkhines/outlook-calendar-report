#!/usr/bin/env python3
"""
Calendar Category Usage Report (Linux / Obsidian)

This script reads Obsidian daily notes for the current workweek and produces a
report showing time spent in each category versus your budgets. Categories are
assigned via inline #Tags on meeting lines.
"""

import datetime
import os
import re
import sys
import pytz
from collections import defaultdict
from typing import Dict, List, Optional

# ---------- Global Variables and Configuration ----------
DEBUG_MODE = "--verbose" in sys.argv

# Display help/usage information
if "--help" in sys.argv or "-h" in sys.argv:
    print(
        "Calendar Category Usage Report (Obsidian)\n\n"
        "Usage: python calendar_report.py [options]\n\n"
        "Options:\n"
        "  --lastweek   Analyze previous workweek (Mon-Fri)\n"
        "  --nextweek   Analyze next workweek\n"
        "  --start DATE Analyze starting from DATE (format: yyyy-MM-dd)\n"
        "  --end DATE   Analyze ending at DATE (format: yyyy-MM-dd)\n"
        "  --path PATH  Obsidian daily notes directory (default: ~/obsidian/Daily)\n"
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

# Parse --path parameter
OBSIDIAN_PATH = os.path.expanduser("~/obsidian/Daily")
if "--path" in sys.argv:
    path_idx = sys.argv.index("--path")
    if path_idx + 1 >= len(sys.argv):
        print("Error: --path requires a directory argument")
        sys.exit(1)
    OBSIDIAN_PATH = os.path.expanduser(sys.argv[path_idx + 1])

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

# Mapping from inline #Tags to category names (case-insensitive lookup)
TAG_TO_CATEGORY = {
    "focustime": "Focus Time",
    "communication": "Communication",
    "collaboration": "Collaboration",
    "unavailable": "Unavailable",
    "workmeeting": "Work Meeting",
    "holidayvacation": "Holiday/Vacation",
}

# Regex for meeting lines: - [x] or - [ ] followed by HH:MM-HH:MM subject
MEETING_RE = re.compile(r"^- \[[ x]\] (\d{2}:\d{2})-(\d{2}:\d{2}) (.+)$")

# Regex for inline tag at end of subject
TAG_RE = re.compile(r"\s+#(\S+)\s*$")


class ObsidianCalendarReporter:
    """Class to handle calendar reporting from Obsidian daily notes."""

    def __init__(self, notes_path: str):
        """Initialize with the path to Obsidian daily notes directory."""
        self.notes_path = notes_path
        self.pacific_tz = pytz.timezone(TIMEZONE)

    def _parse_tag(self, subject: str) -> tuple:
        """Extract category from inline #Tag at end of subject.
        Returns (cleaned_subject, category)."""
        match = TAG_RE.search(subject)
        if not match:
            return subject, "Work Meeting"

        tag_name = match.group(1).lower()
        cleaned_subject = subject[:match.start()].rstrip()
        category = TAG_TO_CATEGORY.get(tag_name)
        if category is None:
            # Convert CamelCase to spaced: insert space before uppercase letters
            spaced = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", match.group(1))
            category = spaced
        return cleaned_subject, category

    def _parse_note(self, file_path: str, note_date: datetime.date) -> List[dict]:
        """Parse a single Obsidian daily note and extract meeting events."""
        events = []
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
        except (OSError, IOError):
            return events

        in_meetings_section = False
        for line in lines:
            stripped = line.rstrip("\n")

            # Detect section boundaries
            if stripped.startswith("### Meetings"):
                in_meetings_section = True
                continue
            if stripped.startswith("###") and in_meetings_section:
                in_meetings_section = False
                continue
            if not in_meetings_section:
                continue

            match = MEETING_RE.match(stripped)
            if not match:
                continue

            start_str, end_str, subject = match.group(1), match.group(2), match.group(3)
            start_hour, start_min = int(start_str[:2]), int(start_str[3:])
            end_hour, end_min = int(end_str[:2]), int(end_str[3:])

            start_time = self.pacific_tz.localize(
                datetime.datetime.combine(note_date, datetime.time(start_hour, start_min))
            )
            end_time = self.pacific_tz.localize(
                datetime.datetime.combine(note_date, datetime.time(end_hour, end_min))
            )

            # Handle midnight-to-midnight (00:00-00:00) as a full-day marker
            if start_hour == 0 and start_min == 0 and end_hour == 0 and end_min == 0:
                end_time += datetime.timedelta(days=1)

            cleaned_subject, category = self._parse_tag(subject)

            events.append({
                "subject": cleaned_subject,
                "start": start_time,
                "end": end_time,
                "categories": category,
            })

        return events

    def get_current_workweek_events(self, week_offset: int = 0,
                                     start_date: Optional[datetime.date] = None,
                                     end_date: Optional[datetime.date] = None) -> List[dict]:
        """Get all meeting events for the specified date range from Obsidian notes."""
        if start_date is not None:
            range_start = start_date
            if end_date is not None:
                range_end = end_date
            else:
                days_since_monday = start_date.weekday()
                monday = start_date - datetime.timedelta(days=days_since_monday)
                range_end = monday + datetime.timedelta(days=4)
        else:
            now = datetime.datetime.now()
            monday = now - datetime.timedelta(days=now.weekday())
            monday = monday.replace(hour=0, minute=0, second=0, microsecond=0)
            monday += datetime.timedelta(days=7 * week_offset)
            range_start = monday.date()
            range_end = range_start + datetime.timedelta(days=4)

        events = []
        current_date = range_start
        while current_date <= range_end:
            file_name = current_date.strftime("%Y-%m-%d") + ".md"
            file_path = os.path.join(self.notes_path, file_name)
            if os.path.isfile(file_path):
                day_events = self._parse_note(file_path, current_date)
                events.extend(day_events)
            current_date += datetime.timedelta(days=1)

        # Calculate work hours overlap for each event
        for event in events:
            event["duration_hours"] = calculate_work_hours_overlap(
                event["start"], event["end"], self.pacific_tz
            )

        return events

    def is_all_day_holiday(self, event: dict) -> bool:
        """Check if an event is an all-day Holiday/Vacation event."""
        category = event.get("categories", "")
        if category != "Holiday/Vacation":
            return False

        start_time = event["start"]
        end_time = event["end"]

        # Check midnight-to-midnight boundaries (00:00-00:00 spanning >= 24h)
        start_is_midnight = start_time.hour == 0 and start_time.minute == 0 and start_time.second == 0
        end_is_midnight = end_time.hour == 0 and end_time.minute == 0 and end_time.second == 0
        duration_hours = (end_time - start_time).total_seconds() / 3600.0
        is_full_day = duration_hours >= 24

        if start_is_midnight and end_is_midnight and is_full_day:
            return True

        # Also treat as holiday if it spans the full configured work day
        weekday = start_time.date().weekday()
        daily_hours = DAILY_WORK_HOURS.get(weekday, 0)
        start_hour = DAILY_START_TIMES.get(weekday, 9)
        if daily_hours > 0:
            work_start = self.pacific_tz.localize(
                datetime.datetime.combine(start_time.date(), datetime.time(start_hour, 0))
            )
            work_end = self.pacific_tz.localize(
                datetime.datetime.combine(start_time.date(), datetime.time(start_hour + daily_hours, 0))
            )
            if start_time <= work_start and end_time >= work_end:
                return True

        return False


def calculate_work_hours_overlap(start_time: datetime.datetime, end_time: datetime.datetime,
                                  pacific_tz: pytz.BaseTzInfo) -> float:
    """Calculate how many hours of an event overlap with work hours.
    Assumes start_time and end_time are already in Pacific Time."""
    total_overlap = 0.0

    current_date = start_time.date()
    end_date = end_time.date()

    while current_date <= end_date:
        weekday = current_date.weekday()
        daily_hours = DAILY_WORK_HOURS.get(weekday, 0)
        start_hour = DAILY_START_TIMES.get(weekday, 9)

        if daily_hours == 0:
            current_date += datetime.timedelta(days=1)
            continue

        work_start = pacific_tz.localize(
            datetime.datetime.combine(current_date, datetime.time(start_hour, 0))
        )
        work_end = pacific_tz.localize(
            datetime.datetime.combine(current_date, datetime.time(start_hour + daily_hours, 0))
        )

        event_start_today = max(start_time, work_start)
        event_end_today = min(end_time, work_end)

        if event_start_today < event_end_today:
            overlap_seconds = (event_end_today - event_start_today).total_seconds()
            total_overlap += overlap_seconds / 3600.0

        current_date += datetime.timedelta(days=1)

    return max(0.0, total_overlap)


def categorize_event(categories_str: str) -> str:
    """Determine the category for an event based on its categories."""
    if not categories_str or not categories_str.strip():
        return "Work Meeting"

    # Take the first category if multiple exist
    category = categories_str.split(",")[0].strip()

    if category not in KNOWN_CATEGORIES and category not in BUDGETS:
        return "Work Meeting"

    return category


def build_report(events: List[dict], reporter: ObsidianCalendarReporter, debug: bool = False,
                  monday_start: datetime.datetime = None, friday_end: datetime.datetime = None) -> None:
    """Generate the weekly calendar usage report."""
    if not events:
        print("No calendar events found for the current workweek.")
        return

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


def main():
    """Main function to generate the calendar report."""
    if not os.path.isdir(OBSIDIAN_PATH):
        print(f"Error: Obsidian daily notes directory not found: {OBSIDIAN_PATH}")
        print("Use --path to specify the correct directory.")
        sys.exit(1)

    try:
        reporter = ObsidianCalendarReporter(OBSIDIAN_PATH)

        print(f"Reading Obsidian daily notes from {OBSIDIAN_PATH}...")

        monday_start, friday_end = get_workweek_bounds(WEEK_OFFSET, START_DATE, END_DATE)
        events = reporter.get_current_workweek_events(week_offset=WEEK_OFFSET, start_date=START_DATE, end_date=END_DATE)

        build_report(events, reporter, DEBUG_MODE, monday_start, friday_end)

    except KeyboardInterrupt:
        print("\nReport generation cancelled by user.")
    except (OSError, ValueError, RuntimeError) as e:
        print(f"An unexpected error occurred: {e}")


if __name__ == "__main__":
    main()
