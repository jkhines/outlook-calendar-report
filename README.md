# Calendar Report

This project produces a weekly report showing time spent in each calendar category versus your budgets. The script uses allocation recommendations from Atlassian's excellent "Redesign your workweek" training at [https://community.atlassian.com/learning/course/redesign-your-workweek](https://community.atlassian.com/learning/course/redesign-your-workweek).

Two platform-specific implementations are provided:

| Directory | Platform | Data Source |
|-----------|----------|-------------|
| `win32/`  | Windows  | Outlook COM (win32com) |
| `linux/`  | Linux    | Obsidian daily notes |

Each subdirectory is self-contained with its own `pyproject.toml` and dependencies.

## Setup

This project uses [uv](https://github.com/astral-sh/uv) for dependency management.

### Installing uv

On Linux:
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

On Windows (PowerShell):
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### Installing Dependencies

```bash
cd win32/  # or linux/
uv sync
```

## Windows (win32/)

Queries Outlook calendar events via the COM interface. Requires Outlook to be installed.

### Running

```powershell
cd win32
uv run python calendar_report.py [options]
```

### Options

- `--lastweek`   Analyze previous workweek (Mon-Fri)
- `--nextweek`   Analyze next workweek
- `--start DATE` Analyze starting from DATE (format: yyyy-MM-dd)
- `--end DATE`   Analyze ending at DATE (format: yyyy-MM-dd)
- `--verbose`    Show every meeting in the output
- `--help, -h`   Show help message and exit

## Linux (linux/)

Reads Obsidian daily notes (Markdown files named `yyyy-MM-dd.md`) and extracts meetings from the `### Meetings` section.

### Running

```bash
cd linux
uv run python calendar_report.py [options]
```

### Options

Same as Windows, plus:

- `--path PATH`  Obsidian daily notes directory (default: `~/obsidian/Daily`)

### Obsidian Note Format

The script reads meetings from the `### Meetings` section of each daily note. Each meeting is a checklist item with a time range:

```markdown
### Meetings
- [x] 09:00-10:00 Sprint Planning
- [ ] 10:30-11:00 1:1 with Alice #Collaboration
- [x] 14:00-15:00 Focus block #FocusTime
- [x] 16:30-17:00 Communication #Communication
```

- Lines must match the pattern `- [x] HH:MM-HH:MM Subject` or `- [ ] HH:MM-HH:MM Subject`
- Both checked (`[x]`) and unchecked (`[ ]`) meetings are included
- Only lines between `### Meetings` and the next `###` heading are parsed

### Category Tags

Assign categories using inline `#Tags` at the end of the meeting subject:

| Tag | Category |
|-----|----------|
| `#FocusTime` | Focus Time |
| `#Communication` | Communication |
| `#Collaboration` | Collaboration |
| `#Unavailable` | Unavailable |
| `#WorkMeeting` | Work Meeting |
| `#HolidayVacation` | Holiday/Vacation |

Tags are case-insensitive. Untagged meetings default to "Work Meeting".

### Holiday/Vacation

To mark a full-day holiday, use a `00:00-00:00` entry with the `#HolidayVacation` tag:

```markdown
- [x] 00:00-00:00 Holiday #HolidayVacation
```

Or use a meeting spanning the full work day (e.g., `08:00-17:00`) with the same tag.

## Configuration

Both scripts can be customized by modifying constants at the top of their respective `calendar_report.py`:

### Work Schedule

- **`DAILY_WORK_HOURS`**: Hours worked per day (Monday=0, Sunday=6)
  - Default: Monday-Wednesday 8 hours, Thursday-Friday 9 hours, weekends 0 hours

- **`DAILY_START_TIMES`**: Work start hour (24-hour format) for each day
  - Default: Monday-Wednesday 9am, Thursday-Friday 8am

### Timezone

- **`TIMEZONE`**: Timezone string for calendar event processing
  - Default: `"US/Pacific"`
  - Use any valid pytz timezone name (e.g., `"US/Eastern"`, `"Europe/London"`)

### Category Budgets

- **`BUDGETS`**: Target time ranges (in hours) for each calendar category
  - Each category has `min`, `max`, and `warn` thresholds
  - Default categories: Focus Time, Collaboration, Communication, Work Meeting, Unavailable

- **`KNOWN_CATEGORIES`**: Set of recognized calendar categories
  - Events with unrecognized categories default to "Work Meeting"
  - Default: Focus Time, Communication, Unavailable, Collaboration, Holiday/Vacation
