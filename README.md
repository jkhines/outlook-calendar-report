# Outlook Calendar Report

This script queries your Outlook calendar for meetings in the current workweek and produces a report showing time spent in each category versus your budgets. The script uses allocation recommendations from Atlassian's excellent "Redesign your workweek" training at [https://community.atlassian.com/learning/course/redesign-your-workweek](https://community.atlassian.com/learning/course/redesign-your-workweek).

## Setup

This project uses [uv](https://github.com/astral-sh/uv) for dependency management.

### Installing uv

On Windows (PowerShell):
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Or using pip:
```powershell
pip install uv
```

### Installing Dependencies

```powershell
uv sync
```

This will create a virtual environment and install all dependencies.

### Running the Script

```powershell
uv run python calendar_report.py
```

Or activate the virtual environment first:
```powershell
uv venv
.\.venv\Scripts\Activate.ps1
python calendar_report.py
```

## Usage

```powershell
uv run python calendar_report.py [options]
```

Or activate the virtual environment first:
```powershell
uv venv
.\.venv\Scripts\Activate.ps1
python calendar_report.py [options]
```

Options:
- `--lastweek`   Analyze previous workweek (Mon-Fri)
- `--nextweek`   Analyze next workweek
- `--verbose`    Show every meeting in the output
- `--help, -h`   Show help message and exit

## Configuration

The script can be customized by modifying constants at the top of `calendar_report.py`:

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
  - Default: Focus Time, Communication, Unavailable, Collaboration, Holiday

To customize these settings, edit the constants in `calendar_report.py` before running the script.

