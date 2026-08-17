# Personal Attendance Tracker

## Run locally

Open this folder in VS Code, activate your virtual environment, then:

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app reads `timetable.xlsx` automatically.

### Current features
- Dynamically discovers electives from the Excel timetable
- Treats Sec A/B/C as separate selectable electives
- Builds your personal timetable from your selections
- Shows today's selected classes
- Mark each class Present / OD / Skip
- Stores attendance in SQLite
- Subject and overall attendance
- Class-by-class history
- Exports your selected timetable as an `.ics` calendar file

The Excel file is the source of truth for dates, times, subjects and faculty.
