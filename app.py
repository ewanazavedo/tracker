
import re
import sqlite3
from datetime import datetime, date, time
from zoneinfo import ZoneInfo
from pathlib import Path
import pdfplumber
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from icalendar import Calendar, Event

APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "attendance.db"
TIMETABLE_PATH = APP_DIR / "timetable.xlsx"

st.set_page_config(
    page_title="Tracker",
    page_icon="📚",
    layout="wide",
)
# -----------------------------
# Simple user login
# -----------------------------

if "user_id" not in st.query_params:
    st.title("Welcome to RollCalls")

    user_name = st.text_input(
        "Enter your name / ID",
        placeholder="e.g. Ewan"
    ).strip()

    if st.button(
        "Continue",
        type="primary",
        disabled=not user_name
    ):
        st.query_params["user_id"] = user_name
        st.rerun()

    st.stop()

USER_ID = st.query_params["user_id"].strip().lower()
st.markdown("""
<style>
/* ---------- MESS MENU ---------- */

.menu-day {
    font-size: 20px;
    font-weight: 700;
    color: #a5b4fc;
    margin: 8px 0 20px 0;
}

.mess-card {
    width: 100%;
    box-sizing: border-box;

    background: rgba(255,255,255,0.045);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 16px;

    padding: 18px 20px;
    margin-bottom: 14px;
}

.mess-meal {
    font-size: 12px;
    font-weight: 700;
    color: #a5b4fc;

    text-transform: uppercase;
    letter-spacing: 1px;

    margin-bottom: 8px;
}

.mess-items {
    font-size: 15px;
    line-height: 1.6;
    color: #e2e8f0;
}
/* ---------- SUBJECT RINGS ---------- */

.subject-ring {
    width: 180px;
    height: 180px;

    position: relative;

    display: flex;
    align-items: center;
    justify-content: center;

    margin: 15px 0;
}

.subject-svg {
    position: absolute;

    top: 0;
    left: 0;

    width: 180px;
    height: 180px;

    transform: rotate(-90deg);

    z-index: 1;
}

.subject-bg {
    fill: none;
    stroke: rgba(255,255,255,0.10);
    stroke-width: 14;
}

.subject-progress {
    fill: none;

    stroke: #6366f1;
    stroke-width: 14;

    stroke-linecap: round;

    stroke-dasharray: 471.24;
    stroke-dashoffset: 471.24;

    animation: subject-ring-animation 1.5s ease-out forwards;
}

@keyframes subject-ring-animation {
    from {
        stroke-dashoffset: 471.24;
    }

    to {
        stroke-dashoffset: var(--subject-dash-offset);
    }
}

.subject-percentage {
    position: absolute;

    inset: 0;

    display: flex;
    align-items: center;
    justify-content: center;

    z-index: 5;

    font-size: 27px;
    font-weight: 800;

    text-align: center;
}

/* ---------- SUBJECT ATTENDANCE ---------- */

.subject-card {
    width: 100%;
    height: 390px;
    margin-bottom: 15px;
    box-sizing: border-box;

    padding: 22px 18px;

    border-radius: 18px;

    background: rgba(255, 255, 255, 0.045);
    border: 1px solid rgba(255, 255, 255, 0.08);

    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: space-between;

    text-align: center;
}

/* Standard spacing around every subject column */

div[data-testid="column"] {
    padding-left: 8px;
    padding-right: 8px;
    margin-bottom: 16px;
}

.subject-card:hover {
    transform: translateY(-4px);

    border-color: rgba(129, 140, 248, 0.25);

    box-shadow:
        0 15px 35px rgba(0, 0, 0, 0.18);
}

.subject-name {
    width: 100%;

    min-height: 52px;

    font-size: 16px;
    font-weight: 700;

    color: #f8fafc;

    display: flex;
    align-items: center;
    justify-content: center;

    line-height: 1.35;
}

.subject-ring {
    width: 180px;
    height: 180px;

    position: relative;

    display: flex;
    align-items: center;
    justify-content: center;

    margin: 15px 0;

    flex-shrink: 0;
}

.subject-attendance-svg {
    position: absolute;

    top: 0;
    left: 0;

    width: 180px;
    height: 180px;

    transform: rotate(-90deg);

    display: block;

    overflow: visible;

    z-index: 1;
}

.subject-ring-background {
    fill: none;

    stroke: rgba(255, 255, 255, 0.10);

    stroke-width: 14;
}

.subject-ring-progress {
    fill: none;

    stroke-width: 14;

    stroke-linecap: round;

    stroke-dasharray: 471.24;
    stroke-dashoffset: 471.24;

    animation:
        subject-draw-attendance 1.5s ease-out forwards;
}

@keyframes subject-draw-attendance {
    from {
        stroke-dashoffset: 471.24;
    }

    to {
        stroke-dashoffset: var(--subject-dash-offset);
    }
}

.subject-ring-value {
    position: absolute;
    inset: 0;

    display: flex;
    align-items: center;
    justify-content: center;

    z-index: 5;

    font-size: 27px;
    font-weight: 800;
}

.subject-stats {
    width: 100%;

    padding-top: 12px;

    border-top: 1px solid rgba(255,255,255,0.06);

    font-size: 12px;

    color: #94a3b8;

    line-height: 1.6;
}

.attendance-ring {
    width: 220px;
    height: 220px;
    margin: 20px auto 35px auto;

    position: relative;

    display: flex;
    align-items: center;
    justify-content: center;
}

.attendance-svg {
    position: absolute;
    inset: 0;

    width: 220px;
    height: 220px;

    transform: rotate(-90deg);
}

.ring-background {
    fill: none;
    stroke: rgba(255,255,255,0.08);
    stroke-width: 20;
}

.ring-progress {
    fill: none;
    stroke: #6366f1;
    stroke-width: 20;
    stroke-linecap: round;

    stroke-dasharray: 565.49;
    stroke-dashoffset: 565.49;

    animation: draw-attendance 1.8s ease-out forwards;
}

@keyframes draw-attendance {
    from {
        stroke-dashoffset: 565.49;
    }

    to {
        stroke-dashoffset: var(--dash-offset);
    }
}

.attendance-ring-value {
    position: absolute;
    inset: 0;

    display: flex;
    align-items: center;
    justify-content: center;

    z-index: 5;

    font-size: 34px;
    font-weight: 800;
    text-align: center;
}

.attendance-ring::before {
    content: "";
    position: absolute;

    width: 178px;
    height: 178px;

    border-radius: 50%;

    background: #111827;
}

.attendance-ring-value,
.attendance-ring-label {
    position: relative;
    z-index: 2;
}

.attendance-ring-value {
    font-size: 34px;
    font-weight: 750;
    color: #f8fafc;
    text-align: center;
}

.attendance-ring-label {
    font-size: 10px;
    letter-spacing: 1.5px;
    color: #94a3b8;
    text-align: center;
}

    /* ---------- GLOBAL ---------- */

.stApp {
    background:
        radial-gradient(
            circle at 15% 15%,
            rgba(99, 102, 241, 0.18),
            transparent 32%
        ),
        radial-gradient(
            circle at 85% 25%,
            rgba(56, 189, 248, 0.10),
            transparent 30%
        ),
        radial-gradient(
            circle at 50% 100%,
            rgba(139, 92, 246, 0.12),
            transparent 35%
        ),
        linear-gradient(
            135deg,
            #0b1020 0%,
            #111827 50%,
            #0b1220 100%
        );

    background-attachment: fixed;
    color: #f5f7fb;
}

    .block-container {
        max-width: 1200px;
        padding-top: 2.5rem;
        padding-bottom: 4rem;
    }

    /* ---------- HEADINGS ---------- */

    h1 {
        font-size: 2.6rem !important;
        font-weight: 700 !important;
        letter-spacing: -1px;
    }

    h2 {
        font-size: 1.8rem !important;
        font-weight: 650 !important;
    }

    h3 {
        font-size: 1.15rem !important;
        font-weight: 600 !important;
    }

  /* ---------- SIDEBAR ---------- */

section[data-testid="stSidebar"] {
    background:
        linear-gradient(
            160deg,
            #151D3A 0%,
            #111A34 45%,
            #0D162D 100%
        );

    border-right: 1px solid rgba(129, 140, 248, 0.18);

    box-shadow:
        8px 0 35px rgba(0, 0, 0, 0.20);
} {
    background:
        radial-gradient(
            circle at 20% 10%,
            rgba(99, 102, 241, 0.12),
            transparent 35%
        ),
        radial-gradient(
            circle at 90% 85%,
            rgba(56, 189, 248, 0.07),
            transparent 35%
        ),
        linear-gradient(
            180deg,
            #10172A 0%,
            #0D1424 55%,
            #0B1020 100%
        );

    border-right: 1px solid rgba(129, 140, 248, 0.10);
}

section[data-testid="stSidebar"] .block-container {
    padding: 1.5rem 1rem;
    overflow: hidden;
}

section[data-testid="stSidebar"] > div:first-child {
    overflow: hidden;
}

section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
    overflow: hidden;
}

/* Brand */

.sidebar-brand {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 10px 8px 25px 8px;
}

.brand-icon {
    width: 42px;
    height: 42px;

    display: flex;
    align-items: center;
    justify-content: center;

    border-radius: 12px;

    background: linear-gradient(
        135deg,
        #6366f1,
        #8b5cf6
    );

    color: white;
    font-size: 18px;
    font-weight: 800;

    box-shadow:
        0 8px 25px rgba(99,102,241,0.25);
}

.brand-title {
    font-size: 17px;
    font-weight: 700;
    color: #f8fafc;
}

.brand-subtitle {
    font-size: 11px;
    color: #7f8ba3;
    margin-top: 2px;
}

/* Navigation label */

.nav-label {
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 1.5px;
    color: #64748b;
    padding: 0 10px 8px 10px;
}

/* Navigation buttons */

section[data-testid="stSidebar"]
div[role="radiogroup"] {
    gap: 6px;
}

section[data-testid="stSidebar"]
div[role="radiogroup"] label {
    border-radius: 11px;
    padding: 10px 12px;
    margin: 0;
    border: 1px solid transparent;
    transition: all 0.2s ease;
}

section[data-testid="stSidebar"]
div[role="radiogroup"] label:hover {
    background: rgba(255,255,255,0.055);
    border-color: rgba(255,255,255,0.06);
}

section[data-testid="stSidebar"]
div[role="radiogroup"] label[data-checked="true"] {
    background:
        linear-gradient(
            90deg,
            rgba(99,102,241,0.20),
            rgba(139,92,246,0.08)
        );

    border-color: rgba(129,140,248,0.22);
}

/* Hide radio circles */

section[data-testid="stSidebar"]
div[role="radiogroup"] label > div:first-child {
    display: none;
}

/* Footer */

.sidebar-spacer {
    height: 25vh;
}

.sidebar-footer {
    display: flex;
    align-items: center;
    gap: 11px;

    padding: 12px;

    border-radius: 13px;

    background: rgba(255,255,255,0.035);
    border: 1px solid rgba(255,255,255,0.06);
}

.footer-icon {
    width: 32px;
    height: 32px;

    display: flex;
    align-items: center;
    justify-content: center;

    border-radius: 9px;

    background: rgba(99,102,241,0.15);
    color: #a5b4fc;

    font-size: 18px;
}

.footer-title {
    font-size: 12px;
    font-weight: 600;
    color: #e2e8f0;
}

.footer-subtitle {
    font-size: 10px;
    color: #64748b;
    margin-top: 2px;
}

    section[data-testid="stSidebar"] {
        background: #0b0f1a;
        border-right: 1px solid rgba(255,255,255,0.06);
    }

    section[data-testid="stSidebar"] .block-container {
        padding-top: 2rem;
    }

    /* ---------- BUTTONS ---------- */

.stButton > button {
    width: 100%;
    min-height: 48px;
    border-radius: 12px;
    border: 1px solid rgba(255,255,255,0.12);
    background: rgba(255,255,255,0.06);
    color: #f5f7fb;
    font-weight: 650;
    transition: all 0.15s ease;
}

.stButton > button:hover {
    border-color: #818cf8;
    background: rgba(99,102,241,0.18);
    transform: translateY(-2px);
}

.stButton > button:active {
    transform: scale(0.97);
}

.stButton > button[kind="primary"] {
    background: linear-gradient(
        135deg,
        #6366f1,
        #8b5cf6
    ) !important;

    border: 1px solid #a5b4fc !important;
    box-shadow: 0 5px 20px rgba(99,102,241,0.25);
    color: white !important;
}

    .stButton > button {
        border-radius: 10px;
        border: 1px solid rgba(255,255,255,0.08);
        background: rgba(255,255,255,0.05);
        color: #f5f7fb;
        font-weight: 600;
        transition: all 0.2s ease;
    }

    .stButton > button:hover {
        border-color: rgba(99,102,241,0.7);
        background: rgba(99,102,241,0.15);
        transform: translateY(-1px);
    }

    /* ---------- METRIC CARDS ---------- */

    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.045);
        border: 1px solid rgba(255,255,255,0.07);
        border-radius: 16px;
        padding: 20px;
    }

    div[data-testid="stMetricLabel"] {
        color: #9ca3af;
    }

    div[data-testid="stMetricValue"] {
        font-weight: 700;
    }

    /* ---------- DATAFRAMES ---------- */

    div[data-testid="stDataFrame"] {
        border-radius: 14px;
        overflow: hidden;
    }

    /* ---------- DIVIDERS ---------- */

    hr {
        border-color: rgba(255,255,255,0.07);
    }

</style>
""", unsafe_allow_html=True)

# -----------------------------
# Database
# -----------------------------
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_settings (
            user_id TEXT NOT NULL,
            key TEXT NOT NULL,
            value TEXT,
            PRIMARY KEY (user_id, key)
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_attendance (
            user_id TEXT NOT NULL,
            class_key TEXT NOT NULL,
            class_date TEXT NOT NULL,
            start_time TEXT NOT NULL,
            end_time TEXT NOT NULL,
            subject TEXT NOT NULL,
            faculty TEXT,
            status TEXT NOT NULL,
            PRIMARY KEY (user_id, class_key)
        )
    """)

    conn.commit()
    conn.close()

def get_setting(key):
    conn = get_conn()

    row = conn.execute(
        """
        SELECT value
        FROM user_settings
        WHERE user_id = ? AND key = ?
        """,
        (USER_ID, key)
    ).fetchone()

    conn.close()

    return row["value"] if row else None

def set_setting(key, value):
    conn = get_conn()

    conn.execute(
        """
        INSERT OR REPLACE INTO user_settings
        (user_id, key, value)
        VALUES (?, ?, ?)
        """,
        (USER_ID, key, value)
    )

    conn.commit()
    conn.close()

def save_attendance(class_row, status):
    class_key = make_class_key(class_row)

    conn = get_conn()

    conn.execute("""
        INSERT OR REPLACE INTO user_attendance
        (
            user_id,
            class_key,
            class_date,
            start_time,
            end_time,
            subject,
            faculty,
            status
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        USER_ID,
        class_key,
        class_row["date"].strftime("%Y-%m-%d"),
        class_row["start"],
        class_row["end"],
        class_row["subject"],
        class_row["faculty"],
        status,
    ))

    conn.commit()
    conn.close()


def attendance_statuses():
    conn = get_conn()

    rows = conn.execute(
        """
        SELECT
            class_key,
            class_date,
            start_time,
            end_time,
            subject,
            faculty,
            status
        FROM user_attendance
        WHERE user_id = ?
        """,
        (USER_ID,)
    ).fetchall()

    conn.close()

    return pd.DataFrame([dict(r) for r in rows])

def get_mess_menu(pdf_path, target_date):
    days = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday"
    ]

    day_name = target_date.strftime("%A")

    if day_name not in days:
        return None

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]

        # Coordinates for the August 2026 IIM Bodh Gaya menu PDF
        x = [100.7, 231.8, 456.0, 755.6, 984.1, 1339.3]

        # Horizontal boundaries for Monday → Sunday
        y = [
            157.9,
            238.2,
            318.4,
            406.7,
            487.0,
            567.2,
            647.5,
            742.5
        ]

        day_index = days.index(day_name)

        meals = {}

        meal_names = [
            "Breakfast",
            "Lunch",
            "Snacks",
            "Dinner"
        ]

        for i, meal in enumerate(meal_names, start=1):

            crop = page.crop(
                (
                    x[i] + 2,
                    y[day_index] + 2,
                    x[i + 1] - 2,
                    y[day_index + 1] - 2
                )
            )

            text = crop.extract_text(
                x_tolerance=2,
                y_tolerance=3
            )

            if text:
                text = " ".join(text.split())

                text = text.replace('<div class="mess-items">', "")
                text = text.replace("</div>", "")
            else:
                text = "Menu unavailable"

            meals[meal] = text

        return {
            "day": day_name,
            "date": target_date,
            "meals": meals
        }

# -----------------------------
# Timetable parser
# -----------------------------
EVENT_WORDS = (
    "SUMMER INTERNSHIP",
    "EXAMINATIONS",
    "HOLIDAYS",
    "JAYANTI",
    "E-MILAD",
    "GYANODAYA",
    "QUIZ SLOT",
)

TIME_RE = re.compile(r"^\s*(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*$")

def normalize_subject(s):
    s = re.sub(r"\s+", " ", str(s)).strip()
    s = re.sub(r"\s+\(", " (", s)
    s = re.sub(r"\(Sec\s*([A-Z])\)", r"(Sec \1)", s, flags=re.I)
    return s

def is_event(subject):
    upper = subject.upper()
    return any(word in upper for word in EVENT_WORDS)

def split_lines(value):
    if pd.isna(value):
        return []
    return [x.strip() for x in str(value).splitlines() if x.strip()]

@st.cache_data
def load_timetable():
    raw = pd.read_excel(TIMETABLE_PATH, sheet_name=0, header=None)

    # Row 0 contains the time slots.
    time_slots = {}
    for col in range(2, raw.shape[1]):
        match = TIME_RE.match(str(raw.iloc[0, col]))
        if match:
            time_slots[col] = (match.group(1), match.group(2))

    records = []

    # Each date occupies two rows: subject row followed by faculty row.
    for row in range(1, raw.shape[0], 2):
        if row + 1 >= raw.shape[0]:
            break

        raw_date = raw.iloc[row, 0]
        raw_day = raw.iloc[row, 1]

        if pd.isna(raw_date):
            continue

        try:
            class_date = pd.to_datetime(raw_date).date()
        except Exception:
            continue

        day_name = str(raw_day).strip() if not pd.isna(raw_day) else class_date.strftime("%A")

        for col, (start, end) in time_slots.items():
            subject_lines = split_lines(raw.iloc[row, col])
            faculty_lines = split_lines(raw.iloc[row + 1, col])

            for idx, subject in enumerate(subject_lines):
                subject = normalize_subject(subject)
                if not subject or is_event(subject):
                    continue

                faculty = faculty_lines[idx] if idx < len(faculty_lines) else ""
                records.append({
                    "date": class_date,
                    "day": day_name,
                    "start": start,
                    "end": end,
                    "subject": subject,
                    "faculty": faculty,
                })

    result = pd.DataFrame(records)

    if result.empty:
        return result

    return result.sort_values(["date", "start", "subject"]).reset_index(drop=True)

def make_class_key(row):
    return f'{row["date"].isoformat()}|{row["start"]}|{row["end"]}|{row["subject"]}|{row["faculty"]}'

def get_electives(tt):
    # The timetable itself is the source of truth.
    # Every actual course found in the workbook becomes selectable.
    subjects = sorted(tt["subject"].dropna().unique().tolist())
    return subjects

# -----------------------------
# Attendance calculations
# -----------------------------
def calculate_stats(selected_tt):
    att = attendance_statuses()

    if selected_tt.empty:
        return pd.DataFrame()

    if att.empty:
        return pd.DataFrame()

    selected_keys = set(selected_tt.apply(make_class_key, axis=1))
    att = att[att["class_key"].isin(selected_keys)].copy()

    if att.empty:
        return pd.DataFrame()

    rows = []
    for subject, group in att.groupby("subject"):
        present = int((group["status"] == "PRESENT").sum())
        od = int((group["status"] == "OD").sum())
        skipped = int((group["status"] == "SKIP").sum())

        denominator = present + od + skipped
        pct = ((present + od) / denominator * 100) if denominator else 0

        rows.append({
            "Subject": subject,
            "Present": present,
            "OD": od,
            "Skipped": skipped,
            "Attendance": pct,
        })

    return pd.DataFrame(rows).sort_values("Subject")

def make_ics(selected_tt):
    cal = Calendar()
    cal.add("prodid", "-//Personal Attendance Tracker//EN")
    cal.add("version", "2.0")
    cal.add("X-WR-CALNAME", "My Class Timetable")

    for _, row in selected_tt.iterrows():
        start_dt = datetime.combine(
            row["date"],
            datetime.strptime(row["start"], "%H:%M").time()
        )
        end_dt = datetime.combine(
            row["date"],
            datetime.strptime(row["end"], "%H:%M").time()
        )

        event = Event()
        event.add("summary", row["subject"])
        event.add("dtstart", start_dt)
        event.add("dtend", end_dt)
        if row["faculty"]:
            event.add("description", f'Faculty: {row["faculty"]}')
        event.add("uid", make_class_key(row))
        cal.add_component(event)

    return cal.to_ical()

# -----------------------------
# App
# -----------------------------
init_db()
tt = load_timetable()

if tt.empty:
    st.error("No timetable classes were found. Check timetable.xlsx.")
    st.stop()

all_electives = [
    x for x in get_electives(tt)
    if x.lower() != "independence day"
]

saved = get_setting("selected_electives")

if saved:
    selected_electives = [
        x for x in saved.split("|||")
        if x
    ]
else:
    selected_electives = []

# Sidebar
# Sidebar
with st.sidebar:

    st.markdown("""
    <div class="sidebar-brand">
        <div class="brand-icon">RC</div>
        <div>
            <div class="brand-title">Tracker</div>
            <div class="brand-subtitle">MBA • Term IV</div>
        </div>
    </div>
""", unsafe_allow_html=True)

    st.markdown("<div class='nav-label'>MENU</div>", unsafe_allow_html=True)

    page = st.radio(
        "Navigation",
        [
            "Setup",
            "Today",
            "Attendance",
            "Timetable",
            "Subjects",
            "Mess Menu"
            
        ],
        label_visibility="collapsed",
    )

    st.markdown("<div class='sidebar-spacer'></div>", unsafe_allow_html=True)

    st.markdown("""
        <div class="sidebar-footer">
            <div class="footer-icon">◎</div>
            <div>
                <div class="footer-title">Your Attendance</div>
                <div class="footer-subtitle">Track • Plan • Attend</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

# Setup page
if page == "Setup":
    st.title("Choose your electives")
    st.caption("These options are discovered automatically from your Excel timetable.")

    new_selection = st.multiselect(
        "Select the classes that belong to you",
        options=all_electives,
        default=[x for x in selected_electives if x in all_electives],
    )

    st.info(
        "Sections such as Sec A, Sec B and Sec C are treated as separate electives, "
        "so you can simply select the exact one you take."
    )

    if st.button("Save & Build My Timetable", type="primary"):
        set_setting(
            "selected_electives",
            "|||".join(new_selection)
        )

        st.success("✓ Your electives have been saved successfully!")
        st.info("Your timetable is now personalized to your selected subjects.")

    st.divider()

    st.subheader("All electives found in the Excel")
    st.write(f"{len(all_electives)} selectable courses found.")

    for subject in all_electives:
        st.write(f"• {subject}")

    st.stop()

# If no electives selected, guide user to setup.
# First-run setup: show the elective selector immediately.
# if not selected_electives:
#     st.title("Welcome to your Attendance Tracker")
#     st.write(
#         "Select the electives that belong to you. "
#         "Sections such as Sec A/B/C are separate options."
#     )

#     new_selection = st.multiselect(
#         "Your electives",
#         options=all_electives,
#         placeholder="Select your electives..."
#     )

#     if new_selection:
#         st.write("Selected:")
#         for subject in new_selection:
#             st.write(f"• {subject}")

#     if st.button(
#         "Save & Build My Timetable",
#         type="primary",
#         disabled=not new_selection
#     ):
#         st.session_state.selected_electives = new_selection
#         st.success("Your timetable has been created.")
#         st.rerun()

#     st.stop() 

selected_tt = tt[tt["subject"].isin(selected_electives)].copy()

# Today
if page == "Today":
    today = datetime.now(ZoneInfo("Asia/Kolkata")).date()
    today_tt = selected_tt[selected_tt["date"] == today].copy()

    st.title("Today")    
    st.markdown("""
<div class="semester-card">
    <div>
        <div class="semester-label">TERM IV</div>
        <div class="semester-title">Your Semester</div>
        <div class="semester-subtitle">
            Keep track of every class, attendance and OD.
        </div>
    </div>
    <div class="semester-icon">◷</div>
</div>
""", unsafe_allow_html=True)
    st.caption(today.strftime("%A, %d %B %Y"))



    if today_tt.empty:
        st.success("No selected classes today.")
    else:
        for idx, (_, row) in enumerate(today_tt.iterrows()):
            st.subheader(row["subject"])
            st.write(f'**{row["start"]} – {row["end"]}**  •  {row["faculty"]}')

            current = attendance_statuses()
            existing = ""
            if not current.empty:
                key = make_class_key(row)
                match = current[current["class_key"] == key]
                if not match.empty:
                    existing = match.iloc[0]["status"]

            cols = st.columns(3)

            if cols[0].button(
                "✓ Present",
                key=f"present_{idx}_{make_class_key(row)}",
                type="primary" if existing == "PRESENT" else "secondary",
            ):
                save_attendance(row, "PRESENT")
                st.rerun()

            if cols[1].button(
                "🟡 OD",
                key=f"od_{idx}_{make_class_key(row)}",
                type="primary" if existing == "OD" else "secondary",
            ):
                save_attendance(row, "OD")
                st.rerun()

            if cols[2].button(
                "✕ Skip",
                key=f"skip_{idx}_{make_class_key(row)}",
                type="primary" if existing == "SKIP" else "secondary",
            ):
                save_attendance(row, "SKIP")
                st.rerun()

            if existing:
                st.caption(f"Recorded: **{existing}**")
            st.divider()

# Attendance
# Attendance
elif page == "Attendance":
    st.title("Attendance")

    # -----------------------------------------
    # Overall attendance summary
    # -----------------------------------------
    stats = calculate_stats(selected_tt)

    if stats.empty:
        st.info("No attendance has been marked yet.")
    else:
        total_present = stats["Present"].sum()
        total_od = stats["OD"].sum()
        total_skipped = stats["Skipped"].sum()

        denom = total_present + total_od + total_skipped
        overall = ((total_present + total_od) / denom * 100) if denom else 0
        color = "#22c55e" if overall >= 80 else "#ef4444"

        circumference = 565.49
        dash_offset = circumference * (1 - overall / 100)
        color = "#22c55e" if overall >= 80 else "#ef4444"
        st.markdown(
            f"""<div class="attendance-ring" style="--dash-offset: {dash_offset}px;">
<svg class="attendance-svg" viewBox="0 0 220 220">
<circle class="ring-background" cx="110" cy="110" r="90"></circle>
<circle class="ring-progress" cx="110" cy="110" r="90"></circle>
</svg>
<div class="attendance-ring-value" style="color: {color};">{overall:.1f}%</div>
</div>""",
            unsafe_allow_html=True
        )
        c1, c2, c3, c4 = st.columns(4)

        c1.metric("Overall", f"{overall:.1f}%")
        c2.metric("Present", int(total_present))
        c3.metric("OD", int(total_od))
        c4.metric("Skipped", int(total_skipped))

        st.divider()

    # -----------------------------------------
    # Select any date
    # -----------------------------------------
    st.subheader("Edit Attendance")

    available_dates = sorted(selected_tt["date"].unique())

    if available_dates:

        selected_date = st.date_input(
            "Select class date",
            value=datetime.now(ZoneInfo("Asia/Kolkata")).date(),
            min_value=available_dates[0],
            max_value=available_dates[-1],
        )

        # Classes on selected date
        date_classes = selected_tt[
            selected_tt["date"] == selected_date
        ].copy()

        if date_classes.empty:
            st.info("You have no selected classes on this date.")

        else:

            # Existing attendance records
            history = attendance_statuses()

            for idx, (_, row) in enumerate(date_classes.iterrows()):

                st.markdown("---")

                st.subheader(row["subject"])

                faculty_text = (
                    row["faculty"]
                    if row["faculty"]
                    else "Faculty not listed"
                )

                st.write(
                    f'**{row["start"]} – {row["end"]}**  •  '
                    f'{faculty_text}'
                )

                # Find existing status
                existing_status = ""

                if not history.empty:
                    class_key = make_class_key(row)

                    match = history[
                        history["class_key"] == class_key
                    ]

                    if not match.empty:
                        existing_status = match.iloc[0]["status"]

                # Buttons
                col1, col2, col3 = st.columns(3)

                if col1.button(
                    "✓ Present",
                    key=f"attendance_present_{idx}_{selected_date}",
                    type=(
                        "primary"
                        if existing_status == "PRESENT"
                        else "secondary"
                    ),
                ):
                    save_attendance(row, "PRESENT")
                    st.rerun()

                if col2.button(
                    "🟡 OD",
                    key=f"attendance_od_{idx}_{selected_date}",
                    type=(
                        "primary"
                        if existing_status == "OD"
                        else "secondary"
                    ),
                ):
                    save_attendance(row, "OD")
                    st.rerun()

                if col3.button(
                    "✕ Skip",
                    key=f"attendance_skip_{idx}_{selected_date}",
                    type=(
                        "primary"
                        if existing_status == "SKIP"
                        else "secondary"
                    ),
                ):
                    save_attendance(row, "SKIP")
                    st.rerun()

                if existing_status:
                    st.caption(
                        f"Current status: **{existing_status}**"
                    )
                else:
                    st.caption("Not marked yet.")

    st.divider()

    # -----------------------------------------
    # Subject-wise attendance
    # -----------------------------------------
    st.subheader("Subject Attendance")

    stats = calculate_stats(selected_tt)

    if stats.empty:
        st.info("No attendance data available yet.")
    else:
        display = stats.copy()

        display["Attendance"] = display["Attendance"].map(
            lambda x: f"{x:.1f}%"
        )

        st.dataframe(
            display,
            use_container_width=True,
            hide_index=True
        )

    # -----------------------------------------
    # Full attendance history
    # -----------------------------------------
    st.subheader("Attendance History")

    history = attendance_statuses()

    if history.empty:
        st.info("No attendance records yet.")

    else:
        history = history[
            history["subject"].isin(selected_electives)
        ].copy()

        history = history.sort_values(
            ["class_date", "start_time"]
        )

        st.dataframe(
            history[
                [
                    "class_date",
                    "start_time",
                    "end_time",
                    "subject",
                    "faculty",
                    "status",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )


# Timetable
elif page == "Timetable":
    st.title("My Timetable")

    st.download_button(
        "📅 Export My Timetable to Calendar",
        data=make_ics(selected_tt),
        file_name="my_term_iv_timetable.ics",
        mime="text/calendar",
    )

    view = selected_tt.copy()
    view["Date"] = view["date"].map(lambda x: x.strftime("%d %b %Y"))
    view = view.rename(columns={
        "day": "Day",
        "start": "Start",
        "end": "End",
        "subject": "Subject",
        "faculty": "Faculty",
    })

    st.dataframe(
        view[["Date", "Day", "Start", "End", "Subject", "Faculty"]],
        use_container_width=True,
        hide_index=True,
    )

# Subjects
elif page == "Subjects":
    st.title("My Subjects")

    stats = calculate_stats(selected_tt)

    if stats.empty:
        st.info("No attendance has been marked yet.")

    else:
        subject_columns = st.columns(3)
        for i, (_, subject_row) in enumerate(stats.iterrows()):

            subject = subject_row["Subject"]
            percentage = float(subject_row["Attendance"])

            color = "#22c55e" if percentage >= 80 else "#ef4444"

            circumference = 471.24
            dash_offset = circumference * (1 - percentage / 100)

            with subject_columns[i % 3]:

                circumference = 471.24
                dash_offset = circumference * (1 - percentage / 100)

                color = "#22c55e" if percentage >= 80 else "#ef4444"
                st.markdown(
                    f"""<div class="subject-card">
<div class="subject-name">{subject}</div>

<div class="subject-ring" style="--subject-dash-offset: {dash_offset}px;">

<svg class="subject-svg" viewBox="0 0 180 180">
<circle class="subject-bg" cx="90" cy="90" r="75"></circle>
<circle class="subject-progress" cx="90" cy="90" r="75"></circle>
</svg>

<div class="subject-percentage" style="color: {color};">
{percentage:.1f}%
</div>

</div>

<div class="subject-stats">
Present: {int(subject_row["Present"])}
&nbsp; • &nbsp;
OD: {int(subject_row["OD"])}
&nbsp; • &nbsp;
Skipped: {int(subject_row["Skipped"])}
</div>

</div>""",
                    unsafe_allow_html=True
                )
                
elif page == "Mess Menu":

    st.title("Mess Menu")

    # Always use Indian date
    menu_today = datetime.now(
        ZoneInfo("Asia/Kolkata")
    ).date()

    st.caption(
        menu_today.strftime("%A, %d %B %Y")
    )

    # August 2026 menu PDF
    menu_pdf = APP_DIR / "August Menu 2026.pdf"

    if not menu_pdf.exists():

        st.error(
            "Mess menu PDF not found."
        )

    else:

        menu = get_mess_menu(
            menu_pdf,
            menu_today
        )

        if menu is None:

            st.info(
                "No mess menu available for today."
            )

        else:

            st.markdown(
                f"""
                <div class="menu-day">
                    {menu["day"]}
                </div>
                """,
                unsafe_allow_html=True
            )

            for meal, items in menu["meals"].items():
                clean_items = re.sub(r"<[^>]*>", "", str(items))
                clean_items = " ".join(clean_items.split())

                html = (
                    f'<div class="mess-card">'
                    f'<div class="mess-meal">{meal}</div>'
                    f'<div class="mess-items">{clean_items}</div>'
                    f'</div>'
                )

                st.markdown(html, unsafe_allow_html=True)
# Safety fallback
else:
    st.title("Attendance Tracker")
    st.write("Use the sidebar to navigate.")

