"""
data.database
=============
SQLite setup for Resume AI Master.
Single database file: data/awis.db

Tables:
    email_logs   — every email sent through the system
    candidates   — CRM candidate records
    submissions  — employer submissions per candidate
    engagements  — every VSC interaction logged by type and subtype
"""

import sqlite3
import os
from datetime import datetime, timedelta

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "awis.db")

# ── Pipeline ──────────────────────────────────────────────────────────────────

PIPELINE_STAGES = [
    "NEW",
    "CONTACTED",
    "MEETING_SCHEDULED",
    "MEETING_COMPLETE",
    "RESUME_ANALYZED",
    "SUBMITTED",
    "HIRED",
    "REJECTED",
    "WITHDRAWN",
    "HELD_REVIEW",
]

STAGE_LABELS = {
    "NEW":              "New Lead",
    "CONTACTED":        "Contacted",
    "MEETING_SCHEDULED":"Meeting Scheduled",
    "MEETING_COMPLETE": "Meeting Complete",
    "RESUME_ANALYZED":  "Resume Analyzed",
    "SUBMITTED":        "Submitted",
    "HIRED":            "Hired",
    "REJECTED":         "Rejected",
    "WITHDRAWN":        "Withdrawn",
    "HELD_REVIEW":      "Held — Check SJB DB",
}

STAGE_COLORS = {
    "NEW":              "#667eea",
    "CONTACTED":        "#2196F3",
    "MEETING_SCHEDULED":"#FF9800",
    "MEETING_COMPLETE": "#9C27B0",
    "RESUME_ANALYZED":  "#00BCD4",
    "SUBMITTED":        "#FF5722",
    "HIRED":            "#4CAF50",
    "REJECTED":         "#f44336",
    "WITHDRAWN":        "#9E9E9E",
    "HELD_REVIEW":      "#B71C1C",
}

MILITARY_BRANCHES = [
    "Army", "Navy", "Marine Corps", "Air Force",
    "Space Force", "Coast Guard", "National Guard", "Reserves"
]

# ── Engagements ───────────────────────────────────────────────────────────────

ENGAGEMENT_TYPES = ["INTAKE", "COUNSELING", "MEETING", "COACHING", "SUBMITTED"]

ENGAGEMENT_SUBTYPES = {
    "INTAKE":     [],
    "COUNSELING": ["PHONE", "EMAIL", "IN_PERSON"],
    "MEETING":    ["VIRTUAL", "IN_PERSON", "PHONE", "MISSED"],
    "COACHING":   ["PHONE", "EMAIL", "IN_PERSON", "MOCK_INTERVIEW", "RESUME_REVIEW", "INTERVIEW_PREP"],
    "SUBMITTED":  [],
}

ENGAGEMENT_TYPE_LABELS = {
    "INTAKE":     "Intake",
    "COUNSELING": "Counseling",
    "MEETING":    "Meeting",
    "COACHING":   "Coaching",
    "SUBMITTED":  "Submitted",
}

ENGAGEMENT_SUBTYPE_LABELS = {
    "PHONE":          "Phone",
    "EMAIL":          "Email",
    "IN_PERSON":      "In-Person",
    "VIRTUAL":        "Virtual",
    "MISSED":         "Missed",
    "MOCK_INTERVIEW": "Mock Interview",
    "RESUME_REVIEW":  "Resume Review",
    "INTERVIEW_PREP": "Interview Prep",
}

# ── Connection ────────────────────────────────────────────────────────────────

def get_connection():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


# ── Init ──────────────────────────────────────────────────────────────────────

def init_db():
    """Create all tables. Safe to call on every startup."""
    conn = get_connection()
    c = conn.cursor()

    # Email logs
    c.execute("""
        CREATE TABLE IF NOT EXISTS email_logs (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            sent_at         TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            vsc_name        TEXT NOT NULL,
            vsc_email       TEXT,
            to_name         TEXT NOT NULL,
            to_email        TEXT NOT NULL,
            template_key    TEXT NOT NULL,
            template_label  TEXT NOT NULL,
            subject         TEXT NOT NULL,
            body            TEXT NOT NULL,
            status          TEXT NOT NULL DEFAULT 'sent',
            error_message   TEXT,
            candidate_id    INTEGER
        )
    """)

    # Add candidate_id to email_logs if it doesn't exist (migration)
    try:
        c.execute("ALTER TABLE email_logs ADD COLUMN candidate_id INTEGER")
    except Exception:
        pass  # Column already exists

    # Candidates
    c.execute("""
        CREATE TABLE IF NOT EXISTS candidates (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            is_deleted      INTEGER NOT NULL DEFAULT 0,

            vsc_name        TEXT NOT NULL,

            first_name      TEXT NOT NULL,
            last_name       TEXT NOT NULL,
            email           TEXT,
            phone           TEXT,
            city            TEXT,
            state           TEXT,

            branch          TEXT,
            rank            TEXT,
            mos             TEXT,
            years_served    TEXT,

            stage           TEXT NOT NULL DEFAULT 'NEW',
            notes           TEXT,
            resume_text     TEXT
        )
    """)

    # Submissions
    c.execute("""
        CREATE TABLE IF NOT EXISTS submissions (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            submitted_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            candidate_id    INTEGER NOT NULL,
            job_title       TEXT NOT NULL,
            company_name    TEXT NOT NULL,
            outcome         TEXT NOT NULL DEFAULT 'PENDING',
            notes           TEXT
        )
    """)

    # DB metadata — tracks last jobs CSV upload
    c.execute("""
        CREATE TABLE IF NOT EXISTS db_metadata (
            key     TEXT PRIMARY KEY,
            value   TEXT
        )
    """)

    # Job listings table
    c.execute("""
        CREATE TABLE IF NOT EXISTS job_listings (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            loaded_at       TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            job_id          TEXT,
            job_title       TEXT NOT NULL,
            company_name    TEXT,
            city            TEXT,
            state           TEXT,
            job_description TEXT,
            qualifications  TEXT,
            salary_from     TEXT,
            salary_to       TEXT,
            salary_period   TEXT,
            job_type        TEXT,
            latitude        REAL,
            longitude       REAL,
            board           TEXT,
            is_active       INTEGER NOT NULL DEFAULT 1
        )
    """)

    # Add source column to candidates (migration)
    try:
        c.execute("ALTER TABLE candidates ADD COLUMN source TEXT DEFAULT 'manual'")
    except Exception:
        pass  # Column already exists

    # Engagements — every VSC interaction
    c.execute("""
        CREATE TABLE IF NOT EXISTS engagements (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            candidate_id    INTEGER NOT NULL,
            vsc_name        TEXT NOT NULL,
            type            TEXT NOT NULL,
            subtype         TEXT,
            notes           TEXT,
            FOREIGN KEY (candidate_id) REFERENCES candidates(id)
        )
    """)

    # Custom email templates — VSC-saved templates
    c.execute("""
        CREATE TABLE IF NOT EXISTS custom_templates (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            key             TEXT UNIQUE NOT NULL,
            label           TEXT NOT NULL,
            stage           TEXT DEFAULT 'GENERAL',
            when_to_send    TEXT DEFAULT '',
            subject         TEXT NOT NULL,
            body            TEXT NOT NULL
        )
    """)

    # Users — login accounts for VSCs and admins
    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            email         TEXT UNIQUE NOT NULL,
            display_name  TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            role          TEXT NOT NULL DEFAULT 'vsc',
            is_active     INTEGER NOT NULL DEFAULT 1
        )
    """)

    conn.commit()
    conn.close()


# ── Users ─────────────────────────────────────────────────────────────────────

def get_user_by_email(email: str):
    conn = get_connection()
    row = conn.execute(
        "SELECT * FROM users WHERE email = ? AND is_active = 1",
        (email.lower().strip(),)
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def get_user_by_id(user_id: int):
    conn = get_connection()
    row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def seed_users(users: list):
    """Insert users that don't exist yet. Each dict needs email, display_name, password_hash, role."""
    conn = get_connection()
    for u in users:
        conn.execute(
            "INSERT OR IGNORE INTO users (email, display_name, password_hash, role) VALUES (?, ?, ?, ?)",
            (u['email'].lower().strip(), u['display_name'], u['password_hash'], u['role'])
        )
    conn.commit()
    conn.close()


# ── Email Logs ────────────────────────────────────────────────────────────────

def log_email(vsc_name, vsc_email, to_name, to_email,
              template_key, template_label, subject, body,
              status="sent", error_message=None, candidate_id=None):
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        INSERT INTO email_logs
            (vsc_name, vsc_email, to_name, to_email,
             template_key, template_label, subject, body,
             status, error_message, candidate_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (vsc_name, vsc_email, to_name, to_email,
          template_key, template_label, subject, body,
          status, error_message, candidate_id))
    row_id = c.lastrowid
    conn.commit()
    conn.close()
    return row_id


def get_email_history(limit=50, candidate_email=None):
    conn = get_connection()
    c = conn.cursor()
    if candidate_email:
        c.execute("""
            SELECT id, sent_at, vsc_name, to_name, to_email,
                   template_label, subject, status, candidate_id
            FROM email_logs
            WHERE to_email = ?
            ORDER BY sent_at DESC
            LIMIT ?
        """, (candidate_email, limit))
    else:
        c.execute("""
            SELECT id, sent_at, vsc_name, to_name, to_email,
                   template_label, subject, status, candidate_id
            FROM email_logs
            ORDER BY sent_at DESC
            LIMIT ?
        """, (limit,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_email_stats():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM email_logs WHERE status = 'sent'")
    total_sent = c.fetchone()[0]
    c.execute("""
        SELECT vsc_name, COUNT(*) as count FROM email_logs
        WHERE status = 'sent' GROUP BY vsc_name ORDER BY count DESC
    """)
    by_vsc = [dict(r) for r in c.fetchall()]
    c.execute("""
        SELECT template_label, COUNT(*) as count FROM email_logs
        WHERE status = 'sent' GROUP BY template_label ORDER BY count DESC
    """)
    by_template = [dict(r) for r in c.fetchall()]
    conn.close()
    return {"total_sent": total_sent, "by_vsc": by_vsc, "by_template": by_template}


# ── Candidates ────────────────────────────────────────────────────────────────

def create_candidate(data: dict) -> int:
    """Insert a new candidate. Returns new id."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        INSERT INTO candidates
            (vsc_name, first_name, last_name, email, phone,
             city, state, branch, rank, mos, years_served,
             stage, notes, resume_text, source)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data.get("vsc_name", ""),
        data.get("first_name", ""),
        data.get("last_name", ""),
        data.get("email", ""),
        data.get("phone", ""),
        data.get("city", ""),
        data.get("state", ""),
        data.get("branch", ""),
        data.get("rank", ""),
        data.get("mos", ""),
        data.get("years_served", ""),
        data.get("stage", "NEW"),
        data.get("notes", ""),
        data.get("resume_text", ""),
        data.get("source", "manual"),
    ))
    row_id = c.lastrowid
    conn.commit()
    conn.close()
    return row_id


def get_candidates(vsc_name=None, stage=None, search=None) -> list:
    conn = get_connection()
    c = conn.cursor()
    query = """
        SELECT c.*,
               (SELECT COUNT(*) FROM submissions s WHERE s.candidate_id = c.id) AS submission_count,
               (SELECT MAX(sent_at) FROM email_logs e WHERE e.to_email = c.email) AS last_email_at
        FROM candidates c
        WHERE c.is_deleted = 0
    """
    params = []
    if vsc_name:
        query += " AND c.vsc_name = ?"
        params.append(vsc_name)
    if stage:
        query += " AND c.stage = ?"
        params.append(stage)
    if search:
        query += " AND (c.first_name LIKE ? OR c.last_name LIKE ? OR c.email LIKE ? OR c.company LIKE ?)"
        s = f"%{search}%"
        params.extend([s, s, s, s])
    query += " ORDER BY c.updated_at DESC"
    c.execute(query, params)
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_candidate(candidate_id: int) -> dict:
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT * FROM candidates WHERE id = ? AND is_deleted = 0", (candidate_id,))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def update_candidate(candidate_id: int, data: dict) -> bool:
    allowed = {
        "vsc_name", "first_name", "last_name", "email", "phone",
        "city", "state", "branch", "rank", "mos", "years_served",
        "stage", "notes", "resume_text", "source"
    }
    fields = {k: v for k, v in data.items() if k in allowed}
    if not fields:
        return False
    fields["updated_at"] = datetime.now().isoformat()
    set_clause = ", ".join(f"{k} = ?" for k in fields)
    values = list(fields.values()) + [candidate_id]
    conn = get_connection()
    c = conn.cursor()
    c.execute(f"UPDATE candidates SET {set_clause} WHERE id = ?", values)
    changed = c.rowcount > 0
    conn.commit()
    conn.close()
    return changed


def delete_candidate(candidate_id: int) -> bool:
    conn = get_connection()
    c = conn.cursor()
    c.execute("UPDATE candidates SET is_deleted = 1 WHERE id = ?", (candidate_id,))
    changed = c.rowcount > 0
    conn.commit()
    conn.close()
    return changed


def get_pipeline_counts(vsc_name=None) -> dict:
    conn = get_connection()
    c = conn.cursor()
    query = """
        SELECT stage, COUNT(*) as count FROM candidates
        WHERE is_deleted = 0
    """
    params = []
    if vsc_name:
        query += " AND vsc_name = ?"
        params.append(vsc_name)
    query += " GROUP BY stage"
    c.execute(query, params)
    counts = {row["stage"]: row["count"] for row in c.fetchall()}
    c.execute(
        "SELECT COUNT(*) FROM candidates WHERE is_deleted = 0" +
        (" AND vsc_name = ?" if vsc_name else ""),
        params
    )
    total = c.fetchone()[0]
    conn.close()
    return {"total": total, "by_stage": counts}


# ── Submissions ───────────────────────────────────────────────────────────────

def add_submission(candidate_id: int, job_title: str,
                   company_name: str, notes: str = "") -> int:
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        INSERT INTO submissions (candidate_id, job_title, company_name, notes)
        VALUES (?, ?, ?, ?)
    """, (candidate_id, job_title, company_name, notes))
    # Advance candidate stage to SUBMITTED
    c.execute("""
        UPDATE candidates SET stage = 'SUBMITTED', updated_at = ?
        WHERE id = ? AND stage NOT IN ('HIRED', 'REJECTED', 'WITHDRAWN')
    """, (datetime.now().isoformat(), candidate_id))
    row_id = c.lastrowid
    conn.commit()
    conn.close()
    return row_id


def get_submissions(candidate_id: int) -> list:
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        SELECT * FROM submissions WHERE candidate_id = ?
        ORDER BY submitted_at DESC
    """, (candidate_id,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def update_submission(sub_id: int, outcome: str, notes: str = None) -> bool:
    conn = get_connection()
    c = conn.cursor()
    if notes is not None:
        c.execute("UPDATE submissions SET outcome = ?, notes = ? WHERE id = ?",
                  (outcome, notes, sub_id))
    else:
        c.execute("UPDATE submissions SET outcome = ? WHERE id = ?", (outcome, sub_id))

    # If hired, advance candidate stage
    if outcome == "HIRED":
        c.execute("""
            UPDATE candidates SET stage = 'HIRED', updated_at = ?
            WHERE id = (SELECT candidate_id FROM submissions WHERE id = ?)
        """, (datetime.now().isoformat(), sub_id))
    elif outcome == "REJECTED":
        # Only move to rejected if not already hired elsewhere
        c.execute("""
            UPDATE candidates SET stage = 'REJECTED', updated_at = ?
            WHERE id = (SELECT candidate_id FROM submissions WHERE id = ?)
            AND stage NOT IN ('HIRED')
        """, (datetime.now().isoformat(), sub_id))

    changed = c.rowcount > 0
    conn.commit()
    conn.close()
    return changed


# ── Job Listings ──────────────────────────────────────────────────────────────

def save_jobs(jobs: list) -> int:
    """
    Replace all job listings with new list.
    Deactivates all existing, inserts new batch.
    Returns count inserted.
    """
    conn = get_connection()
    c = conn.cursor()

    # Deactivate all existing
    c.execute("UPDATE job_listings SET is_active = 0")

    rows = []
    for job in jobs:
        job_title = str(job.get('Job Title') or job.get('job_title') or '')
        if not job_title:
            continue
        rows.append((
            str(job.get('Job Id')        or job.get('job_id')        or ''),
            job_title,
            str(job.get('Company Name')  or job.get('company_name')  or ''),
            str(job.get('City')          or job.get('city')          or ''),
            str(job.get('State')         or job.get('state')         or ''),
            str(job.get('Job Description') or job.get('job_description') or ''),
            str(job.get('Qualifications') or job.get('qualifications') or ''),
            str(job.get('Salary From')   or job.get('salary_from')   or ''),
            str(job.get('Salary To')     or job.get('salary_to')     or ''),
            str(job.get('Salary Period') or job.get('salary_period') or ''),
            str(job.get('Job Type')      or job.get('job_type')      or ''),
            job.get('latitude'),
            job.get('longitude'),
            str(job.get('Board')         or job.get('board')         or ''),
        ))

    c.executemany("""
        INSERT INTO job_listings
            (job_id, job_title, company_name, city, state,
             job_description, qualifications, salary_from, salary_to,
             salary_period, job_type, latitude, longitude, board, is_active)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
    """, rows)

    conn.commit()
    conn.close()
    return len(rows)


def get_jobs_near(lat: float, lon: float, radius_miles: float = 50) -> list:
    """
    Return jobs within radius_miles of lat/lon.
    Uses haversine approximation via SQL.
    Returns list of dicts with distance_miles added.
    Sorted by distance ascending.
    """
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        SELECT * FROM (
            SELECT *,
              (3958.8 * acos(
                cos(radians(?)) * cos(radians(latitude)) *
                cos(radians(longitude) - radians(?)) +
                sin(radians(?)) * sin(radians(latitude))
              )) AS distance_miles
            FROM job_listings
            WHERE is_active = 1
              AND latitude IS NOT NULL
              AND longitude IS NOT NULL
        ) WHERE distance_miles <= ?
        ORDER BY distance_miles ASC
        LIMIT 100
    """, (lat, lon, lat, radius_miles))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_job_by_title(title: str) -> dict:
    """Find best matching job by title. Exact match first, then LIKE. None if not found."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        SELECT * FROM job_listings
        WHERE LOWER(job_title) = LOWER(?) AND is_active = 1
        LIMIT 1
    """, (title,))
    row = c.fetchone()
    if not row:
        c.execute("""
            SELECT * FROM job_listings
            WHERE job_title LIKE ? AND is_active = 1
            LIMIT 1
        """, (f'%{title}%',))
        row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def get_job_by_id(job_id: str) -> dict:
    """Return single job by job_id field (not primary key). None if not found."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        SELECT * FROM job_listings
        WHERE job_id = ? AND is_active = 1
        LIMIT 1
    """, (job_id,))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def get_all_active_jobs(limit: int = 500) -> list:
    """Return all active jobs, ordered by job_title."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        SELECT * FROM job_listings
        WHERE is_active = 1
        ORDER BY job_title ASC
        LIMIT ?
    """, (limit,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def load_jobs_to_memory() -> list:
    """Return all active jobs as list of dicts for JOBS_DB in-memory cache."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        SELECT
            id,
            job_id          AS "Job Id",
            job_title       AS "Job Title",
            company_name    AS "Company Name",
            city            AS "City",
            state           AS "State",
            job_description AS "Job Description",
            qualifications  AS "Qualifications",
            salary_from     AS "Salary From",
            salary_to       AS "Salary To",
            salary_period   AS "Salary Period",
            job_type        AS "Job Type",
            latitude,
            longitude,
            board           AS "Board"
        FROM job_listings
        WHERE is_active = 1
        ORDER BY job_title ASC
    """)
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def log_job_upload(count: int):
    """Record the timestamp and count of the most recent jobs CSV upload."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO db_metadata (key, value) VALUES ('last_upload_at', ?)",
              (datetime.utcnow().isoformat(),))
    c.execute("INSERT OR REPLACE INTO db_metadata (key, value) VALUES ('last_upload_count', ?)",
              (str(count),))
    conn.commit()
    conn.close()


def get_last_job_upload() -> dict:
    """Return last upload timestamp, age in days, and job count."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT key, value FROM db_metadata WHERE key IN ('last_upload_at','last_upload_count')")
    rows = {r['key']: r['value'] for r in c.fetchall()}
    conn.close()
    last_at  = rows.get('last_upload_at')
    count    = int(rows.get('last_upload_count', 0))
    age_days = None
    if last_at:
        try:
            age_days = (datetime.utcnow() - datetime.fromisoformat(last_at)).days
        except Exception:
            pass
    return {'last_upload_at': last_at, 'age_days': age_days, 'count': count}


def delete_job_by_title(title: str) -> int:
    """Remove all job listings matching title (case-insensitive). Returns rows deleted."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("DELETE FROM job_listings WHERE LOWER(job_title) = LOWER(?)", (title.strip(),))
    deleted = c.rowcount
    conn.commit()
    conn.close()
    return deleted


# ── Engagements ───────────────────────────────────────────────────────────────

def log_engagement(candidate_id: int, vsc_name: str, eng_type: str,
                   subtype: str = None, notes: str = None) -> int:
    """
    Record a VSC engagement with a candidate.
    type must be one of ENGAGEMENT_TYPES.
    subtype must be valid for the given type, or None.
    Returns new engagement id.
    """
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        INSERT INTO engagements (candidate_id, vsc_name, type, subtype, notes)
        VALUES (?, ?, ?, ?, ?)
    """, (candidate_id, vsc_name, eng_type, subtype, notes))
    # Update candidate updated_at so ghost detection resets
    c.execute("""
        UPDATE candidates SET updated_at = ? WHERE id = ?
    """, (datetime.now().isoformat(), candidate_id))
    row_id = c.lastrowid
    conn.commit()
    conn.close()
    return row_id


def get_engagements(candidate_id: int) -> list:
    """Return all engagements for a candidate, newest first."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        SELECT * FROM engagements
        WHERE candidate_id = ?
        ORDER BY created_at DESC
    """, (candidate_id,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_candidates_needing_followup(days: int = 7) -> list:
    """
    Return candidates who have had no engagement in `days` days
    and are not in a terminal stage (HIRED, REJECTED, WITHDRAWN).
    Used by ghost detection.
    """
    cutoff = (datetime.now() - timedelta(days=days)).isoformat()
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        SELECT c.*,
               (SELECT MAX(created_at) FROM engagements e WHERE e.candidate_id = c.id) AS last_engagement_at
        FROM candidates c
        WHERE c.is_deleted = 0
          AND c.stage NOT IN ('HIRED', 'REJECTED', 'WITHDRAWN')
          AND (
              (SELECT MAX(created_at) FROM engagements e WHERE e.candidate_id = c.id) < ?
              OR
              (SELECT COUNT(*) FROM engagements e WHERE e.candidate_id = c.id) = 0
          )
          AND c.created_at < ?
        ORDER BY c.updated_at ASC
    """, (cutoff, cutoff))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_intake_submissions(limit: int = 50) -> list:
    """Return candidates where source='intake', newest first, with submission_count."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        SELECT c.*,
               (SELECT COUNT(*) FROM submissions s WHERE s.candidate_id = c.id) AS submission_count,
               (SELECT MAX(sent_at) FROM email_logs e WHERE e.to_email = c.email) AS last_email_at
        FROM candidates c
        WHERE c.is_deleted = 0
          AND c.source = 'intake'
        ORDER BY c.created_at DESC
        LIMIT ?
    """, (limit,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


# ── Custom Templates ──────────────────────────────────────────────────────────

def save_custom_template(label: str, subject: str, body: str,
                         stage: str = 'GENERAL', when_to_send: str = '') -> dict:
    """
    Save a VSC-authored template. Generates a unique key from the label.
    Returns the saved template as a dict.
    """
    import re
    import time
    base = re.sub(r'[^a-z0-9]+', '_', label.lower()).strip('_')
    key  = f"custom_{base}_{int(time.time())}"
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
        INSERT INTO custom_templates (key, label, stage, when_to_send, subject, body)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (key, label, stage, when_to_send, subject, body))
    conn.commit()
    conn.close()
    return {'key': key, 'label': label, 'stage': stage,
            'when': when_to_send, 'subject': subject, 'body': body, 'fields': []}


def get_custom_templates() -> list:
    """Return all VSC-saved custom templates, newest first."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT * FROM custom_templates ORDER BY created_at DESC")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def delete_custom_template(key: str) -> bool:
    """Delete a custom template by key. Returns True if deleted."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("DELETE FROM custom_templates WHERE key = ?", (key,))
    deleted = c.rowcount > 0
    conn.commit()
    conn.close()
    return deleted
