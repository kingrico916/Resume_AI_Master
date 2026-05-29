"""
Resume AI Master - Work for Warriors
Flask application for resume analysis with job matching and CalCareers packaging.
"""

from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
import requests
import os
from dotenv import load_dotenv
load_dotenv(override=True)
import re
import json
import logging
import pandas as pd
from datetime import datetime
import smtplib
import imaplib
import email
import email.utils
from email.header import decode_header as _decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import tempfile
import zipfile
import threading
from pathlib import Path
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
from data.database import (
    init_db, log_email, get_email_history, get_email_stats,
    create_candidate, get_candidates, get_candidate, update_candidate,
    delete_candidate, get_pipeline_counts, add_submission,
    get_submissions, update_submission,
    save_jobs, get_jobs_near, get_job_by_id, get_job_by_title, get_all_active_jobs,
    load_jobs_to_memory, get_intake_submissions,
    log_engagement, get_engagements, get_candidates_needing_followup,
    save_custom_template, get_custom_templates, delete_custom_template,
    log_job_upload, get_last_job_upload, delete_job_by_title,
    get_user_by_email, get_user_by_id, seed_users, get_dashboard_stats,
    archive_email_log, restore_email_log, hard_delete_email_log,
    PIPELINE_STAGES, STAGE_LABELS, STAGE_COLORS, MILITARY_BRANCHES,
    ENGAGEMENT_TYPES, ENGAGEMENT_SUBTYPES, ENGAGEMENT_TYPE_LABELS, ENGAGEMENT_SUBTYPE_LABELS
)
from data.email_templates import TEMPLATES
from services.email_service import send_email as svc_send_email, preview_email as svc_preview_email
from services.resume_parser import parse_resume
from services.automation_engine import run_automation

# CalCareers imports
from core.data_models import (
    CandidateProfile, JobTarget, VeteransPreference,
    PackageInput, EducationEntry, WorkExperienceEntry,
    TemplateTrack, ApplicationBasis
)
from core.decision_engine import DecisionEngine
from core.package_generator import PackageGenerator
from core.checklist_generator import ChecklistGenerator
from core.audit_logger import AuditLogger

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'dev-change-me-in-production')
app.config['DEBUG'] = os.getenv('FLASK_DEBUG', 'false').lower() == 'true'
logging.getLogger('werkzeug').setLevel(logging.ERROR)

_print = __builtins__['print'] if isinstance(__builtins__, dict) else __builtins__.print
def print(*args, **kwargs):
    kwargs.setdefault('flush', True)
    _print(f"[{datetime.now().strftime('%H:%M:%S')}]", *args, **kwargs)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB max

# Initialize database tables on startup
init_db()

_SEED_USERS = [
    {"email": "tracey.huerta@workforwarriors.org",     "display_name": "Tracey Huerta",     "role": "vsc"},
    {"email": "philip.downs@workforwarriors.org",      "display_name": "Philip Downs",      "role": "vsc"},
    {"email": "damon.oliver@workforwarriors.org",      "display_name": "Damon Oliver",      "role": "vsc"},
    {"email": "chuck.callahan@workforwarriors.org",    "display_name": "Chuck Callahan",    "role": "vsc"},
    {"email": "martin.rivera@workforwarriors.org",     "display_name": "Martin Rivera",     "role": "vsc"},
    {"email": "stevenrosales@workforwarriors.org",     "display_name": "Steven Rosales",    "role": "vsc"},
    {"email": "shayal.prasad@workforwarriors.org",     "display_name": "Shayal Prasad",     "role": "vsc"},
    {"email": "anthony.antonucci@workforwarriors.org", "display_name": "Anthony Antonucci", "role": "admin"},
    {"email": "ryan.mcgrath@workforwarriors.org",      "display_name": "Ryan McGrath",      "role": "admin"},
]
_pw = generate_password_hash("WFW2026!")
seed_users([{**u, "password_hash": _pw} for u in _SEED_USERS])

# Configuration — read directly from .env file to bypass Windows system env var interference
def _read_env_key(key_name):
    """Read a key directly from .env file, bypassing os.environ entirely."""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
    try:
        with open(env_path) as _ef:
            for _line in _ef:
                _line = _line.strip()
                if _line.startswith(f'{key_name}=') and not _line.startswith('#'):
                    return _line.split('=', 1)[1].strip()
    except Exception:
        pass
    return os.getenv(key_name, '')

HF_API_KEY          = _read_env_key("HF_API_KEY")
DEEPSEEK_API_KEY    = _read_env_key("DEEPSEEK_API_KEY")
GEMINI_API_KEY      = _read_env_key("GEMINI_API_KEY")
ANTHROPIC_API_KEY   = _read_env_key("ANTHROPIC_API_KEY")
OUTLOOK_USER        = os.getenv("OUTLOOK_USER", "anthony.antonucci@workforwarriors.org")
OUTLOOK_PASSWORD    = os.getenv("OUTLOOK_PASSWORD", "")
GRAPH_TENANT_ID     = os.getenv("GRAPH_TENANT_ID", "")
GRAPH_CLIENT_ID     = os.getenv("GRAPH_CLIENT_ID", "")
GRAPH_CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET", "")
RYAN_EMAIL          = os.getenv("RYAN_EMAIL", "ryan.mcgrath@workforwarriors.org")
VSC_DISPLAY_NAME    = os.getenv("VSC_DISPLAY_NAME", "VSC")


# ── Microsoft Graph API helpers ───────────────────────────────────────────────

def _graph_get_token():
    """Fetch a client-credentials OAuth2 token for Microsoft Graph."""
    url = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     GRAPH_CLIENT_ID,
        "client_secret": GRAPH_CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def _graph_send_email(to_addr, to_name, subject, body, reply_to=None):
    """Send a plain-text email via Microsoft Graph on behalf of OUTLOOK_USER."""
    token = _graph_get_token()
    message = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": [{"emailAddress": {"address": to_addr, "name": to_name}}],
    }
    if reply_to:
        message["replyTo"] = [{"emailAddress": {"address": reply_to}}]
    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{OUTLOOK_USER}/sendMail",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"message": message, "saveToSentItems": True},
        timeout=30,
    )
    resp.raise_for_status()


def _graph_mark_read(token, msg_id):
    """Mark a mailbox message as read."""
    try:
        requests.patch(
            f"https://graph.microsoft.com/v1.0/users/{OUTLOOK_USER}/messages/{msg_id}",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"isRead": True},
            timeout=15,
        )
    except Exception as e:
        print(f"Graph: mark-read failed: {e}")


def _fetch_attachment_text(token, msg_id):
    """Download the first PDF/DOCX/TXT attachment from a Graph message and return its text."""
    import base64
    import io
    try:
        resp = requests.get(
            f"https://graph.microsoft.com/v1.0/users/{OUTLOOK_USER}/messages/{msg_id}/attachments",
            headers={"Authorization": f"Bearer {token}"},
            timeout=30,
        )
        if not resp.ok:
            return ""
        for att in resp.json().get("value", []):
            name = att.get("name", "")
            b64  = att.get("contentBytes", "")
            if not b64:
                continue
            ext = os.path.splitext(name)[1].lower()
            if ext not in ('.pdf', '.docx', '.txt'):
                continue
            raw = base64.b64decode(b64)
            stream = io.BytesIO(raw)
            if ext == '.pdf':
                try:
                    from pypdf import PdfReader
                    pages = [p.extract_text() for p in PdfReader(stream).pages if p.extract_text()]
                    text = "\n".join(pages).strip()
                    if text:
                        print(f"  Attachment: parsed '{name}' ({len(text)} chars)")
                        return text
                except Exception as e:
                    print(f"  Attachment: PDF parse error on '{name}': {e}")
            elif ext == '.docx':
                try:
                    import docx as _docx
                    doc   = _docx.Document(stream)
                    paras = [p.text for p in doc.paragraphs if p.text.strip()]
                    text  = "\n".join(paras).strip()
                    if text:
                        print(f"  Attachment: parsed '{name}' ({len(text)} chars)")
                        return text
                except Exception as e:
                    print(f"  Attachment: DOCX parse error on '{name}': {e}")
            elif ext == '.txt':
                text = raw.decode('utf-8', errors='replace').strip()
                if text:
                    print(f"  Attachment: parsed '{name}' ({len(text)} chars)")
                    return text
    except Exception as e:
        print(f"  Attachment fetch error: {e}")
    return ""

# Global storage for loaded jobs
JOBS_DB = []
geolocator = Nominatim(user_agent="resume_ai_master", timeout=10)


# Background job-load state
JOB_LOAD_STATUS = {
    'state':    'idle',   # idle | loading | done | error
    'progress': '',
    'total':    0,
    'geocoded': 0,
}

# Pre-load jobs from DB into memory on startup
_db_jobs = load_jobs_to_memory()
if _db_jobs:
    JOBS_DB.extend(_db_jobs)
    print(f"Loaded {len(JOBS_DB)} jobs from database")

# Military rank crosswalk (from Master Prompt)
MOS_CROSSWALK = {
    # ── Army ──────────────────────────────────────────────────────────────────
    "11B": "Infantry → Security Operations / Law Enforcement Support / Protective Services",
    "11C": "Indirect Fire Infantryman → Tactical Operations Coordinator / Field Operations Specialist",
    "12B": "Combat Engineer → Construction Manager / Project Coordinator / Infrastructure Specialist",
    "12N": "Horizontal Construction Engineer → Heavy Equipment Operator / Site Supervisor",
    "13B": "Cannon Crewmember → Heavy Equipment Operator / Ballistics Technician",
    "15T": "UH-60 Helicopter Repairer → Aviation Maintenance Technician / Rotary Wing Mechanic",
    "25B": "IT Specialist → Systems Administrator / Network Administrator / IT Support Specialist",
    "25U": "Signal Support Specialist → IT Support Specialist / Network Technician / Communications Technician",
    "27D": "Paralegal Specialist → Legal Assistant / Compliance Coordinator / Paralegal",
    "31B": "Military Police → Law Enforcement Officer / Security Manager / Loss Prevention Specialist",
    "35F": "Intelligence Analyst → Intelligence Analyst / Data Analyst / Research Analyst",
    "42A": "Human Resources Specialist → HR Coordinator / Talent Acquisition Specialist / Benefits Administrator",
    "56M": "Religious Affairs NCO → Counseling Support Specialist / Community Outreach Coordinator",
    "68D": "Operating Room Specialist → Surgical Technologist / OR Technician / Perioperative Tech",
    "68K": "Medical Laboratory Specialist → Clinical Laboratory Technician / Lab Scientist",
    "68P": "Radiology Specialist → Radiologic Technologist / X-Ray Technician",
    "68W": "Combat Medic Specialist → EMT / Medical Technician / Clinical Support Specialist",
    "79S": "Career Counselor → Career Development Specialist / HR Advisor / Workforce Coach",
    "88M": "Motor Transport Operator → CDL Driver / Logistics Coordinator / Fleet Operator",
    "88N": "Transportation Management Coordinator → Logistics Coordinator / Supply Chain Analyst",
    "89D": "Explosive Ordnance Disposal → EOD Technician / Hazmat Specialist / Safety Officer",
    "91B": "Wheeled Vehicle Mechanic → Automotive Mechanic / Fleet Maintenance Technician",
    "92A": "Automated Logistical Specialist → Supply Chain Coordinator / Inventory Manager / Warehouse Supervisor",
    "92F": "Petroleum Supply Specialist → Fuel Operations Specialist / Logistics Coordinator",
    "92Y": "Unit Supply Specialist → Supply Chain Specialist / Warehouse Coordinator / Inventory Control Specialist",
    # ── Navy Ratings ──────────────────────────────────────────────────────────
    "HM":  "Hospital Corpsman → EMT / Medical Technician / Clinical Support Specialist",
    "IT":  "Information Systems Technician → IT Administrator / Network Specialist / Systems Analyst",
    "YN":  "Yeoman → Administrative Coordinator / Office Manager / Executive Assistant",
    "BM":  "Boatswain's Mate → Maritime Operations Supervisor / Deck Supervisor",
    "EM":  "Electrician's Mate → Electrician / Electrical Systems Technician",
    "EN":  "Engineman → Marine Diesel Mechanic / Engine Technician / Marine Technician",
    "MA":  "Master-at-Arms → Security Officer / Law Enforcement / Access Control Specialist",
    "CS":  "Culinary Specialist → Food Service Manager / Executive Chef / Dining Operations Supervisor",
    "ET":  "Electronics Technician → Electronics Technician / Systems Technician / Avionics Specialist",
    "LS":  "Logistics Specialist → Logistics Coordinator / Supply Chain Specialist / Procurement Analyst",
    "PS":  "Personnel Specialist → HR Coordinator / Benefits Specialist / Personnel Administrator",
    "GM":  "Gunner's Mate → Weapons Systems Technician / Ordnance Specialist / Range Safety Officer",
    # ── Marine Corps MOS ──────────────────────────────────────────────────────
    "0311": "Rifleman → Security Operations Specialist / Law Enforcement Support",
    "0331": "Machine Gunner → Weapons Systems Operator / Range Safety Officer",
    "0811": "Field Artillery Cannoneer → Heavy Equipment Operator / Ballistics Technician",
    "1371": "Combat Engineer → Construction Supervisor / Infrastructure Project Specialist",
    "2651": "Intelligence Analyst → Intelligence Analyst / Geospatial Analyst / Data Analyst",
    "3043": "Supply Administration → Supply Chain Coordinator / Logistics Specialist",
    "3051": "Warehouse Specialist → Warehouse Manager / Inventory Control Specialist",
    "3531": "Motor Vehicle Operator → CDL Driver / Fleet Coordinator / Transportation Specialist",
    "4421": "Legal Services Specialist → Legal Assistant / Paralegal / Compliance Specialist",
    "5811": "Military Police → Law Enforcement Officer / Security Supervisor",
    "6511": "Aviation Mechanic → Aviation Maintenance Technician / Airframe Mechanic",
    # ── Air Force AFSC ────────────────────────────────────────────────────────
    "1A2X1": "Aircraft Loadmaster → Cargo Operations Specialist / Aviation Logistics Coordinator",
    "1N0X1": "Operations Intelligence → Intelligence Analyst / Operations Analyst",
    "2A3X3": "Tactical Aircraft Maintenance → Aircraft Maintenance Technician / Avionics Inspector",
    "2A5X1": "Aerospace Maintenance → Aerospace Maintenance Technician / Quality Assurance Inspector",
    "2E1X1": "Ground Radio Communications → Communications Systems Technician / RF Specialist",
    "3D1X2": "Cyber Systems Operations → Cybersecurity Analyst / IT Systems Administrator",
    "3D0X2": "Cyber Transport Systems → Network Administrator / Systems Engineer",
    "3E0X1": "Electrical Systems → Licensed Electrician / Electrical Systems Supervisor",
    "3E3X1": "Structural → Structural Technician / Construction Specialist / Facilities Engineer",
    "3E4X3": "Utilities Systems → Facilities Maintenance Technician / HVAC Specialist",
    "4N0X1": "Aerospace Medical Technician → EMT / Medical Technician / Clinical Support",
    "4Y0X1": "Dental Assistant → Dental Assistant / Oral Health Technician",
    "6C0X1": "Contracting → Procurement Specialist / Contracts Administrator / Acquisition Analyst",
    # ── Coast Guard ───────────────────────────────────────────────────────────
    "ME":  "Maritime Enforcement Specialist → Law Enforcement Officer / Maritime Security Specialist",
    "MK":  "Machinery Technician → Marine Engineer / Diesel Mechanic / Marine Systems Technician",
}


def _extract_mos_translations(resume_text: str) -> str:
    """Scan resume for known MOS/rating/AFSC codes. Returns formatted translation block or empty string."""
    found = []
    text_upper = resume_text.upper()
    for code, translation in MOS_CROSSWALK.items():
        pattern = rf'\b{re.escape(code)}\b'
        if re.search(pattern, text_upper):
            found.append(f"  {code}: {translation}")
    if not found:
        return ""
    return "MILITARY OCCUPATIONAL SPECIALTY TRANSLATIONS (extracted from resume):\n" + "\n".join(found)


MILITARY_CROSSWALK = """
U.S. Army / Marine Corps (Enlisted & NCOs)
Private (E-1/E-2)  Entry-Level Team Member / Trainee
Private First Class (E-3)  Junior Technician / Operator
Specialist (E-4)  Technical Specialist / Skilled Technician
Corporal (E-4)  Junior Supervisor / Team Lead
Sergeant (E-5)  Operations Supervisor / First-Line Manager
Staff Sergeant (E-6)  Department Supervisor / Section Manager
Sergeant First Class (E-7)  Senior Manager / Operations Coordinator
Master Sergeant (E-8)  Senior Operations Manager
First Sergeant (E-8)  Senior Personnel Manager / HR Operations Lead
Sergeant Major (E-9)  Senior Executive Advisor / Organizational Director
Command Sergeant Major (E-9)  Chief of Staff (Operational Oversight)

U.S. Navy / Coast Guard (Enlisted & NCOs)
Seaman Recruit (E-1)  Entry-Level Crew Member / Trainee
Seaman Apprentice (E-2)  Junior Technician / Operator
Seaman (E-3)  Skilled Operator / Technician
Petty Officer Third Class (E-4)  Team Lead / Junior Supervisor
Petty Officer Second Class (E-5)  Supervisor / Section Lead
Petty Officer First Class (E-6)  Operations Supervisor / Department Lead
Chief Petty Officer (E-7)  Senior Supervisor / Operations Manager
Senior Chief Petty Officer (E-8)  Senior Manager / Division Head
Master Chief Petty Officer (E-9)  Senior Executive Advisor / Organizational Director

U.S. Air Force / Space Force (Enlisted & NCOs)
Airman Basic (E-1)  Entry-Level Trainee
Airman (E-2)  Junior Technician / Operator
Airman First Class (E-3)  Skilled Operator / Technician
Senior Airman (E-4)  Technical Specialist / Assistant Supervisor
Staff Sergeant (E-5)  Operations Supervisor / First-Line Manager
Technical Sergeant (E-6)  Section Manager / Operations Lead
Master Sergeant (E-7)  Department Manager / Senior Supervisor
Senior Master Sergeant (E-8)  Division Manager / Senior Operations Leader
Chief Master Sergeant (E-9)  Senior Executive Advisor / Organizational Director

Warrant Officers (All Branches)
Warrant Officer (WO1)  Technical Specialist / Junior Manager
Chief Warrant Officer 2 (CW2)  Senior Technical Manager / Operations Lead
Chief Warrant Officer 3 (CW3)  Technical Director / Program Manager
Chief Warrant Officer 4 (CW4)  Senior Technical Director / Department Head
Chief Warrant Officer 5 (CW5)  Executive Technical Advisor / Senior Program Director

Commissioned Officers (All Branches)
Second Lieutenant (O-1)  Assistant Manager / Junior Officer / Project Lead
First Lieutenant (O-2)  Manager / Department Lead
Captain (O-3)  Operations Manager / Program Manager
Major (O-4)  Senior Operations Manager / Strategy Lead
Lieutenant Colonel (O-5)  Director / Senior Department Head
Colonel (O-6)  Senior Director / Division Head
Brigadier General (O-7)  Executive Director / VP of Operations
Major General (O-8)  Senior Executive / Regional VP
Lieutenant General (O-9)  Executive Vice President / COO-level Leadership
General (O-10)  C-Suite Executive (CEO/President Equivalent)
"""


# Pre-built coordinate cache for CA cities — loaded from data/ca_cities.py static dict.
# Any city not found here falls through to Nominatim and gets saved to geocode_cache.json
# so subsequent lookups are instant without network calls.
_GEOCODE_CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "geocode_cache.json")

def _load_geocode_cache() -> dict:
    """Load runtime geocode cache from disk, merging with the full hardcoded city list."""
    from data.ca_cities import ALL_CA_COORDS
    cache = {k: tuple(v) for k, v in ALL_CA_COORDS.items()}
    if os.path.exists(_GEOCODE_CACHE_PATH):
        try:
            with open(_GEOCODE_CACHE_PATH, "r") as f:
                saved = json.load(f)
            cache.update({k: tuple(v) for k, v in saved.items()})
        except Exception:
            pass
    return cache

CA_CITY_COORDS = _load_geocode_cache()

def geocode_location(city, state):
    """Convert city/state to lat/lon. Checks cache first, falls back to Nominatim.
    Successful Nominatim results are persisted to geocode_cache.json for future startups."""
    key = city.strip().lower()
    if key in CA_CITY_COORDS:
        return CA_CITY_COORDS[key]
    try:
        location = geolocator.geocode(f"{city}, {state}, USA")
        if location:
            coords = (location.latitude, location.longitude)
            CA_CITY_COORDS[key] = coords
            try:
                existing = {}
                if os.path.exists(_GEOCODE_CACHE_PATH):
                    with open(_GEOCODE_CACHE_PATH, "r") as f:
                        existing = json.load(f)
                existing[key] = list(coords)
                with open(_GEOCODE_CACHE_PATH, "w") as f:
                    json.dump(existing, f, indent=2)
            except Exception:
                pass
            return coords
        return None, None
    except Exception as e:
        print(f"Geocoding error for {city}, {state}: {e}")
        return None, None


def calculate_distance(lat1, lon1, lat2, lon2):
    """Calculate distance between two points in miles"""
    if None in [lat1, lon1, lat2, lon2]:
        return None
    try:
        return geodesic((lat1, lon1), (lat2, lon2)).miles
    except:
        return None


def send_management_email(candidate_name, job_title, status="Complete"):
    """Send minimal summary to management"""
    if not OUTLOOK_PASSWORD:
        print(" Email credentials not configured")
        return False
    
    try:
        msg = MIMEMultipart()
        msg['From'] = OUTLOOK_USER
        msg['To'] = RYAN_EMAIL
        msg['Subject'] = f"Resume Processing - {job_title}"
        
        body = f"""RESUME ANALYSIS SUMMARY

Candidate: {candidate_name}
Position: {job_title}
Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}
Status: {status}

---
Automated notification - Work for Warriors Resume AI
"""
        
        msg.attach(MIMEText(body, 'plain'))
        
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(OUTLOOK_USER, OUTLOOK_PASSWORD)
        server.send_message(msg)
        server.quit()
        
        print(f" Summary sent to {RYAN_EMAIL}")
        return True
        
    except Exception as e:
        print(f" Email failed: {e}")
        return False


def analyze_with_ai(resume_text, candidate_name, candidate_location, target_job, alternative_jobs):
    """AI analysis via DeepSeek direct."""
    try:
        # Build alternative jobs text
        alt_jobs_text = "\n".join([
            f"- {job['Job Title']} at {job['Company Name']} ({job['City']}, {job['State']}) - Distance: {job.get('distance', 'N/A')} miles"
            for job in alternative_jobs[:10]
        ])
        
        prompt = f"""Resume Evaluation, Job Fit Analysis, and Ranked Job Recommendations

You are an expert hiring analyst specializing in resume evaluation and job fit analysis for veterans transitioning to civilian careers.

# CANDIDATE INFORMATION
Candidate: {candidate_name}
Location: {candidate_location}

# TARGET JOB (Position Being Applied For)
Job Title: {target_job.get('Job Title', 'Not specified')}
Company: {target_job.get('Company Name', 'Not specified')}
Location: {target_job.get('City', '')}, {target_job.get('State', '')}
Job Description: {target_job.get('Job Description', 'Not provided')}
Qualifications: {target_job.get('Qualifications', 'Not provided')}
Salary Range: {target_job.get('Salary From', '')} - {target_job.get('Salary To', '')}

# ALTERNATIVE JOB OPTIONS (Pre-filtered by location)
{alt_jobs_text}

# CANDIDATE RESUME
{resume_text}

---

# YOUR TASK

## PART 1: TARGET JOB ANALYSIS

### 1. Match Grade (A+ to F)
Assign letter grade for resume fit to TARGET JOB ONLY.
- A+: 90% alignment
- A: 80-89%
- B: 70-79%
- C: 60-69%
- D: 50-59%
- F: <50%

### 2. Justification (3-5 bullets, 20 words each)
Explain grade based on:
- Match to specific job requirements
- Transferable skills (military  civilian)
- Experience relevance
- Qualifications alignment

### 3. Missing Requirements
List ONLY requirements from target job that are ABSENT from resume.
- Max 7 items
- If none: "None clearly missing"
- DO NOT infer or guess

### 4. Improvement Suggestions
Provide 3 actionable changes (15 words each) to strengthen resume for this specific job.

### 5. Resume Rewrite (350 words, 1 page)
Rewrite resume optimized for TARGET JOB:
- Translate military experience using this crosswalk:

{MILITARY_CROSSWALK}

- Align structure to job requirements
- Quantify achievements
- Remove irrelevant details
- DO NOT fabricate experience
- If data missing, omit

---

## PART 2: ALTERNATIVE JOB RECOMMENDATIONS

From the alternative jobs list, recommend TOP 3 BEST FITS.

For each alternative:
**Job Title:** [title]
**Company:** [company]
**Location:** [city, state]
**Distance:** [miles] miles from candidate
**Fit Score:** X/100
**Why It Fits (2 sentences):** [explanation]
**Salary:** [range if available]

Calculate Fit Score using:
Fit Score = (Resume Match  2) + (Salary Alignment  1.5) + (Proximity Bonus  1)

Where:
- Resume Match = 0-100 based on skills/experience alignment
- Salary Alignment = 0-100 based on candidate level vs job level
- Proximity Bonus = 100 at 0 miles, 75 at 15 miles, 50 at 25 miles, 25 at 50 miles, 0 beyond

Rank alternatives highest to lowest fit score.

---

# OUTPUT FORMAT

## TARGET JOB ANALYSIS
**Job:** [title] at [company]
**Grade:** [Letter]

**Justification:**
- [Bullet 1]
- [Bullet 2]
- [Bullet 3]

**Missing Requirements:**
- [Item 1]
- [Item 2]
...

**Improvement Suggestions:**
1. [Suggestion 1]
2. [Suggestion 2]
3. [Suggestion 3]

**Resume Rewrite:**
[Full rewritten resume]

---

## RECOMMENDED ALTERNATIVES

**Alternative 1:**
Job: [title] at [company]
Location: [city, state]
Distance: [X] miles
Fit Score: X/100
Why: [Explanation]
Salary: [range]

**Alternative 2:**
[Same format]

**Alternative 3:**
[Same format]

---

# RULES
- NO greetings or commentary
- NO fabricated data
- Military ranks  civilian equivalents using provided crosswalk
- Preserve only information from original resume
- If unsure, omit rather than guess
"""
        
        messages = [{"role": "user", "content": prompt}]
        print(" Sending to DeepSeek...")
        content, used = _call_deepseek(messages, max_tokens=4000)
        if content:
            print(" Analysis complete")
            return content
        return "Analysis error: DeepSeek unavailable"

    except Exception as e:
        return f"Analysis error: {str(e)}"


def quick_extract_location(resume_text):
    """Extract city/state from resume header using regex. Returns (city, state) or (None, None)."""
    header = resume_text[:600]
    m = re.search(r'\b([A-Z][a-zA-Z\s]{2,20}),\s*([A-Z]{2})\b', header)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return None, None


def _extract_contact_fields(body_text, resume_text):
    """Search every available text source for every contact field.
    No source is exclusive to any field — all patterns run on all sources.
    Returns dict: phones (str), address (str), city (str), state (str).
    """
    sources = [s for s in [body_text or '', resume_text or ''] if s.strip()]

    PHONE_LABELED = re.compile(
        r'(?:Phone|Tel(?:ephone)?|Mobile|Cell|Direct)[:\s]+(\+?1?[\s.\-]?\(?\d{3}\)?[\s.\-]\d{3}[\s.\-]\d{4})',
        re.IGNORECASE
    )
    PHONE_BARE = re.compile(r'(?<!\d)(\+?1?\s?\(?\d{3}\)?[\s.\-]\d{3}[\s.\-]\d{4})(?!\d)')

    STREET_RE = re.compile(
        r'(\d{1,5}\s+[A-Za-z0-9\.\s]{3,40}'
        r'(?:Street|St|Avenue|Ave|Drive|Dr|Road|Rd|Lane|Ln|Boulevard|Blvd|Way|'
        r'Court|Ct|Place|Pl|Circle|Cir|Terrace|Ter|Highway|Hwy|Parkway|Pkwy)'
        r'\.?(?:\s*(?:Apt|Suite|Ste|Unit|#)\s*[\w-]+)?)',
        re.IGNORECASE
    )

    CITY_STATE_RE = re.compile(r'\b([A-Z][a-zA-Z\s]{2,20}),\s*([A-Z]{2})(?:\s+\d{5})?\b')

    all_phones  = []
    seen_digits = set()
    address     = ''
    city        = ''
    state       = ''

    for text in sources:
        # Phones — labeled pattern (full text)
        for m in PHONE_LABELED.finditer(text):
            digits = re.sub(r'\D', '', m.group(1))
            if len(digits) >= 10 and digits not in seen_digits and len(set(digits)) > 3:
                seen_digits.add(digits)
                all_phones.append(m.group(1).strip())

        # Phones — bare pattern (header only to avoid ZIP/date false positives)
        for m in PHONE_BARE.finditer(text[:1000]):
            digits = re.sub(r'\D', '', m.group(1))
            if len(digits) >= 10 and digits not in seen_digits and len(set(digits)) > 3:
                seen_digits.add(digits)
                all_phones.append(m.group(1).strip())

        # Street address
        if not address:
            m = STREET_RE.search(text[:2000])
            if m:
                address = m.group(1).strip()

        # City / State
        if not city:
            m = CITY_STATE_RE.search(text[:1500])
            if m:
                city  = m.group(1).strip()
                state = m.group(2).strip()

    return {
        'phones':  ' / '.join(all_phones),
        'address': address,
        'city':    city,
        'state':   state,
    }


def _extract_military_profile(resume_text):
    """Extract branch, rank, MOS, and years_served from resume text.
    Returns dict — all keys present, values may be empty string.
    """
    # Use full text but focus early sections for header fields
    text      = resume_text[:4000]
    text_up   = text.upper()
    result    = {'branch': '', 'rank': '', 'mos': '', 'years_served': ''}

    # ── Branch ────────────────────────────────────────────────────────────────
    for pattern, label in [
        (r'\bU\.?S\.?\s*Army\b|\bUnited\s+States\s+Army\b',                'Army'),
        (r'\bU\.?S\.?\s*Marine\s+Corps\b|\bUnited\s+States\s+Marine\b|\bUSMC\b|\bMarines\b', 'Marine Corps'),
        (r'\bU\.?S\.?\s*Navy\b|\bUnited\s+States\s+Navy\b|\bUSN\b',        'Navy'),
        (r'\bU\.?S\.?\s*Air\s+Force\b|\bUnited\s+States\s+Air\s+Force\b|\bUSAF\b', 'Air Force'),
        (r'\bU\.?S\.?\s*Coast\s+Guard\b|\bUnited\s+States\s+Coast\s+Guard\b|\bUSCG\b', 'Coast Guard'),
        (r'\bU\.?S\.?\s*Space\s+Force\b|\bUnited\s+States\s+Space\s+Force\b|\bUSSF\b', 'Space Force'),
        (r'\bArmy\s+National\s+Guard\b|\bAir\s+National\s+Guard\b|\bNational\s+Guard\b', 'National Guard'),
        (r'\bArmy\s+Reserve\b|\bNaval\s+Reserve\b|\bAir\s+Force\s+Reserve\b|\bMarine\s+Reserve\b', 'Reserve'),
    ]:
        if re.search(pattern, text, re.IGNORECASE):
            result['branch'] = label
            break

    # ── Rank — spelled-out (specific before general) ─────────────────────────
    SPELLED_RANKS = [
        # Warrant Officers
        'Chief Warrant Officer 5', 'Chief Warrant Officer 4', 'Chief Warrant Officer 3',
        'Chief Warrant Officer 2', 'Chief Warrant Officer',
        # Army/Marines NCO (highest-first to avoid partial matches)
        'Command Sergeant Major', 'Sergeant Major of the Army',
        'Sergeant Major', 'First Sergeant', 'Master Sergeant',
        'Sergeant First Class', 'Staff Sergeant', 'Sergeant', 'Corporal', 'Specialist',
        'Private First Class', 'Private',
        # Navy/CG
        'Master Chief Petty Officer', 'Senior Chief Petty Officer', 'Chief Petty Officer',
        'Petty Officer First Class', 'Petty Officer Second Class', 'Petty Officer Third Class',
        'Seaman', 'Seaman Apprentice', 'Seaman Recruit',
        # Air/Space Force
        'Chief Master Sergeant', 'Senior Master Sergeant',
        'Technical Sergeant', 'Senior Airman', 'Airman First Class', 'Airman Basic', 'Airman',
        # Officers
        'General of the Army', 'General', 'Lieutenant General', 'Major General',
        'Brigadier General', 'Colonel', 'Lieutenant Colonel',
        'Major', 'Captain', 'First Lieutenant', 'Second Lieutenant',
    ]
    for r in SPELLED_RANKS:
        if re.search(rf'\b{re.escape(r)}\b', text, re.IGNORECASE):
            result['rank'] = r
            break

    # Fallback: common abbreviations (word-boundary, uppercase only to reduce false positives)
    if not result['rank']:
        for abbrev, full in [
            ('CSM', 'Command Sergeant Major'), ('SGM', 'Sergeant Major'),
            ('1SG', 'First Sergeant'),         ('MSG', 'Master Sergeant'),
            ('SFC', 'Sergeant First Class'),   ('SSG', 'Staff Sergeant'),
            ('SGT', 'Sergeant'),               ('CPL', 'Corporal'),
            ('SPC', 'Specialist'),             ('PFC', 'Private First Class'),
            ('CW5', 'Chief Warrant Officer 5'), ('CW4', 'Chief Warrant Officer 4'),
            ('CW3', 'Chief Warrant Officer 3'), ('CW2', 'Chief Warrant Officer 2'),
            ('WO1', 'Warrant Officer'),
            ('MCPO', 'Master Chief Petty Officer'), ('SCPO', 'Senior Chief Petty Officer'),
            ('CPO', 'Chief Petty Officer'),
            ('GEN', 'General'), ('LTG', 'Lieutenant General'), ('MG', 'Major General'),
            ('BG', 'Brigadier General'), ('COL', 'Colonel'), ('LTC', 'Lieutenant Colonel'),
            ('MAJ', 'Major'), ('CPT', 'Captain'), ('CAPT', 'Captain'),
            ('1LT', 'First Lieutenant'), ('2LT', 'Second Lieutenant'),
            ('CMSGT', 'Chief Master Sergeant'), ('SMSGT', 'Senior Master Sergeant'),
            ('MSGT', 'Master Sergeant'), ('TSGT', 'Technical Sergeant'),
            ('SSGT', 'Staff Sergeant'), ('SRA', 'Senior Airman'), ('A1C', 'Airman First Class'),
        ]:
            if re.search(rf'\b{re.escape(abbrev)}\b', text_up):
                result['rank'] = full
                break

    # ── MOS — scan for known codes ────────────────────────────────────────────
    found_codes = [code for code in MOS_CROSSWALK if re.search(rf'\b{re.escape(code)}\b', text_up)]
    if found_codes:
        result['mos'] = ', '.join(found_codes[:3])

    # ── Years served ──────────────────────────────────────────────────────────
    m = re.search(
        r'(\d{1,2})\s+years?\s+(?:of\s+)?(?:military\s+|active[\s-]+duty\s+)?(?:service|experience)',
        text, re.IGNORECASE
    )
    if m:
        result['years_served'] = f"{m.group(1)} years"
    else:
        # Date range near a military section (e.g. "U.S. Army  2005 – 2019")
        mil_m = re.search(
            r'(?:Army|Navy|Marine|Air Force|Coast Guard|Military|Service).*?(\b(?:19|20)\d{2}\b).*?(\b(?:19|20)\d{2}\b|[Pp]resent)',
            text, re.IGNORECASE | re.DOTALL
        )
        if mil_m:
            try:
                start = int(mil_m.group(1))
                end_raw = mil_m.group(2)
                end = 2026 if re.match(r'[Pp]resent', end_raw) else int(end_raw)
                span = end - start
                if 1 <= span <= 40:
                    result['years_served'] = f"{span} years"
            except (ValueError, TypeError):
                pass

    return result


# Degree hierarchy: higher index = higher credential
_DEGREE_LEVELS = [
    (3, 'HS Diploma/GED', [
        r'high school diploma', r'high school degree', r'hs diploma', r'h\.s\. diploma',
        r'\bged\b', r'g\.e\.d', r'high school graduate', r'secondary education',
    ]),
    (4, "Associate's", [
        r"associate'?s?\b", r'\ba\.s\.', r'\ba\.a\.', r'associate of \w', r'associate degree',
        r'2.year degree', r'two.year degree',
    ]),
    (5, "Bachelor's", [
        r"bachelor'?s?\b", r'\bb\.s\.', r'\bb\.a\.', r'\bbba\b', r'\bbfa\b',
        r'bachelor of \w', r'undergraduate degree', r'4.year degree', r'four.year degree',
    ]),
    (6, "Master's", [
        r"master'?s?\b", r'\bm\.s\.', r'\bm\.a\.', r'\bmba\b', r'\bmpa\b',
        r'\bmsw\b', r'\bmph\b', r'\bm\.ed\.?', r'master of \w', r'graduate degree',
        r'post.?graduate', r'postgraduate',
    ]),
    (7, 'PhD/Doctorate', [
        r'\bph\.?d\.?\b', r'\bdoctorate\b', r'\bdoctoral\b', r'\bd\.phil\.?\b',
    ]),
]


def _highest_degree_in(text):
    """Return (level, label) of highest degree credential found in text. Level 0 = none found."""
    t = text.lower()
    for level, label, patterns in reversed(_DEGREE_LEVELS):
        for p in patterns:
            if re.search(p, t):
                return level, label
    return 0, None


def check_education_requirement(resume_text, qualifications_text):
    """
    Deterministically compare the degree required by the job against the highest
    degree held by the candidate. Returns an injection block string for AI prompts,
    or None if no degree requirement is detected.
    """
    required_level, required_label = _highest_degree_in(qualifications_text or '')
    if required_level == 0:
        return None  # no degree requirement detected — nothing to inject

    resume_level, resume_label = _highest_degree_in(resume_text or '')

    if resume_level == 0:
        verdict = 'UNVERIFIED'
        detail  = (
            f"No degree credential detected in resume text. "
            f"If the resume shows U.S. military service the HS/GED rule applies. "
            f"Otherwise treat as NOT_PROVIDED."
        )
    elif resume_level >= required_level:
        verdict = 'CONFIRMED'
        detail  = (
            f"Candidate holds {resume_label} (level {resume_level}), "
            f"which meets or exceeds the required {required_label} (level {required_level}). "
            f"Mark the education requirement CONFIRMED — do not override this."
        )
    else:
        verdict = 'POTENTIAL_GAP'
        detail  = (
            f"Candidate's highest detected credential is {resume_label} (level {resume_level}), "
            f"below the required {required_label} (level {required_level}). "
            f"Treat as NOT_PROVIDED unless the resume body contains evidence otherwise."
        )

    return (
        f"DETERMINISTIC EDUCATION PRE-CHECK (Python-verified — ground truth, do not override):\n"
        f"  Required degree:   {required_label}\n"
        f"  Candidate holds:   {resume_label or 'not detected'}\n"
        f"  Verdict:           {verdict}\n"
        f"  {detail}\n"
    )


_MODEL_DISPLAY = {
    "deepseek-chat":                    "DeepSeek",
    "anthropic/claude-haiku-4-5-20251001": "Claude Haiku",
    "gemini/gemini-2.0-flash":          "Gemini",
}

def _model_name(model_id):
    return _MODEL_DISPLAY.get(model_id, model_id)


import logging as _logging
_dbg_logger = _logging.getLogger('awis.debug')
if not _dbg_logger.handlers:
    _dbg_handler = _logging.FileHandler('webhook_debug.log', encoding='utf-8')
    _dbg_handler.setFormatter(_logging.Formatter('%(asctime)s %(message)s'))
    _dbg_logger.addHandler(_dbg_handler)
    _dbg_logger.setLevel(_logging.DEBUG)

def _dbg_log(msg):
    _dbg_logger.debug(msg)


_MODEL_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'model_debug.log')

def _mlog(msg):
    line = f"{datetime.now().strftime('%H:%M:%S')} {msg}"
    _print(line, flush=True)
    try:
        with open(_MODEL_LOG, 'a', encoding='utf-8') as f:
            f.write(line + "\n")
    except Exception as _mlog_err:
        _print(f"[MLOG WRITE FAIL] {_mlog_err}", flush=True)

def _call_chat_api(base_url, api_key, model, messages, max_tokens=1500, timeout=180):
    """Call a chat completions API endpoint."""
    if not api_key:
        return None, None
    import time
    for attempt in range(3):
        try:
            resp = requests.post(
                f"{base_url.rstrip('/')}/chat/completions",
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                json={"model": model, "messages": messages, "max_tokens": max_tokens},
                timeout=timeout,
            )
            if resp.status_code == 200:
                choices = resp.json().get("choices", [])
                if choices:
                    text = choices[0].get("message", {}).get("content", "")
                    if text and text.strip():
                        _mlog(f"OK    {model} (direct)  {len(text)} chars")
                        return text.strip(), model
                _mlog(f"EMPTY {model}  200 no content: {resp.text[:200]}")
                return None, None
            if resp.status_code == 429:
                wait = 5 * (2 ** attempt)
                _mlog(f"WAIT  {model}  429 retry in {wait}s")
                time.sleep(wait)
                continue
            _mlog(f"ERR   {model}  HTTP {resp.status_code}: {resp.text[:300]}")
            return None, None
        except Exception as e:
            _mlog(f"EXC   {model}  {e}")
            return None, None
    return None, None



def _call_deepseek(messages, max_tokens=1500, model="deepseek-chat"):
    """Call DeepSeek direct API."""
    key = _read_env_key("DEEPSEEK_API_KEY")
    _mlog(f"DS    key={'found' if key else 'MISSING'}  model={model}")
    if not key:
        return None, None
    _mlog(f"CALL  deepseek-direct/{model}  max_tokens={max_tokens}")
    content, used = _call_chat_api(
        "https://api.deepseek.com", key, model, messages, max_tokens=max_tokens, timeout=120
    )
    if content:
        _mlog(f"OK    deepseek-direct/{model}  {len(content)} chars")
        return content, f"deepseek-direct/{model}"
    _mlog(f"FAIL  deepseek-direct/{model}")
    return None, None


def _call_gemini(messages, max_tokens=900, model="gemini-2.0-flash"):
    """Call Gemini API directly. Used only on Opportunities track (resume header only — no full PII)."""
    key = _read_env_key("GEMINI_API_KEY")
    _mlog(f"GEMINI  key={'found' if key else 'MISSING'}  model={model}")
    if not key:
        return None, None
    try:
        text_in = "\n".join(m.get("content", "") for m in messages if m.get("role") == "user")
        resp = requests.post(
            f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}",
            headers={"Content-Type": "application/json"},
            json={"contents": [{"parts": [{"text": text_in}]}],
                  "generationConfig": {"maxOutputTokens": max_tokens}},
            timeout=120,
        )
        if resp.status_code == 200:
            candidates = resp.json().get("candidates", [])
            if candidates:
                text = candidates[0].get("content", {}).get("parts", [{}])[0].get("text", "")
                if text and text.strip():
                    _mlog(f"GEMINI OK  {len(text)} chars")
                    return text.strip(), f"gemini/{model}"
            _mlog(f"GEMINI EMPTY  200 no content")
            return None, None
        _mlog(f"GEMINI ERR  HTTP {resp.status_code}: {resp.text[:200]}")
        return None, None
    except Exception as e:
        _mlog(f"GEMINI EXC  {e}")
        return None, None


_HF_FALLBACK_MODELS = [
    "meta-llama/Meta-Llama-3-8B-Instruct",
    "mistralai/Mistral-7B-Instruct-v0.3",
    "Qwen/Qwen2.5-7B-Instruct",
    "HuggingFaceH4/zephyr-7b-beta",
]


def _call_claude(messages, max_tokens=1500, model="claude-haiku-4-5-20251001"):
    """Call Anthropic API directly."""
    key = _read_env_key("ANTHROPIC_API_KEY")
    _mlog(f"CLAUDE  key={'found' if key else 'MISSING'}  model={model}")
    if not key:
        return None, None
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": key,
                "anthropic-version": "2023-06-01",
                "Content-Type": "application/json",
            },
            json={
                "model": model,
                "max_tokens": max_tokens,
                "messages": [m for m in messages if m.get("role") in ("user", "assistant")],
            },
            timeout=120,
        )
        if resp.status_code == 200:
            content_blocks = resp.json().get("content", [])
            text = "".join(b.get("text", "") for b in content_blocks if b.get("type") == "text")
            if text and text.strip():
                _mlog(f"CLAUDE OK  {len(text)} chars")
                return text.strip(), f"anthropic/{model}"
            _mlog(f"CLAUDE EMPTY  200 no content")
            return None, None
        _mlog(f"CLAUDE ERR  HTTP {resp.status_code}: {resp.text[:300]}")
        return None, None
    except Exception as e:
        _mlog(f"CLAUDE EXC  {e}")
        return None, None


def _call_with_fallback(messages, max_tokens=1500):
    """Try DeepSeek direct first, then HuggingFace."""
    content, used = _call_deepseek(messages, max_tokens=max_tokens)
    if content:
        return content, used
    if HF_API_KEY:
        hf_max = min(max_tokens, 1024)
        for hf_model in _HF_FALLBACK_MODELS:
            _mlog(f"HF    {hf_model.split('/')[-1]}  max_tokens={hf_max}")
            content, used = _call_hf_model(hf_model, messages, max_tokens=hf_max)
            if content:
                _mlog(f"HF OK {hf_model.split('/')[-1]}  {len(content)} chars")
                return content, used
            _mlog(f"HF NO {hf_model.split('/')[-1]}")
    return None, None


def _call_hf_model(model, messages, max_tokens=1500, timeout=180):
    """Call a Hugging Face Inference API model. Returns (content, display_name) or (None, None)."""
    if not HF_API_KEY:
        return None, None
    try:
        resp = requests.post(
            f"https://api-inference.huggingface.co/models/{model}/v1/chat/completions",
            headers={"Authorization": f"Bearer {HF_API_KEY}", "Content-Type": "application/json"},
            json={"model": model, "messages": messages, "max_tokens": max_tokens},
            timeout=timeout,
        )
        if resp.status_code == 200:
            choices = resp.json().get('choices', [])
            if choices:
                raw = choices[0].get('message', {}).get('content', '')
                if raw and raw.strip():
                    display = model.split('/')[-1].replace('-Instruct', '').replace('-instruct', '')
                    return raw.strip(), display
        _mlog(f"HF ERR {model.split('/')[-1]}  HTTP {resp.status_code}: {resp.text[:200]}")
        return None, None
    except Exception as e:
        _mlog(f"HF EXC {model.split('/')[-1]}  {e}")
        return None, None


def analyze_for_vsc(resume_text, candidate_name, target_job, alternative_jobs):
    _mlog("=== ANALYSIS START — DeepSeek direct only ===")
    """
    AWIS Council v1.0 — 4 specialized tracks run in parallel, assembled into one report.
    Returns (assembled_text, track_credits) or (None, None).
    """
    alt_jobs_text = "\n".join([
        f"- {job.get('Job Title', '')} at {job.get('Company Name', '')} "
        f"({job.get('City', '')}, {job.get('State', '')})  {job.get('distance', 'N/A')} miles"
        for job in alternative_jobs[:30]
    ])

    job_title      = target_job.get('Job Title', 'Not specified')
    company        = target_job.get('Company Name', 'Not specified')
    job_city       = target_job.get('City', '')
    job_state      = target_job.get('State', '')
    job_desc       = target_job.get('Job Description', 'Not provided')
    qualifications = target_job.get('Qualifications', 'Not provided')

    # Deterministic education check — injected into T1 and T3 as ground truth
    edu_pre_check = check_education_requirement(resume_text, qualifications)
    edu_block = f"\n{edu_pre_check}\n" if edu_pre_check else ""

    # ── Track 1: Eligibility ──────────────────────────────────────────────────
    t1 = f"""You are a veteran employment eligibility analyst for Work for Warriors.

RULES: Do NOT fabricate qualifications. Do NOT assume missing data. Job description is the source of truth.

VETERAN EDUCATION RULE — CRITICAL:
U.S. military enlistment legally requires a minimum of a high school diploma or GED.
If the resume shows ANY U.S. military service (any branch, any rank, any MOS/rate), you MUST mark
"High School diploma", "GED", "H.S. diploma", or any equivalent education requirement as CONFIRMED.
Do NOT mark it NOT_PROVIDED simply because education is not listed in a separate section.

AGE RULE — CRITICAL:
Age is a legally protected category. NEVER flag any minimum age requirement as NOT_PROVIDED or NOT_MET.
Always mark it CONFIRMED and do not surface it to the VSC.

EDUCATION HIERARCHY RULE — CRITICAL:
A higher degree always satisfies a lower degree requirement. Use this hierarchy (highest to lowest):
PhD / Doctorate → Master's (MA/MS/MBA/etc.) → Bachelor's (BA/BS/etc.) → Associate's (AA/AS) → HS Diploma/GED
If the resume shows ANY degree at or above the required level, mark that education requirement as CONFIRMED.
Example: job requires BA → candidate has PhD → mark CONFIRMED. Never mark a degree requirement NOT_MET or
NOT_PROVIDED when the candidate holds a higher credential.

STANDARD SKILLS RULE:
Common workplace tools are implied by relevant work experience. If the resume shows administrative,
clerical, office coordination, management, or customer service experience, mark the following as CONFIRMED
without requiring explicit mention: MS Office / Microsoft Office (Word, Excel, Outlook, PowerPoint),
basic computer use, data entry, email, internet. Do NOT flag these as NOT_PROVIDED for experienced workers.

CANDIDATE: {candidate_name}
JOB: {job_title} at {company} ({job_city}, {job_state})
JOB DESCRIPTION: {job_desc}
REQUIRED QUALIFICATIONS: {qualifications}
RESUME:
{resume_text}
{edu_block}
ELIGIBILITY RULES:
SUITABLE = all requirements CONFIRMED → recommend submission
PENDING REVIEW = any requirement NOT_PROVIDED or NOT_MET → VSC must verify before deciding
NOT SUITABLE — FINAL = candidate is actively, provably disqualified — use ONLY when the resume
  explicitly contradicts a hard requirement (e.g. requires CA license, resume states licensed in TX only;
  requires 5 years experience, resume shows 6 months). A missing item alone is NEVER grounds for
  NOT SUITABLE — FINAL. When in doubt, use PENDING REVIEW.

GRADING RULES:
TRUE_FIT_GRADE = pure resume-to-job match (A+/A/A-/B+/B/B-/C/D/F)
PRIMARY_GRADE: SUITABLE→same as TRUE FIT | PENDING REVIEW→cap at B | NOT SUITABLE FINAL→F
SHADOW_GRADE: what candidate scores if all gaps confirmed | N/A if SUITABLE

OUTPUT — EXACT FORMAT ONLY:
ELIGIBILITY: [SUITABLE | PENDING REVIEW | NOT SUITABLE — FINAL]

REQUIREMENTS:
[Requirement] — [CONFIRMED | NOT_PROVIDED | NOT_MET] — [brief note]

PRIMARY_GRADE: [grade]
TRUE_FIT_GRADE: [grade]
SHADOW_GRADE: [grade or N/A]

CLASSIFICATION RULES — FOLLOW EXACTLY:
Direct Fit   = ELIGIBLE is SUITABLE AND candidate experience directly aligns to this role → submit
Redirect     = ELIGIBLE is SUITABLE but candidate is clearly better suited to a DIFFERENT job category — NOT for minor gaps, NOT for near-matches, NOT for B-range grades
Develop      = ELIGIBLE is PENDING REVIEW AND there are fixable skill gaps → coach before submitting
Non-viable   = ELIGIBLE is NOT SUITABLE — FINAL only

IMPORTANT: If ELIGIBILITY = SUITABLE, use Direct Fit in almost all cases. Only use Redirect when the
role type itself is genuinely wrong for this candidate (e.g. a pilot applying for an accounting role).
A B- grade, a missing soft skill, or a single unconfirmed requirement is NOT grounds for Redirect.
When in doubt, choose Direct Fit and let the employer decide.
CLASSIFICATION: [Direct Fit | Redirect | Develop | Non-viable]
END"""

    # ── Track 2: Development (split into T2a + T2b to stay under per-request token cap) ──
    mos_translations = _extract_mos_translations(resume_text)
    mos_block = f"\n{mos_translations}\n" if mos_translations else ""

    _t2_context = f"""CANDIDATE: {candidate_name}
JOB: {job_title} at {company}
JOB DESCRIPTION:
{job_desc}

REQUIRED QUALIFICATIONS:
{qualifications}

RESUME:
{resume_text}

RANK-TO-CIVILIAN CROSSWALK:
{MILITARY_CROSSWALK}{mos_block}"""

    t2a = f"""You are a veteran career coach for Work for Warriors.

{_t2_context}

RULES:
- No fabrication. Only use experience and credentials already in the resume.
- Translate military jargon to civilian language.

Identify 10-15 key terms the employer used in the job description (skills, certifications, tools, action verbs — their exact words).

OUTPUT — EXACT FORMAT ONLY:
ATS_KEYWORDS:
[comma-separated list of 10-15 key terms from the job description]

JUSTIFICATION:
- [how candidate's background aligns to this role]
- [transferable skills or military experience mapping to requirements]
- [any risk, gap, or notable strength for the VSC]

IMPROVEMENTS:
- [specific ATS language change — original phrase → improved phrase]
- [specific ATS language change]
- [specific ATS language change]
END"""

    t2b = f"""You are a veteran resume writer for Work for Warriors.

{_t2_context}

RULES:
- No fabrication. Never invent experience, credentials, or metrics not in the original resume.
- The Professional Summary must ONLY claim expertise explicitly demonstrated in the resume. Never add domain names, system names, or specializations (e.g. HCM, ERP, Workday, SAP) that do not appear in the original resume.
- Quantify only what the resume already supports. If a number isn't there, don't add one.
- ALWAYS include the candidate's most recent/current employer as the first entry under Professional Experience. Never drop it.
- Translate all military jargon to civilian language (e.g. "deployed" → "mobilized", "TDY" → "temporary assignment").
- Use the rank crosswalk and MOS translations to reframe military titles and roles.
- Keep the resume concise. Prioritize the most recent 10 years and most relevant experience. Do not cut the current employer to save space.
- Integrate the employer's own key terms naturally where the candidate's experience already supports them.

OUTPUT — EXACT FORMAT ONLY:
RESUME_REWRITE:
[Full civilian-language, ATS-optimized resume. Sections: Professional Summary | Core Competencies | Professional Experience | Military Service | Education & Certifications. No invented experience.]
END"""

    # ── Track 3: Verification (GPT-OSS 120B → Nemotron) ──────────────────────
    t3 = f"""You are a veteran services coordinator for Work for Warriors. Your job is to identify true gaps and write VSC outreach questions.

CANDIDATE: {candidate_name}
JOB: {job_title} at {company}
REQUIRED QUALIFICATIONS: {qualifications}
RESUME:
{resume_text}
{edu_block}
RULES — READ CAREFULLY:
- A requirement is CONFIRMED if the resume shows it directly OR through equivalent demonstrated experience.
  Example: "sensitivity to behavioral health populations" is CONFIRMED if the resume shows caregiving for mentally ill patients.
- A requirement is NOT_PROVIDED only if there is zero evidence — no experience, no credential, no equivalent anywhere in the resume.
- A requirement is NOT_MET only if the resume actively contradicts it or falls clearly short.
- Do NOT list a requirement as missing if it was met by equivalent experience, even if not stated word-for-word.
- Credentials (GED, specific license, certification) that are completely absent from the resume are legitimate gaps.
- Keep the list short — only genuine gaps, not every possible edge case.
- For VERIFICATION_REQUIRED: write one short yes/no question per gap. Questions go to the VSC, not the candidate.
- If the candidate meets all requirements, write N/A for both sections.

VETERAN EDUCATION RULE — CRITICAL:
U.S. military enlistment legally requires a high school diploma or GED.
If the resume shows ANY U.S. military service, do NOT list "High School diploma", "GED", or equivalent
as a missing requirement. It is confirmed by the fact of their service.

AGE RULE — CRITICAL:
Age is a legally protected category. NEVER list any minimum age requirement as missing or ask the VSC to verify age.
Always treat age requirements as CONFIRMED.

EDUCATION HIERARCHY RULE — CRITICAL:
A higher degree always satisfies a lower degree requirement (PhD → Master's → Bachelor's → Associate's → HS Diploma).
NEVER list a degree requirement as missing if the candidate holds a higher credential.
Example: job requires BA, candidate has PhD → do NOT list BA as a gap.

STANDARD SKILLS RULE:
Do NOT list MS Office, Microsoft Office, basic computer use, data entry, email, or internet as missing
requirements if the resume shows any administrative, clerical, office, management, or customer service
experience. These are implied baseline skills for experienced workers.

OUTPUT — EXACT FORMAT ONLY:
MISSING_REQUIREMENTS:
- [genuine gap — absent credential, license, or qualification with zero resume evidence]

VERIFICATION_REQUIRED:
- [direct yes/no question for VSC to ask candidate about each gap above]
END"""

    # ── Track 4: Opportunities (GPT-OSS 120B → Nemotron) ─────────────────────
    # Pull a brief career summary from the resume header for relevance filtering
    resume_header = resume_text[:800]
    t4 = f"""You are a veteran job placement specialist for Work for Warriors.

CANDIDATE: {candidate_name}
RESUME HEADER (for career context):
{resume_header}

TARGET JOB: {job_title} at {company}
TARGET JOB LOCATION: {job_city}, {job_state}

NEARBY JOBS LIST:
{alt_jobs_text}

TASK:
1. Extract candidate city/state from the resume header.
2. From the nearby jobs list, select the 3 best ALTERNATIVE jobs. Each alternative must:
   - Match the candidate's background, skills, and experience level as shown in the resume
   - Exclude the target job itself ({job_title} at {company})
   - Be within a reasonable commute distance
   - Represent a realistic fit — not a career change into an unrelated field
3. Score each on fit (skills/experience match × 70% + proximity × 30%).

Distance tiers: IDEAL under 15mi | ACCEPTABLE 15-30mi | BORDERLINE 30-50mi

OUTPUT — EXACT FORMAT ONLY:
LOCATION:
Candidate: [city], [state]
Job: {job_city}, {job_state}
Distance: [X miles]
Tier: [IDEAL | ACCEPTABLE | BORDERLINE]

ALTERNATIVE_1:
Title: [job title]
Company: [company name]
City: [city], [state]
Distance: [X] miles
Score: [X]/100
Why: [1-2 sentences explaining why this fits the candidate's background]

ALTERNATIVE_2:
Title: [job title]
Company: [company name]
City: [city], [state]
Distance: [X] miles
Score: [X]/100
Why: [1-2 sentences explaining why this fits the candidate's background]

ALTERNATIVE_3:
Title: [job title]
Company: [company name]
City: [city], [state]
Distance: [X] miles
Score: [X]/100
Why: [1-2 sentences explaining why this fits the candidate's background]
END"""

    track_results = {}

    def run_track_split(track_id, prompt_a, max_a, prompt_b, max_b):
        _print(f"  [TRACK] {track_id} START (split)", flush=True)
        try:
            msgs = lambda p: [{"role": "user", "content": p}]
            content_a, used_a = _call_with_fallback(msgs(prompt_a), max_tokens=max_a)
            content_b, used_b = _call_with_fallback(msgs(prompt_b), max_tokens=max_b)
            combined = "\n\n".join(filter(None, [content_a, content_b]))
            track_results[track_id] = (combined if combined else None, used_a or used_b)
            _print(f"  [TRACK] {track_id} DONE: {'OK' if combined else 'None'}", flush=True)
        except Exception as _te:
            import traceback as _tb
            _print(f"  [TRACK] {track_id} EXCEPT: {_te}", flush=True)
            _print(_tb.format_exc(), flush=True)
            track_results[track_id] = (None, None)

    def run_track(track_id, prompt, max_tokens, primary=None):
        _print(f"  [TRACK] {track_id} START", flush=True)
        try:
            msgs = [{"role": "user", "content": prompt}]
            if primary is not None:
                content, used = primary(msgs, max_tokens=max_tokens)
                if not content:
                    content, used = _call_with_fallback(msgs, max_tokens=max_tokens)
            else:
                content, used = _call_with_fallback(msgs, max_tokens=max_tokens)
            track_results[track_id] = (content, used)
            _print(f"  [TRACK] {track_id} DONE: {'OK' if content else 'None'}", flush=True)
        except Exception as _te:
            import traceback as _tb
            _print(f"  [TRACK] {track_id} EXCEPT: {_te}", flush=True)
            _print(_tb.format_exc(), flush=True)
            track_results[track_id] = (None, None)

    def run_track_hf(track_id, hf_model, prompt, max_tokens):
        try:
            content, used = _call_hf_model(
                hf_model, [{"role": "user", "content": prompt}], max_tokens=max_tokens
            )
            track_results[track_id] = (content, used)
        except Exception as e:
            import traceback
            print(f"  Council track_hf '{track_id}' EXCEPTION: {e}")
            traceback.print_exc()
            track_results[track_id] = (None, None)

    print(f" Running VSC council analysis for {candidate_name}...")

    def run_track_opportunities(prompt, max_tokens):
        _print(f"  [TRACK] opportunities START", flush=True)
        try:
            msgs = [{"role": "user", "content": prompt}]
            content, used = _call_gemini(msgs, max_tokens=max_tokens)
            if not content:
                content, used = _call_deepseek(msgs, max_tokens=max_tokens)
            track_results['opportunities'] = (content, used)
            _print(f"  [TRACK] opportunities DONE: {'OK' if content else 'None'}", flush=True)
        except Exception as _te:
            import traceback as _tb
            _print(f"  [TRACK] opportunities EXCEPT: {_te}", flush=True)
            _print(_tb.format_exc(), flush=True)
            track_results['opportunities'] = (None, None)

    threads = [
        threading.Thread(target=run_track, kwargs={'track_id': 'eligibility', 'prompt': t1, 'max_tokens': 800,  'primary': _call_claude}, daemon=True),
        threading.Thread(target=run_track_split, args=('development', t2a, 1200, t2b, 2500), daemon=True),
        threading.Thread(target=run_track, kwargs={'track_id': 'verification', 'prompt': t3, 'max_tokens': 600, 'primary': _call_claude}, daemon=True),
        threading.Thread(target=run_track_opportunities, args=(t4, 900), daemon=True),
    ]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    e_text, e_model = track_results.get('eligibility',  (None, None))
    d_text, d_model = track_results.get('development',  (None, None))
    v_text, v_model = track_results.get('verification', (None, None))
    o_text, o_model = track_results.get('opportunities',(None, None))

    with open('webhook_debug.log', 'a') as _f:
        _f.write(
            f"[TRACKS] {candidate_name} — "
            f"E={'OK' if e_text else 'FAIL'} "
            f"D={'OK' if d_text else 'FAIL'} "
            f"V={'OK' if v_text else 'FAIL'} "
            f"O={'OK' if o_text else 'FAIL'}\n"
        )

    if not e_text:
        print(f" Council: eligibility track failed — cannot produce report")
        return None, None

    for label, text in [('development', d_text), ('verification', v_text), ('opportunities', o_text)]:
        if not text:
            print(f" Council: {label} track failed — section will be empty")

    assembled = "\n\n".join(filter(None, [e_text, d_text, v_text, o_text])) + "\n\nEND"

    credits = {
        'eligibility_model':  _model_name(e_model) if e_model else 'N/A',
        'development_model':  _model_name(d_model) if d_model else 'N/A',
        'verification_model': _model_name(v_model) if v_model else 'N/A',
        'opportunities_model':_model_name(o_model) if o_model else 'N/A',
    }
    print(
        f" Council complete — "
        f"Eligibility:{credits['eligibility_model']} | "
        f"Development:{credits['development_model']} | "
        f"Verification:{credits['verification_model']} | "
        f"Opportunities:{credits['opportunities_model']}"
    )
    return assembled, credits


def parse_vsc_analysis(raw, credits=None):
    """Parse AWIS council assembled output into a dict for the VSC report email."""
    def extract_block(label, text):
        if not text:
            return ''
        m = re.search(rf'{label}:\s*(.*?)(?=\n[A-Z_]{{2,}}:|END|\Z)', text, re.DOTALL)
        return m.group(1).strip() if m else ''

    def extract_line(label, text):
        m = re.search(rf'^{label}:\s*(.+)', text, re.MULTILINE)
        return m.group(1).strip() if m else ''

    eligibility    = extract_line('ELIGIBILITY', raw)
    requirements   = extract_block('REQUIREMENTS', raw)
    primary_grade  = extract_line('PRIMARY_GRADE', raw)
    true_fit_grade = extract_line('TRUE_FIT_GRADE', raw)
    shadow_grade   = extract_line('SHADOW_GRADE', raw)
    classification = extract_line('CLASSIFICATION', raw)
    justification  = extract_block('JUSTIFICATION', raw)
    missing_req    = extract_block('MISSING_REQUIREMENTS', raw)
    verification   = extract_block('VERIFICATION_REQUIRED', raw)
    improvements   = extract_block('IMPROVEMENTS', raw)
    ats_keywords   = extract_block('ATS_KEYWORDS', raw)
    ats_resume     = extract_block('RESUME_REWRITE', raw)

    loc_block = extract_block('LOCATION', raw)
    def loc_line(lbl, text=loc_block):
        m = re.search(rf'^{lbl}:\s*(.+)', text, re.MULTILINE)
        return m.group(1).strip() if m else ''

    candidate_loc = loc_line('Candidate')
    loc_parts     = candidate_loc.split(',')
    city          = loc_parts[0].strip() if loc_parts else 'Unknown'
    state         = loc_parts[1].strip() if len(loc_parts) > 1 else ''

    # ── Post-process: enforce eligibility rules deterministically ────────────
    # 1. If requirements contain any NOT_PROVIDED or NOT_MET, force PENDING REVIEW.
    #    The AI sometimes says SUITABLE despite gaps — override it.
    if requirements and re.search(r'\bNOT_PROVIDED\b|\bNOT_MET\b', requirements, re.IGNORECASE):
        if eligibility.upper() == 'SUITABLE':
            eligibility = 'PENDING REVIEW'
            # Cap primary grade at B if it was higher
            _grade_order = ['A+','A','A-','B+','B','B-','C','D','F']
            try:
                if _grade_order.index(primary_grade) < _grade_order.index('B'):
                    primary_grade = 'B'
            except ValueError:
                pass

    # 2. Shadow grade is N/A when SUITABLE — clear it.
    if eligibility.upper() == 'SUITABLE':
        shadow_grade = ''

    # 3. Enforce classification consistency with eligibility.
    _elig = eligibility.upper()
    _cls  = (classification or '').upper()
    if 'NOT SUITABLE' in _elig:
        if 'NON' not in _cls and 'VIABLE' not in _cls:
            classification = 'Non-viable'
    elif 'PENDING REVIEW' in _elig:
        if 'DEVELOP' not in _cls and 'NON' not in _cls:
            classification = 'Develop'
    elif _elig == 'SUITABLE':
        if 'DIRECT' not in _cls and 'REDIRECT' not in _cls:
            classification = 'Direct Fit'

    # 4. Trust Track 3 for missing/verification — do not filter by Track 1 keywords.
    #    The old filter was too aggressive and was deleting legitimate NOT_PROVIDED gaps.
    if not missing_req:
        missing_req = 'None — all listed requirements confirmed or under review.'

    parsed = {
        'eligibility':           eligibility,
        'requirements':          requirements,
        'primary_grade':         primary_grade,
        'true_fit_grade':        true_fit_grade,
        'shadow_grade':          shadow_grade if shadow_grade and shadow_grade.upper() != 'N/A' else '',
        'classification':        classification,
        'grade':                 primary_grade,
        'grade_summary':         classification,
        'candidate_city':        city,
        'candidate_state':       state,
        'justification':         justification,
        'missing_requirements':  missing_req,
        'verification_required': verification or 'None — candidate meets all stated requirements.',
        'improvements':          improvements,
        'ats_keywords':          ats_keywords,
        'ats_resume':            ats_resume,
        'recommendation':        f"{eligibility} — {classification}",
    }

    for i in range(1, 4):
        block = extract_block(f'ALTERNATIVE_{i}', raw)
        def _get(pattern, text=block):
            m = re.search(pattern, text, re.MULTILINE)
            return m.group(1).strip() if m else ''
        parsed[f'alt_{i}_title']    = _get(r'^Title:\s*(.+)')
        parsed[f'alt_{i}_company']  = _get(r'^Company:\s*(.+)')
        parsed[f'alt_{i}_distance'] = _get(r'^Distance:\s*(.+)')
        parsed[f'alt_{i}_score']    = _get(r'^Score:\s*(.+)')
        parsed[f'alt_{i}_why']      = _get(r'^Why:\s*(.+)')

    parsed.update(credits or {})
    if not credits:
        parsed.update({
            'eligibility_model':  'N/A',
            'development_model':  'N/A',
            'verification_model': 'N/A',
            'opportunities_model':'N/A',
        })
    return parsed


# ── Auth ──────────────────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return decorated


def api_login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'error': 'Authentication required'}), 401
        return f(*args, **kwargs)
    return decorated


@app.context_processor
def inject_user():
    if 'user_id' in session:
        return {'current_user': {
            'name':  session.get('vsc_name', ''),
            'email': session.get('vsc_email', ''),
            'role':  session.get('role', 'vsc'),
        }}
    return {'current_user': None}


@app.route('/login', methods=['GET', 'POST'])
def login_page():
    if 'user_id' in session:
        return redirect(url_for('home'))
    error = None
    if request.method == 'POST':
        email    = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        user     = get_user_by_email(email)
        if user and check_password_hash(user['password_hash'], password):
            session['user_id']   = user['id']
            session['vsc_name']  = user['display_name']
            session['vsc_email'] = user['email']
            session['role']      = user['role']
            return redirect(url_for('home'))
        error = 'Invalid email or password.'
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_page'))


@app.route('/')
@login_required
def home():
    """Mode selector landing page"""
    return render_template('home.html')


@app.route('/manual')
@login_required
def index():
    """Manual mode  VSC-operated Resume AI"""
    return render_template('index.html', jobs_loaded=len(JOBS_DB))


@app.route('/load_jobs', methods=['POST'])
@api_login_required
def load_jobs():
    """Accept a job CSV upload and geocode + save it in a background thread."""
    global JOBS_DB, JOB_LOAD_STATUS

    if JOB_LOAD_STATUS['state'] == 'loading':
        return jsonify({'status': 'loading', 'progress': JOB_LOAD_STATUS['progress']}), 202

    if 'file' not in request.files or request.files['file'].filename == '':
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    raw = file.read()
    df = None
    for enc in ('utf-8-sig', 'utf-8', 'cp1252', 'latin-1'):
        try:
            import io
            df = pd.read_csv(io.BytesIO(raw), encoding=enc)
            print(f"CSV read OK with encoding: {enc}, rows: {len(df)}, cols: {list(df.columns)}")
            break
        except Exception as e:
            print(f"CSV read failed with {enc}: {e}")
            continue
    if df is None:
        return jsonify({'error': 'Could not read CSV — try saving as UTF-8 from Excel'}), 400

    # Strip whitespace from column names (Excel sometimes adds spaces)
    df.columns = [c.strip() for c in df.columns]

    required_cols = ['Job Title', 'City', 'State', 'Job Description']
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        return jsonify({'error': f'Missing columns: {missing}. Found: {list(df.columns)}'}), 400

    # Phase 1: save all jobs immediately without geocoding (sub-second)
    df['latitude']  = None
    df['longitude'] = None
    JOBS_DB = df.to_dict('records')
    saved = save_jobs(JOBS_DB)
    log_job_upload(saved)
    print(f"Jobs saved to DB immediately: {saved}")

    JOB_LOAD_STATUS.update({
        'state':    'geocoding',
        'progress': f'{saved} jobs saved. Geocoding locations in background...',
        'total':    saved,
        'geocoded': 0,
    })

    # Phase 2: geocode unique locations in background, update DB records
    df_copy = df.copy()

    def _background_geocode(df):
        global JOBS_DB, JOB_LOAD_STATUS
        try:
            from data.database import get_connection
            unique_locs = df[['City', 'State']].dropna().drop_duplicates()
            loc_cache   = {}
            done        = 0

            for _, loc_row in unique_locs.iterrows():
                city  = str(loc_row['City']).strip()
                state = str(loc_row['State']).strip()
                key   = (city, state)
                lat, lon = geocode_location(city, state)
                loc_cache[key] = (lat, lon)
                done += 1
                JOB_LOAD_STATUS['geocoded'] = done
                JOB_LOAD_STATUS['progress'] = (
                    f"Geocoding {done}/{len(unique_locs)}: {city}, {state}"
                )

            # Batch-update lat/lon in DB
            conn = get_connection()
            c = conn.cursor()
            for (city, state), (lat, lon) in loc_cache.items():
                if lat and lon:
                    c.execute("""
                        UPDATE job_listings
                        SET latitude=?, longitude=?
                        WHERE city=? AND state=? AND is_active=1
                    """, (lat, lon, city, state))
            conn.commit()
            conn.close()

            # Refresh in-memory copy with coordinates
            for job in JOBS_DB:
                key = (str(job.get('City', '')).strip(), str(job.get('State', '')).strip())
                lat, lon = loc_cache.get(key, (None, None))
                job['latitude']  = lat
                job['longitude'] = lon

            JOB_LOAD_STATUS.update({
                'state':    'done',
                'progress': f'{saved} jobs loaded, {done} locations geocoded',
                'total':    saved,
            })
            print(f"Geocoding complete: {done} locations")

        except Exception as e:
            JOB_LOAD_STATUS.update({'state': 'done',
                                    'progress': f'{saved} jobs loaded (geocoding failed: {e})'})
            print(f"Background geocoding failed: {e}")

    t = threading.Thread(target=_background_geocode, args=(df_copy,), daemon=True)
    t.start()

    return jsonify({
        'status':       'done',
        'jobs_loaded':  saved,
        'message':      f'{saved} jobs loaded. Geocoding {len(df_copy)} locations in background for geo-matching.',
    })


@app.route('/api/jobs/status', methods=['GET'])
@api_login_required
def jobs_status():
    """Return current job-load progress."""
    return jsonify({
        'state':    JOB_LOAD_STATUS['state'],
        'progress': JOB_LOAD_STATUS['progress'],
        'total':    JOB_LOAD_STATUS['total'],
        'geocoded': JOB_LOAD_STATUS['geocoded'],
        'in_memory': len(JOBS_DB),
    })


@app.route('/analyze', methods=['POST'])
@api_login_required
def analyze_resume():
    """Analyze resume against target job + recommend alternatives"""
    try:
        data = request.json
        
        resume_text = data.get('resume_text', '')
        candidate_name = data.get('candidate_name', 'Unknown')
        candidate_city = data.get('candidate_city', '')
        candidate_state = data.get('candidate_state', '')
        target_job_id = data.get('target_job_id', '')
        
        if not resume_text:
            return jsonify({'error': 'Resume text required'}), 400
        
        if not JOBS_DB:
            return jsonify({'error': 'No jobs loaded. Upload CSV first.'}), 400
        
        # Find target job
        target_job = next((j for j in JOBS_DB if str(j.get('Job Id', '')) == target_job_id), None)
        if not target_job:
            return jsonify({'error': 'Target job not found'}), 400
        
        # Geocode candidate location
        candidate_location = f"{candidate_city}, {candidate_state}"
        cand_lat, cand_lon = geocode_location(candidate_city, candidate_state)
        
        if cand_lat is None:
            return jsonify({'error': 'Could not geocode candidate location'}), 400
        
        # Calculate distances for all jobs
        for job in JOBS_DB:
            if job.get('latitude') and job.get('longitude'):
                dist = calculate_distance(cand_lat, cand_lon, job['latitude'], job['longitude'])
                job['distance'] = round(dist, 1) if dist else None
            else:
                job['distance'] = None
        
        # Filter jobs within 50 miles
        nearby_jobs = [j for j in JOBS_DB if j.get('distance') and j['distance'] <= 50]
        
        # Sort by distance
        nearby_jobs.sort(key=lambda x: x.get('distance', 999))
        
        # Run AI analysis
        analysis = analyze_with_ai(
            resume_text,
            candidate_name,
            candidate_location,
            target_job,
            nearby_jobs
        )
        
        # Send management email
        send_management_email(
            candidate_name,
            target_job.get('Job Title', 'Unknown Position')
        )
        
        return jsonify({
            'status': 'success',
            'candidate_name': candidate_name,
            'target_job': target_job.get('Job Title'),
            'nearby_jobs_count': len(nearby_jobs),
            'analysis': analysis
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/search_jobs', methods=['POST'])
@api_login_required
def search_jobs():
    """Search jobs by location"""
    try:
        data = request.json
        city = data.get('city', '')
        state = data.get('state', '')
        
        if not city or not state:
            return jsonify({'error': 'City and state required'}), 400
        
        # Geocode search location
        lat, lon = geocode_location(city, state)
        if lat is None:
            return jsonify({'error': 'Could not geocode location'}), 400
        
        # Calculate distances
        for job in JOBS_DB:
            if job.get('latitude') and job.get('longitude'):
                dist = calculate_distance(lat, lon, job['latitude'], job['longitude'])
                job['distance'] = round(dist, 1) if dist else None
        
        # Filter within 50 miles
        results = [j for j in JOBS_DB if j.get('distance') and j['distance'] <= 50]
        results.sort(key=lambda x: x['distance'])
        
        # Return simplified data
        return jsonify({
            'status': 'success',
            'count': len(results),
            'jobs': [{
                'id': j.get('Job Id'),
                'title': j.get('Job Title'),
                'company': j.get('Company Name'),
                'city': j.get('City'),
                'state': j.get('State'),
                'distance': j.get('distance')
            } for j in results[:50]]  # Limit to 50 results
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/jobs/search', methods=['GET'])
@api_login_required
def jobs_search_by_title():
    """Search jobs by title text. Returns up to 50 matches."""
    q = request.args.get('q', '').strip().lower()
    if not q:
        return jsonify({'jobs': []})
    matches = [
        {
            'id': j.get('Job Id') or str(j.get('id', '')),
            'title': j.get('Job Title') or j.get('job_title', ''),
            'company': j.get('Company Name') or j.get('company_name', ''),
            'city': j.get('City') or j.get('city', ''),
            'state': j.get('State') or j.get('state', ''),
            'distance': j.get('distance'),
        }
        for j in JOBS_DB
        if q in (j.get('Job Title') or j.get('job_title', '') or '').lower()
    ]
    return jsonify({'jobs': matches[:50], 'total': len(matches)})


@app.route('/api/jobs/db-age', methods=['GET'])
@api_login_required
def jobs_db_age():
    """Return last upload timestamp and age in days."""
    return jsonify(get_last_job_upload())


@app.route('/api/jobs/prune', methods=['DELETE'])
@api_login_required
def jobs_prune():
    """Delete all jobs matching a title from the DB and in-memory cache."""
    global JOBS_DB
    title = (request.json or {}).get('title', '').strip()
    if not title:
        return jsonify({'error': 'title is required'}), 400
    deleted = delete_job_by_title(title)
    JOBS_DB = [j for j in JOBS_DB
               if (j.get('Job Title') or j.get('job_title', '')).lower() != title.lower()]
    return jsonify({'deleted': deleted, 'title': title})


@app.route('/generate_package', methods=['POST'])
@api_login_required
def generate_calcareers_package():
    """Generate CalCareers application package"""
    try:
        data = request.json
        
        # Extract candidate data
        candidate = CandidateProfile(
            legal_name=data.get('legal_name', ''),
            date_of_birth=data.get('date_of_birth', ''),
            address=data.get('address', ''),
            phone=data.get('phone', ''),
            email=data.get('email', ''),
            education_entries=[
                EducationEntry(
                    institution=edu.get('institution', ''),
                    degree=edu.get('degree', ''),
                    major=edu.get('major', ''),
                    graduation_date=edu.get('graduation_date', '')
                )
                for edu in data.get('education', [])
            ],
            work_experience_entries=[
                WorkExperienceEntry(
                    employer=exp.get('employer', ''),
                    job_title=exp.get('job_title', ''),
                    start_date=exp.get('start_date', ''),
                    end_date=exp.get('end_date'),
                    duties=exp.get('duties', '')
                )
                for exp in data.get('work_experience', [])
            ],
            ecos_id=data.get('ecos_id')
        )
        
        # Extract job data
        job = JobTarget(
            jc_number=data.get('jc_number', ''),
            classification_title=data.get('classification_title', ''),
            department=data.get('department'),
            final_filing_date=data.get('final_filing_date')
        )
        
        # Extract Veterans' Preference
        vp_data = data.get('veterans_preference', {})
        claiming_vp = vp_data.get('claiming', False)
        
        vp_basis = None
        if claiming_vp and vp_data.get('basis'):
            vp_basis = ApplicationBasis[vp_data['basis']]
        
        vp = VeteransPreference(
            claiming_veterans_preference=claiming_vp,
            application_basis=vp_basis
        )
        
        # Determine template track
        template_track_str = data.get('template_track', 'ANALYST')
        template_track = TemplateTrack[template_track_str]
        
        # Create package input
        pkg_input = PackageInput(
            candidate=candidate,
            job=job,
            template_track=template_track,
            veterans_preference=vp
        )
        
        # Generate package
        temp_dir = tempfile.mkdtemp()
        packager = PackageGenerator(output_root=temp_dir)
        decision_engine = DecisionEngine()
        checklist_gen = ChecklistGenerator()
        audit_logger = AuditLogger()
        
        # Run decision engine
        decisions = decision_engine.analyze_package(pkg_input)
        authorities = decision_engine.get_authorities()
        
        # Validate
        is_valid, missing_data = decision_engine.validate_inputs(pkg_input)
        
        # Create package structure
        package_path = packager.create_package_structure(pkg_input)
        packager.create_placeholder_files(package_path, pkg_input, decisions)
        
        # Generate checklist
        checklist_content = checklist_gen.generate_checklist(pkg_input, decisions)
        checklist_path = package_path / "01_Checklists" / f"{pkg_input.get_filename_prefix()}_Checklist.txt"
        checklist_path.write_text(checklist_content, encoding='utf-8')
        
        # Create audit log
        audit_record = audit_logger.create_audit_record(pkg_input, decisions, missing_data, authorities)
        audit_logger.save_audit_log(audit_record, package_path)
        
        # Generate missing data report if needed
        if missing_data:
            audit_logger.generate_missing_data_report(missing_data, package_path)
        
        # Create ZIP file
        zip_filename = f"{pkg_input.get_package_name()}.zip"
        zip_path = Path(temp_dir) / zip_filename
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(package_path):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(package_path.parent)
                    zipf.write(file_path, arcname)
        
        return send_file(
            zip_path,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/health', methods=['GET'])
def health():
    """Health check"""
    return jsonify({
        'status': 'healthy',
        'jobs_loaded': len(JOBS_DB),
        'api_key_set': bool(DEEPSEEK_API_KEY)
    })


#  Email Manager 

@app.route('/email-manager')
@login_required
def email_manager():
    """Email manager page"""
    return render_template('email_manager.html')


@app.route('/api/templates', methods=['GET'])
@api_login_required
def get_templates():
    """Return built-in templates + VSC-saved custom templates."""
    built_in = [
        {'key': t['key'], 'label': t['label'], 'stage': t['stage'],
         'when': t['when'], 'fields': t['fields'], 'custom': False}
        for t in TEMPLATES
        if t['key'] != 'vsc_analysis_report'   # internal-only template
    ]
    custom = [
        {'key': r['key'], 'label': r['label'], 'stage': r['stage'],
         'when': r.get('when_to_send', ''), 'fields': [], 'custom': True,
         'subject': r['subject'], 'body': r['body']}
        for r in get_custom_templates()
    ]
    return jsonify({'templates': built_in + custom})


@app.route('/api/templates', methods=['POST'])
@api_login_required
def create_template():
    """Save a VSC-authored custom template."""
    data = request.json or {}
    label = (data.get('label') or '').strip()
    subject = (data.get('subject') or '').strip()
    body = (data.get('body') or '').strip()
    stage = (data.get('stage') or 'GENERAL').strip().upper()
    when_to_send = (data.get('when_to_send') or '').strip()

    if not label:
        return jsonify({'success': False, 'error': 'Template name is required'}), 400
    if not subject:
        return jsonify({'success': False, 'error': 'Subject is required'}), 400
    if not body:
        return jsonify({'success': False, 'error': 'Body is required'}), 400

    tmpl = save_custom_template(label=label, subject=subject, body=body,
                                stage=stage, when_to_send=when_to_send)
    return jsonify({'success': True, 'template': tmpl})


@app.route('/api/templates/<key>', methods=['DELETE'])
@api_login_required
def delete_template(key):
    """Delete a custom template."""
    if not key.startswith('custom_'):
        return jsonify({'success': False, 'error': 'Only custom templates can be deleted'}), 403
    deleted = delete_custom_template(key)
    return jsonify({'success': deleted})


@app.route('/api/preview_email', methods=['POST'])
@api_login_required
def preview_email_route():
    """Render a template with variables and return subject + body (no send)"""
    data         = request.json
    template_key = data.get('template_key', '')
    variables    = data.get('variables', {})

    result = svc_preview_email(template_key, variables)
    return jsonify(result)


@app.route('/api/send_email', methods=['POST'])
@api_login_required
def send_email_route():
    """Send an email and log it"""
    data = request.json

    vsc_name     = session['vsc_name']
    vsc_email    = session['vsc_email']
    to_name      = data.get('to_name', '').strip()
    to_email     = data.get('to_email', '').strip()
    template_key = data.get('template_key', '').strip()
    subject      = data.get('subject', '').strip()
    body         = data.get('body', '').strip()

    if not all([to_name, to_email, subject, body]):
        return jsonify({'success': False, 'error': 'Missing required fields'}), 400

    from data.email_templates import get_template
    from data.database import log_email as db_log_email
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    if not template_key:
        template_key = 'custom'

    tmpl = get_template(template_key)
    if tmpl:
        template_label = tmpl['label']
    else:
        custom_tmpl = next((t for t in get_custom_templates() if t['key'] == template_key), None)
        template_label = custom_tmpl['label'] if custom_tmpl else 'Custom Email'

    status        = 'sent'
    error_message = None

    try:
        if not OUTLOOK_PASSWORD:
            raise ValueError('SMTP credentials not configured')

        msg = MIMEMultipart()
        msg['From']     = f'Work for Warriors <{OUTLOOK_USER}>'
        msg['To']       = f'{to_name} <{to_email}>'
        msg['Reply-To'] = f'{vsc_name} <{vsc_email}>' if vsc_email else OUTLOOK_USER
        msg['Subject']  = subject
        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.ehlo()
            server.starttls()
            server.login(OUTLOOK_USER, OUTLOOK_PASSWORD)
            server.send_message(msg)

    except Exception as e:
        status        = 'failed'
        error_message = str(e)

    log_id = db_log_email(
        vsc_name      = vsc_name,
        vsc_email     = vsc_email,
        to_name       = to_name,
        to_email      = to_email,
        template_key  = template_key,
        template_label= template_label,
        subject       = subject,
        body          = body,
        status        = status,
        error_message = error_message,
    )

    if status == 'sent':
        return jsonify({'success': True, 'log_id': log_id})
    else:
        return jsonify({'success': False, 'log_id': log_id, 'error': error_message})


@app.route('/api/email_history', methods=['GET'])
@api_login_required
def email_history_route():
    """Return recent email log. Admins see all; VSCs see only non-archived."""
    candidate_email  = request.args.get('candidate_email')
    is_admin         = session.get('role') == 'admin'
    history = get_email_history(limit=50, candidate_email=candidate_email,
                                include_archived=is_admin)
    return jsonify({'history': history, 'is_admin': is_admin})


@app.route('/api/email_logs/<int:log_id>', methods=['DELETE'])
@api_login_required
def delete_email_log_route(log_id):
    """VSC: soft-archive (hidden from their view, admin retains). Admin: hard delete."""
    is_admin = session.get('role') == 'admin'
    if is_admin:
        ok = hard_delete_email_log(log_id)
    else:
        ok = archive_email_log(log_id)
    if ok:
        return jsonify({'success': True})
    return jsonify({'success': False, 'error': 'Log entry not found'}), 404


@app.route('/api/email_logs/<int:log_id>/restore', methods=['POST'])
@api_login_required
def restore_email_log_route(log_id):
    """Admin only — un-archive a VSC-deleted entry."""
    if session.get('role') != 'admin':
        return jsonify({'error': 'Admin access required'}), 403
    ok = restore_email_log(log_id)
    if ok:
        return jsonify({'success': True})
    return jsonify({'success': False, 'error': 'Log entry not found'}), 404


#  CRM 

@app.route('/crm')
@login_required
def crm():
    return render_template('crm.html',
                           pipeline_stages=PIPELINE_STAGES,
                           stage_labels=STAGE_LABELS,
                           stage_colors=STAGE_COLORS,
                           military_branches=MILITARY_BRANCHES,
                           engagement_types=ENGAGEMENT_TYPES,
                           engagement_subtypes=ENGAGEMENT_SUBTYPES,
                           engagement_type_labels=ENGAGEMENT_TYPE_LABELS,
                           engagement_subtype_labels=ENGAGEMENT_SUBTYPE_LABELS)


@app.route('/api/candidates', methods=['GET'])
@api_login_required
def candidates_list():
    if session.get('role') == 'admin':
        vsc_name = request.args.get('vsc_name')  # admin can filter by any VSC or see all
    else:
        vsc_name = session['vsc_name']
    stage  = request.args.get('stage')
    search = request.args.get('search')
    return jsonify({'candidates': get_candidates(vsc_name=vsc_name, stage=stage, search=search)})


@app.route('/api/candidates', methods=['POST'])
@api_login_required
def candidates_create():
    data = request.json
    if not data.get('first_name') or not data.get('last_name'):
        return jsonify({'error': 'first_name and last_name are required'}), 400
    data['vsc_name'] = session['vsc_name']
    cid = create_candidate(data)
    return jsonify({'success': True, 'id': cid})


@app.route('/api/candidates/<int:cid>', methods=['GET'])
@api_login_required
def candidates_get(cid):
    candidate = get_candidate(cid)
    if not candidate:
        return jsonify({'error': 'Not found'}), 404
    subs   = get_submissions(cid)
    emails = get_email_history(limit=30, candidate_email=candidate.get('email'))
    return jsonify({'candidate': candidate, 'submissions': subs, 'emails': emails})


@app.route('/api/candidates/<int:cid>', methods=['PUT'])
@api_login_required
def candidates_update(cid):
    data = request.json
    ok = update_candidate(cid, data)
    return jsonify({'success': ok})


@app.route('/api/candidates/<int:cid>', methods=['DELETE'])
@api_login_required
def candidates_delete(cid):
    ok = delete_candidate(cid)
    return jsonify({'success': ok})


@app.route('/api/pipeline_counts', methods=['GET'])
@api_login_required
def pipeline_counts():
    if session.get('role') == 'admin':
        vsc_name = request.args.get('vsc_name')
    else:
        vsc_name = session['vsc_name']
    counts = get_pipeline_counts(vsc_name=vsc_name)
    return jsonify(counts)


@app.route('/api/candidates/<int:cid>/submissions', methods=['POST'])
@api_login_required
def submissions_create(cid):
    data         = request.json
    job_title    = data.get('job_title', '').strip()
    company_name = data.get('company_name', '').strip()
    notes        = data.get('notes', '').strip()
    if not job_title or not company_name:
        return jsonify({'error': 'job_title and company_name required'}), 400
    sub_id = add_submission(cid, job_title, company_name, notes)
    return jsonify({'success': True, 'id': sub_id})


@app.route('/api/submissions/<int:sub_id>', methods=['PUT'])
@api_login_required
def submissions_update(sub_id):
    data    = request.json
    outcome = data.get('outcome', '').strip()
    notes   = data.get('notes')
    if not outcome:
        return jsonify({'error': 'outcome required'}), 400
    ok = update_submission(sub_id, outcome, notes)
    return jsonify({'success': ok})


#  Automated Mode 

@app.route('/dashboard')
@login_required
def management_dashboard():
    if session.get('role') != 'admin':
        return redirect(url_for('crm'))
    stats = get_dashboard_stats()
    return render_template('dashboard.html',
                           stats=stats,
                           pipeline_stages=PIPELINE_STAGES,
                           stage_labels=STAGE_LABELS,
                           stage_colors=STAGE_COLORS)


@app.route('/api/dashboard/stats', methods=['GET'])
@api_login_required
def dashboard_stats_api():
    if session.get('role') != 'admin':
        return jsonify({'error': 'Admin access required'}), 403
    return jsonify(get_dashboard_stats())


@app.route('/automated')
@login_required
def automated_dashboard():
    """Automated mode VSC dashboard"""
    return render_template('automated_dashboard.html')


@app.route('/intake')
@login_required
def intake():
    """Candidate intake form"""
    return render_template('intake.html')


@app.route('/intake/confirm')
@login_required
def intake_confirm():
    """Thank you page after intake submission"""
    return render_template('intake_confirm.html')


@app.route('/api/intake/jobs', methods=['POST'])
@api_login_required
def intake_jobs():
    """
    Find jobs near candidate location for intake form.
    Input: {city, state}
    Returns simplified job list with board field included.
    Geocodes location, queries job_listings DB.
    Falls back to JOBS_DB in-memory if DB empty.
    """
    try:
        data  = request.json or {}
        city  = data.get('city', '').strip()
        state = data.get('state', '').strip()

        if not city or not state:
            return jsonify({'error': 'city and state are required'}), 400

        lat, lon = geocode_location(city, state)
        if lat is None:
            return jsonify({'error': f'Could not geocode location: {city}, {state}'}), 400

        # Try DB first
        db_results = get_jobs_near(lat, lon, radius_miles=50)

        if db_results:
            jobs = [{
                'id':       j.get('job_id') or str(j.get('id', '')),
                'job_id':   j.get('job_id', ''),
                'title':    j.get('job_title', ''),
                'company':  j.get('company_name', ''),
                'city':     j.get('city', ''),
                'state':    j.get('state', ''),
                'distance': round(j.get('distance_miles', 0), 1),
                'board':    j.get('board', ''),
            } for j in db_results]
        else:
            # Fall back to in-memory JOBS_DB
            for job in JOBS_DB:
                if job.get('latitude') and job.get('longitude'):
                    dist = calculate_distance(lat, lon, job['latitude'], job['longitude'])
                    job['distance'] = round(dist, 1) if dist else None

            nearby = [j for j in JOBS_DB if j.get('distance') and j['distance'] <= 50]
            nearby.sort(key=lambda x: x['distance'])

            jobs = [{
                'id':       str(j.get('Job Id', '')),
                'job_id':   str(j.get('Job Id', '')),
                'title':    j.get('Job Title', ''),
                'company':  j.get('Company Name', ''),
                'city':     j.get('City', ''),
                'state':    j.get('State', ''),
                'distance': j.get('distance'),
                'board':    j.get('Board', ''),
            } for j in nearby[:50]]

        return jsonify({'status': 'success', 'count': len(jobs), 'jobs': jobs})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/intake/submit', methods=['POST'])
@api_login_required
def intake_submit():
    """
    Process candidate intake form submission.
    Accepts multipart/form-data (has resume file upload).
    Extracts resume text, runs automation engine.
    Returns JSON {success, message, candidate_id}.
    """
    try:
        #  Parse form fields 
        def fv(key, default=''):
            """Get form value, strip whitespace."""
            return (request.form.get(key) or default).strip()

        vsc_name     = session['vsc_name']
        first_name   = fv('first_name')
        last_name    = fv('last_name')
        email        = fv('email')
        phone        = fv('phone')
        city         = fv('city')
        state        = fv('state')
        branch       = fv('branch')
        rank         = fv('rank')
        mos          = fv('mos')
        years_served = fv('years_served')
        target_job_id = fv('target_job_id')

        #  Resume text 
        resume_text = ''
        resume_error = ''

        file = request.files.get('resume')
        if file and file.filename:
            resume_text, resume_error = parse_resume(file)
        else:
            resume_text = fv('resume_text')

        if not resume_text:
            error_msg = resume_error if resume_error else 'No resume text provided'
            return jsonify({'success': False, 'error': error_msg}), 400

        #  CalCareers fields (optional JSON strings) 
        education = []
        work_experience = []
        veterans_preference = {}
        template_track = 'ANALYST'

        try:
            edu_raw = request.form.get('education')
            if edu_raw:
                education = json.loads(edu_raw)
        except Exception:
            pass

        try:
            exp_raw = request.form.get('work_experience')
            if exp_raw:
                work_experience = json.loads(exp_raw)
        except Exception:
            pass

        try:
            vp_raw = request.form.get('veterans_preference')
            if vp_raw:
                veterans_preference = json.loads(vp_raw)
        except Exception:
            pass

        template_track = fv('template_track', 'ANALYST')

        #  Build intake_data dict 
        intake_data = {
            'vsc_name':           vsc_name,
            'vsc_email':          '',       # not collected on intake form; VSC notified via outlook_user
            'first_name':         first_name,
            'last_name':          last_name,
            'email':              email,
            'phone':              phone,
            'city':               city,
            'state':              state,
            'branch':             branch,
            'rank':               rank,
            'mos':                mos,
            'years_served':       years_served,
            'resume_text':        resume_text,
            'target_job_id':      target_job_id,
            'education':          education,
            'work_experience':    work_experience,
            'veterans_preference': veterans_preference,
            'template_track':     template_track,
            'outlook_user':       OUTLOOK_USER,
            'outlook_password':   OUTLOOK_PASSWORD,
            'ryan_email':         RYAN_EMAIL,
        }

        #  Look up target job
        matched = get_job_by_id(target_job_id)
        if not matched:
            return jsonify({'success': False, 'error': 'Selected job not found in database'}), 400

        target_job = {
            'Job Title':       matched.get('job_title', ''),
            'Company Name':    matched.get('company_name', ''),
            'City':            matched.get('city', ''),
            'State':           matched.get('state', ''),
            'Job Description': matched.get('job_description', ''),
            'Qualifications':  matched.get('qualifications', ''),
            'Salary From':     matched.get('salary_from', ''),
            'Salary To':       matched.get('salary_to', ''),
        }
        candidate_name = f"{first_name} {last_name}".strip()

        #  Geocode and find nearby jobs
        lat, lon = geocode_location(city, state) if city and state else (None, None)
        alt_jobs = []
        if lat:
            db_nearby = get_jobs_near(lat, lon, radius_miles=50)
            alt_jobs = [{
                'Job Title':    j.get('job_title', ''),
                'Company Name': j.get('company_name', ''),
                'City':         j.get('city', ''),
                'State':        j.get('state', ''),
                'distance':     round(j.get('distance_miles', 0), 1),
            } for j in db_nearby]
        if not alt_jobs:
            alt_jobs = list(JOBS_DB[:50])

        #  Run AI council analysis
        raw_analysis, track_credits = analyze_for_vsc(resume_text, candidate_name, target_job, alt_jobs)
        if not raw_analysis:
            return jsonify({'success': False, 'error': 'AI analysis failed'}), 500

        parsed = parse_vsc_analysis(raw_analysis, credits=track_credits)
        grade  = parsed.get('grade', 'N/A')

        #  Create CRM record
        candidate_id = create_candidate({
            'vsc_name':    vsc_name,
            'first_name':  first_name,
            'last_name':   last_name,
            'email':       email,
            'phone':       phone,
            'city':        city,
            'state':       state,
            'stage':       'RESUME_ANALYZED',
            'notes':       (
                f"Applied for: {target_job['Job Title']}\n"
                f"Eligibility: {parsed.get('eligibility', '')}\n"
                f"Grade: {grade} | True Fit: {parsed.get('true_fit_grade', '')} | Shadow: {parsed.get('shadow_grade', 'N/A')}\n"
                f"Classification: {parsed.get('classification', '')}\n\n"
                f"VERIFICATION REQUIRED:\n{parsed.get('verification_required', '')}"
            ),
            'resume_text': resume_text,
            'source':      'intake',
        })
        log_engagement(
            candidate_id=candidate_id,
            vsc_name=vsc_name,
            eng_type='INTAKE',
            notes=f"Manual intake for {target_job['Job Title']}. AI Grade: {grade}."
        )

        #  Send VSC report email
        if GRAPH_CLIENT_SECRET:
            try:
                from data.email_templates import render_template as render_email
                report = render_email('vsc_analysis_report', {
                    'candidate_name':        candidate_name,
                    'candidate_email':       email or '(not provided)',
                    'vsc_name':              vsc_name,
                    'job_title':             target_job['Job Title'],
                    'company_name':          target_job['Company Name'],
                    'eligibility':           parsed.get('eligibility', ''),
                    'requirements':          parsed.get('requirements', ''),
                    'primary_grade':         parsed.get('primary_grade', grade),
                    'true_fit_grade':        parsed.get('true_fit_grade', ''),
                    'shadow_grade':          parsed.get('shadow_grade', '') or 'N/A',
                    'classification':        parsed.get('classification', ''),
                    'justification':         parsed.get('justification', ''),
                    'missing_requirements':  parsed.get('missing_requirements', ''),
                    'verification_required': parsed.get('verification_required', ''),
                    'improvements':          parsed.get('improvements', ''),
                    'ats_keywords':          parsed.get('ats_keywords', ''),
                    'ats_resume':            parsed.get('ats_resume', ''),
                    'candidate_city':        city or parsed.get('candidate_city', ''),
                    'alt_1_title':           parsed.get('alt_1_title', ''),
                    'alt_1_company':         parsed.get('alt_1_company', ''),
                    'alt_1_distance':        parsed.get('alt_1_distance', ''),
                    'alt_1_score':           parsed.get('alt_1_score', ''),
                    'alt_1_why':             parsed.get('alt_1_why', ''),
                    'alt_2_title':           parsed.get('alt_2_title', ''),
                    'alt_2_company':         parsed.get('alt_2_company', ''),
                    'alt_2_distance':        parsed.get('alt_2_distance', ''),
                    'alt_2_score':           parsed.get('alt_2_score', ''),
                    'alt_2_why':             parsed.get('alt_2_why', ''),
                    'alt_3_title':           parsed.get('alt_3_title', ''),
                    'alt_3_company':         parsed.get('alt_3_company', ''),
                    'alt_3_distance':        parsed.get('alt_3_distance', ''),
                    'alt_3_score':           parsed.get('alt_3_score', ''),
                    'alt_3_why':             parsed.get('alt_3_why', ''),
                    'eligibility_model':     parsed.get('eligibility_model',  'N/A'),
                    'development_model':     parsed.get('development_model',  'N/A'),
                    'verification_model':    parsed.get('verification_model', 'N/A'),
                    'opportunities_model':   parsed.get('opportunities_model','N/A'),
                })
                _graph_send_email(
                    to_addr  = OUTLOOK_USER,
                    to_name  = vsc_name,
                    subject  = report['subject'],
                    body     = report['body'],
                    reply_to = OUTLOOK_USER,
                )
            except Exception as e:
                print(f"  Intake VSC email failed: {e}")

        return jsonify({
            'success':      True,
            'candidate_id': candidate_id,
            'grade':        grade,
            'errors':       [],
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/intake/submissions', methods=['GET'])
@api_login_required
def intake_submissions():
    """Return recent automated intake submissions for VSC dashboard"""
    try:
        limit = int(request.args.get('limit', 50))
        rows = get_intake_submissions(limit=limit)
        return jsonify({'submissions': rows})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/intake/process', methods=['POST'])
def intake_process():
    """
    Primary automated intake endpoint.

    Called by the inbox poller when a SJB application email arrives.
    Accepts:
        resume_text       full resume as plain text (required)
        subject_line      raw email subject, e.g. "Application for Data Analyst"
        job_title         parsed job title (optional if subject_line provided)
        vsc_name          name of the VSC who owns this inbox (required)
        vsc_email         VSC email address (required  analysis report sent here)
        candidate_name    sender display name
        candidate_email   sender email address
        candidate_phone   optional — all numbers joined by " / "
        candidate_address optional — street address
        candidate_city    optional — pre-extracted city
        candidate_state   optional — pre-extracted state (2-letter)
    """
    import sys
    with open('webhook_debug.log', 'a') as _dbg:
        _dbg.write(f"[WEBHOOK CALLED]\n")
    try:
        data = request.json or {}

        resume_text     = (data.get('resume_text') or '').strip()
        subject_line    = (data.get('subject_line') or '').strip()
        job_title       = (data.get('job_title') or '').strip()
        vsc_name        = (data.get('vsc_name') or 'VSC').strip()
        vsc_email       = (data.get('vsc_email') or OUTLOOK_USER).strip()
        candidate_name    = (data.get('candidate_name') or 'Unknown Candidate').strip()
        candidate_name    = re.sub(r'\s+via\s+.+$', '', candidate_name, flags=re.IGNORECASE).strip()
        candidate_email   = (data.get('candidate_email') or '').strip()
        candidate_phone   = (data.get('candidate_phone') or '').strip()
        candidate_address = (data.get('candidate_address') or '').strip()
        candidate_city_in = (data.get('candidate_city') or '').strip()
        candidate_state_in= (data.get('candidate_state') or '').strip()
        email_body        = (data.get('email_body') or '').strip()

        # Append email body to resume text if provided — often contains
        # cover letter content or qualifications not in the attached resume.
        if resume_text and email_body:
            resume_text = f"{resume_text}\n\n--- CANDIDATE APPLICATION MESSAGE ---\n{email_body}"
        elif email_body and not resume_text:
            resume_text = email_body

        # Parse job title from subject line if not provided directly
        if not job_title and subject_line:
            m = re.search(r'application\s+for\s+(.+)', subject_line, re.IGNORECASE)
            if m:
                job_title = m.group(1).strip()

        if not resume_text:
            return jsonify({'success': False, 'error': 'resume_text is required'}), 400
        if not job_title:
            return jsonify({'success': False, 'error': 'Could not determine job title. Provide job_title or subject_line containing "Application for [title]"'}), 400

        print(f" New application: {candidate_name}  {job_title} (VSC: {vsc_name})")

        #  Step 1: Look up job description from DB 
        matched_job = get_job_by_title(job_title)
        if matched_job:
            target_job = {
                'Job Title':       matched_job.get('job_title', job_title),
                'Company Name':    matched_job.get('company_name', ''),
                'City':            matched_job.get('city', ''),
                'State':           matched_job.get('state', ''),
                'Job Description': matched_job.get('job_description', ''),
                'Qualifications':  matched_job.get('qualifications', ''),
                'Salary From':     matched_job.get('salary_from', ''),
                'Salary To':       matched_job.get('salary_to', ''),
            }
            print(f" Job matched in DB: {target_job['Job Title']} at {target_job['Company Name']}")
        else:
            # Job not in DB — hold the application, notify VSC, do not run AI
            db_info    = get_last_job_upload()
            age_str    = f"{db_info['age_days']} days old" if db_info['age_days'] is not None else "unknown age"
            hold_notes = (
                f"APPLICATION HELD — Job not in database\n"
                f"Applied for: {job_title}\n"
                f"Jobs DB is {age_str} ({db_info['count']} jobs).\n\n"
                f"ACTION REQUIRED:\n"
                f"1. Export a fresh CSV from SmartJobBoard and upload it at /manual\n"
                f"2. Use the Prune tool to remove any expired listings\n"
                f"3. Notify the candidate that their application is under review"
            )
            hold_id = None
            try:
                name_parts = candidate_name.split(' ', 1)
                hold_id = create_candidate({
                    'vsc_name':    vsc_name,
                    'first_name':  name_parts[0],
                    'last_name':   name_parts[1] if len(name_parts) > 1 else '',
                    'email':       candidate_email,
                    'phone':       candidate_phone,
                    'city':        '',
                    'state':       '',
                    'stage':       'HELD_REVIEW',
                    'notes':       hold_notes,
                    'resume_text': resume_text,
                    'source':      'intake',
                })
                log_engagement(
                    candidate_id=hold_id,
                    vsc_name=vsc_name,
                    eng_type='INTAKE',
                    notes=f"Application held — '{job_title}' not found in jobs DB."
                )
            except Exception as e:
                print(f"  CRM hold save failed: {e}")

            _hold_email = "email skipped (Graph not configured)"
            if GRAPH_CLIENT_SECRET:
                try:
                    hold_subject = f"APPLICATION HELD — {candidate_name} | {job_title} not in DB"
                    hold_body = (
                        f"APPLICATION HELD — ACTION REQUIRED\n"
                        f"{'='*50}\n\n"
                        f"Candidate:   {candidate_name}\n"
                        f"Email:       {candidate_email}\n"
                        f"Applied for: {job_title}\n\n"
                        f"REASON: '{job_title}' was not found in the jobs database.\n"
                        f"Jobs DB is currently {age_str} ({db_info['count']} jobs).\n\n"
                        f"REQUIRED ACTIONS:\n"
                        f"1. Export a fresh CSV from SmartJobBoard\n"
                        f"2. Upload it at http://localhost:5001/manual (Step 1)\n"
                        f"3. Use the Prune tool to remove any expired listings\n"
                        f"4. Notify {candidate_name} ({candidate_email}) that their\n"
                        f"   application is under review — do not leave them waiting\n\n"
                        f"The candidate record has been added to your CRM under:\n"
                        f"Stage: Held — Check SJB DB\n\n"
                        f"Once the DB is updated, reprocess the application manually."
                    )
                    _graph_send_email(
                        to_addr  = vsc_email,
                        to_name  = vsc_name,
                        subject  = hold_subject,
                        body     = hold_body,
                        reply_to = OUTLOOK_USER,
                    )
                    _hold_email = "VSC notified"
                except Exception as e:
                    _hold_email = f"VSC email failed: {e}"

            print(f"HELD  {candidate_name} → {job_title} | Job not in DB | {_hold_email}")
            return jsonify({
                'success':      False,
                'held':         True,
                'candidate_id': hold_id,
                'job_title':    job_title,
                'reason':       f"Job '{job_title}' not found in database — application held for VSC review",
                'action':       'Upload updated CSV at /manual, then reprocess manually',
            }), 202

        #  Step 2: Resolve candidate location for geosync
        # Use pre-extracted values from payload (body + resume pass) first;
        # fall back to resume-only quick extract when payload is blank.
        city  = candidate_city_in  or ''
        state = candidate_state_in or ''
        if not city or not state:
            city, state = quick_extract_location(resume_text)
        alt_jobs = []
        if city and state:
            cand_lat, cand_lon = geocode_location(city, state)
            if cand_lat:
                db_nearby = get_jobs_near(cand_lat, cand_lon, radius_miles=50)
                alt_jobs = [{
                    'Job Title':    j.get('job_title', ''),
                    'Company Name': j.get('company_name', ''),
                    'City':         j.get('city', ''),
                    'State':        j.get('state', ''),
                    'distance':     round(j.get('distance_miles', 0), 1),
                } for j in db_nearby]
                print(f" Geosync: {len(alt_jobs)} jobs within 50mi of {city}, {state}")

        # Fall back to job location when candidate location unavailable
        if not alt_jobs:
            job_city  = target_job.get('City', '')
            job_state = target_job.get('State', '')
            if job_city and job_state:
                job_lat, job_lon = geocode_location(job_city, job_state)
                if job_lat:
                    db_nearby = get_jobs_near(job_lat, job_lon, radius_miles=50)
                    alt_jobs = [{
                        'Job Title':    j.get('job_title', ''),
                        'Company Name': j.get('company_name', ''),
                        'City':         j.get('city', ''),
                        'State':        j.get('state', ''),
                        'distance':     round(j.get('distance_miles', 0), 1),
                    } for j in db_nearby]
                    print(f" Geosync: job location fallback ({job_city}) — {len(alt_jobs)} nearby alternatives")

        if not alt_jobs:
            alt_jobs = list(JOBS_DB[:50])
            print(f" Geosync unavailable — passing {len(alt_jobs)} jobs for alternatives")

        #  Step 3: Run AI council analysis
        with open('webhook_debug.log', 'a') as _dbg:
            _dbg.write(f"[BEFORE ANALYZE] candidate={candidate_name}\n")
        raw_analysis, track_credits = analyze_for_vsc(resume_text, candidate_name, target_job, alt_jobs)
        with open('webhook_debug.log', 'a') as _dbg:
            _dbg.write(f"[AFTER ANALYZE] raw_analysis type={type(raw_analysis)} truthy={bool(raw_analysis)}\n")
        if not raw_analysis:
            return jsonify({'success': False, 'error': 'AI analysis failed'}), 500

        #  Step 4: Parse structured output
        parsed = parse_vsc_analysis(raw_analysis, credits=track_credits)
        grade = parsed.get('grade', 'N/A')
        candidate_city = parsed.get('candidate_city') or city or 'Unknown'

        #  Step 5: Create CRM record
        mil = _extract_military_profile(resume_text)
        print(f"  Military profile: branch={mil['branch'] or '?'}  rank={mil['rank'] or '?'}  mos={mil['mos'] or '?'}  years={mil['years_served'] or '?'}")

        candidate_id = None
        try:
            name_parts = candidate_name.split(' ', 1)
            first = name_parts[0]
            last  = name_parts[1] if len(name_parts) > 1 else ''
            candidate_id = create_candidate({
                'vsc_name':     vsc_name,
                'first_name':   first,
                'last_name':    last,
                'email':        candidate_email,
                'phone':        candidate_phone,
                'address':      candidate_address,
                'city':         candidate_city,
                'state':        parsed.get('candidate_state', state or ''),
                'branch':       mil['branch'],
                'rank':         mil['rank'],
                'mos':          mil['mos'],
                'years_served': mil['years_served'],
                'stage':        'RESUME_ANALYZED',
                'notes':        (
                    f"Applied for: {job_title}\n"
                    f"Eligibility: {parsed.get('eligibility', '')}\n"
                    f"Grade: {grade} | True Fit: {parsed.get('true_fit_grade', '')} | Shadow: {parsed.get('shadow_grade', 'N/A')}\n"
                    f"Classification: {parsed.get('classification', '')}\n\n"
                    f"VERIFICATION REQUIRED:\n{parsed.get('verification_required', '')}"
                ),
                'resume_text':  resume_text,
                'source':       'intake',
            })
            print(f" CRM record created: ID {candidate_id}")
            # Auto-log INTAKE engagement
            log_engagement(
                candidate_id=candidate_id,
                vsc_name=vsc_name,
                eng_type='INTAKE',
                notes=f"Application received for {job_title}. AI Grade: {grade}."
            )
        except Exception as e:
            print(f"  CRM save failed: {e}")

        #  Step 6: Build and send VSC analysis report email
        if GRAPH_CLIENT_SECRET:
            try:
                from data.email_templates import render_template as render_email
                report = render_email('vsc_analysis_report', {
                    'candidate_name':        candidate_name,
                    'candidate_email':       candidate_email or '(email not provided)',
                    'vsc_name':              vsc_name,
                    'job_title':             job_title,
                    'company_name':          target_job.get('Company Name', 'Not specified'),
                    'eligibility':           parsed.get('eligibility', ''),
                    'requirements':          parsed.get('requirements', ''),
                    'primary_grade':         parsed.get('primary_grade', grade),
                    'true_fit_grade':        parsed.get('true_fit_grade', ''),
                    'shadow_grade':          parsed.get('shadow_grade', '') or 'N/A',
                    'classification':        parsed.get('classification', ''),
                    'justification':         parsed.get('justification', ''),
                    'missing_requirements':  parsed.get('missing_requirements', ''),
                    'verification_required': parsed.get('verification_required', ''),
                    'improvements':          parsed.get('improvements', ''),
                    'ats_keywords':          parsed.get('ats_keywords', ''),
                    'ats_resume':            parsed.get('ats_resume', ''),
                    'candidate_city':        candidate_city,
                    'alt_1_title':           parsed.get('alt_1_title', ''),
                    'alt_1_company':         parsed.get('alt_1_company', ''),
                    'alt_1_distance':        parsed.get('alt_1_distance', ''),
                    'alt_1_score':           parsed.get('alt_1_score', ''),
                    'alt_1_why':             parsed.get('alt_1_why', ''),
                    'alt_2_title':           parsed.get('alt_2_title', ''),
                    'alt_2_company':         parsed.get('alt_2_company', ''),
                    'alt_2_distance':        parsed.get('alt_2_distance', ''),
                    'alt_2_score':           parsed.get('alt_2_score', ''),
                    'alt_2_why':             parsed.get('alt_2_why', ''),
                    'alt_3_title':           parsed.get('alt_3_title', ''),
                    'alt_3_company':         parsed.get('alt_3_company', ''),
                    'alt_3_distance':        parsed.get('alt_3_distance', ''),
                    'alt_3_score':           parsed.get('alt_3_score', ''),
                    'alt_3_why':             parsed.get('alt_3_why', ''),
                    'eligibility_model':     parsed.get('eligibility_model',  'N/A'),
                    'development_model':     parsed.get('development_model',  'N/A'),
                    'verification_model':    parsed.get('verification_model', 'N/A'),
                    'opportunities_model':   parsed.get('opportunities_model','N/A'),
                })
                _graph_send_email(
                    to_addr  = vsc_email,
                    to_name  = vsc_name,
                    subject  = report['subject'],
                    body     = report['body'],
                    reply_to = OUTLOOK_USER,
                )
                print(f" Analysis report sent to {vsc_name} ({vsc_email})")
            except Exception as e:
                print(f"  VSC email failed: {e}")

        with open('webhook_debug.log', 'a') as _dbg:
            _dbg.write(f"[RESULT] SUCCESS candidate={candidate_name} grade={grade} id={candidate_id}\n")
        return jsonify({
            'success':      True,
            'candidate_id': candidate_id,
            'grade':        grade,
            'grade_summary': parsed.get('grade_summary', ''),
            'job_matched':  bool(matched_job),
            'job_title':    job_title,
            'candidate_city': candidate_city,
        })

    except Exception as e:
        import traceback as _tb
        _trace = _tb.format_exc()
        print(f" Intake process error: {e}\n{_trace}")
        with open('webhook_debug.log', 'a') as _dbg:
            _dbg.write(f"[RESULT] EXCEPTION: {e}\n{_trace}\n")
        return jsonify({'success': False, 'error': str(e)}), 500


#  Engagement API

@app.route('/api/candidates/<int:cid>/engagements', methods=['GET'])
@api_login_required
def engagements_list(cid):
    """Return all engagements for a candidate."""
    return jsonify({
        'engagements': get_engagements(cid),
        'types': ENGAGEMENT_TYPES,
        'subtypes': ENGAGEMENT_SUBTYPES,
        'type_labels': ENGAGEMENT_TYPE_LABELS,
        'subtype_labels': ENGAGEMENT_SUBTYPE_LABELS,
    })


@app.route('/api/candidates/<int:cid>/engagements', methods=['POST'])
@api_login_required
def engagements_create(cid):
    """Log a new engagement for a candidate."""
    data = request.json or {}
    eng_type = (data.get('type') or '').upper()
    subtype  = (data.get('subtype') or '').upper() or None
    notes    = data.get('notes', '')
    vsc_name = session['vsc_name']

    if eng_type not in ENGAGEMENT_TYPES:
        return jsonify({'error': f'Invalid type. Must be one of: {ENGAGEMENT_TYPES}'}), 400

    valid_subtypes = ENGAGEMENT_SUBTYPES.get(eng_type, [])
    if subtype and subtype not in valid_subtypes:
        return jsonify({'error': f'Invalid subtype for {eng_type}. Must be one of: {valid_subtypes}'}), 400

    eid = log_engagement(candidate_id=cid, vsc_name=vsc_name,
                         eng_type=eng_type, subtype=subtype, notes=notes)
    return jsonify({'success': True, 'id': eid})


#  Ghost Detection 

def run_ghost_check():
    """
    Check for candidates with no engagement in 7 days.
    Emails each VSC a reminder for each stale candidate.
    Runs daily in a background thread.
    """
    with app.app_context():
        try:
            stale = get_candidates_needing_followup(days=7)
            if not stale:
                print(" Ghost check: no stale candidates")
                threading.Timer(86400, run_ghost_check).start()
                return

            print(f" Ghost check: {len(stale)} candidates need follow-up")

            if not OUTLOOK_PASSWORD:
                print("  Ghost check: email not configured, skipping notifications")
                threading.Timer(86400, run_ghost_check).start()
                return

            # Group by VSC
            by_vsc = {}
            for c in stale:
                vsc = c.get('vsc_name', 'VSC')
                by_vsc.setdefault(vsc, []).append(c)

            for vsc_name, candidates in by_vsc.items():
                lines = []
                for c in candidates:
                    last = c.get('last_engagement_at', 'Never')
                    lines.append(
                        f"   {c['first_name']} {c['last_name']} "
                        f"| Stage: {STAGE_LABELS.get(c['stage'], c['stage'])} "
                        f"| Last engagement: {last[:10] if last and last != 'Never' else 'Never'} "
                        f"| Email: {c.get('email', 'N/A')}"
                    )

                subject = f"Follow-Up Required  {len(candidates)} Candidate{'s' if len(candidates) > 1 else ''} Awaiting Contact"
                body = f"""FOLLOW-UP REMINDER
Work for Warriors  Automated Alert

{len(candidates)} candidate{'s have' if len(candidates) > 1 else ' has'} had no engagement in 7 or more days.

{chr(10).join(lines)}

Log in to your CRM to review and take action.

---
Work for Warriors Resume AI  Automated Notification"""

                try:
                    msg = MIMEMultipart()
                    msg['From']    = f'Work for Warriors AI <{OUTLOOK_USER}>'
                    msg['To']      = OUTLOOK_USER  # VSC email would go here when per-VSC emails are configured
                    msg['Subject'] = subject
                    msg.attach(MIMEText(body, 'plain'))

                    with smtplib.SMTP('smtp.office365.com', 587) as server:
                        server.ehlo()
                        server.starttls()
                        server.login(OUTLOOK_USER, OUTLOOK_PASSWORD)
                        server.send_message(msg)
                    print(f" Ghost reminder sent for {vsc_name}: {len(candidates)} candidates")
                except Exception as e:
                    print(f"  Ghost reminder email failed for {vsc_name}: {e}")

        except Exception as e:
            print(f" Ghost check error: {e}")

    # Schedule next run in 24 hours
    threading.Timer(86400, run_ghost_check).start()


@app.route('/api/ghost_check', methods=['POST'])
@api_login_required
def ghost_check_manual():
    """Manually trigger ghost detection (for testing or Power Automate scheduled call)."""
    threading.Thread(target=run_ghost_check, daemon=True).start()
    return jsonify({'success': True, 'message': 'Ghost check triggered'})


# ── Job Import Folder Watcher ─────────────────────────────────────────────────
# Watches data/job_imports/ for new CSV files dropped by the weekly Claude Code
# schedule. Auto-loads them exactly like a manual upload, then moves the file
# to data/job_imports/processed/ so it isn't re-loaded next cycle.

JOB_IMPORTS_DIR   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "job_imports")
JOB_IMPORTS_DONE  = os.path.join(JOB_IMPORTS_DIR, "processed")
FOLDER_WATCH_INTERVAL = 60  # seconds


def _auto_load_csv(filepath):
    """Load a single jobs CSV file into the DB and in-memory cache."""
    global JOBS_DB
    import io
    with open(filepath, 'rb') as f:
        raw = f.read()
    df = None
    for enc in ('utf-8-sig', 'utf-8', 'cp1252', 'latin-1'):
        try:
            df = pd.read_csv(io.BytesIO(raw), encoding=enc)
            break
        except Exception:
            continue
    if df is None:
        print(f"Folder watcher: could not read {filepath}")
        return
    df.columns = [c.strip() for c in df.columns]
    required = ['Job Title', 'City', 'State', 'Job Description']
    if any(col not in df.columns for col in required):
        print(f"Folder watcher: {filepath} missing required columns, skipping")
        return
    df['latitude']  = None
    df['longitude'] = None
    jobs = df.to_dict('records')
    saved = save_jobs(jobs)
    log_job_upload(saved)
    JOBS_DB = jobs
    print(f"Folder watcher: auto-loaded {saved} jobs from {os.path.basename(filepath)}")


def run_folder_watcher():
    """Poll job_imports/ for new CSV files and auto-load them."""
    os.makedirs(JOB_IMPORTS_DIR,  exist_ok=True)
    os.makedirs(JOB_IMPORTS_DONE, exist_ok=True)
    try:
        for fname in os.listdir(JOB_IMPORTS_DIR):
            if not fname.lower().endswith('.csv'):
                continue
            fpath = os.path.join(JOB_IMPORTS_DIR, fname)
            if not os.path.isfile(fpath):
                continue
            try:
                _auto_load_csv(fpath)
                done_path = os.path.join(JOB_IMPORTS_DONE, fname)
                os.rename(fpath, done_path)
            except Exception as e:
                print(f"Folder watcher: error processing {fname}: {e}")
    except Exception as e:
        print(f"Folder watcher: error: {e}")
    threading.Timer(FOLDER_WATCH_INTERVAL, run_folder_watcher).start()


# ── Graph API Inbox Poller ────────────────────────────────────────────────────
# Watches the VSC inbox via Microsoft Graph (modern auth — no IMAP/Basic Auth).
# Every 5 minutes, finds unread "Application for" emails and feeds them
# through the /api/intake/process pipeline.

INBOX_POLL_INTERVAL = 300  # seconds
_failed_msg_ids: set = set()  # message IDs that failed this session — skip on retry


_SKIP_SUBJECTS = (
    'undeliverable', 'delivery failure', 'mail delivery failed',
    'out of office', 'automatic reply', 'auto-reply', 'autoreply',
    'read receipt', 'non-delivery',
)
_SKIP_SENDERS = ('microsoft outlook', 'postmaster', 'mailer-daemon')


def _process_inbox_message(token, msg):
    """Process a single inbox message through the pipeline (runs in its own thread)."""
    msg_id   = msg["id"]
    subject  = msg.get("subject", "")

    if msg_id in _failed_msg_ids:
        print(f"Inbox poll: skipping previously failed message '{subject[:50]}' — restart app to retry")
        return
    from_obj = msg.get("from", {}).get("emailAddress", {})
    raw_from_addr = from_obj.get("address", "")
    candidate_name = from_obj.get("name", "") or raw_from_addr.split("@")[0]

    subject_lc = subject.lower()
    sender_lc  = candidate_name.lower()
    if any(s in subject_lc for s in _SKIP_SUBJECTS) or any(s in sender_lc for s in _SKIP_SENDERS):
        print(f"Inbox poll: skipping system/bounce email '{subject[:60]}'")
        _graph_mark_read(token, msg_id)
        return

    # Always capture email body — it often contains cover letter content,
    # listed qualifications, or tool experience not present in the resume.
    body_obj  = msg.get("body", {})
    body_text = body_obj.get("content", "")
    if body_obj.get("contentType", "").lower() == "html":
        body_text = re.sub(r"<[^>]+>", " ", body_text)
        body_text = re.sub(r"\s+", " ", body_text).strip()

    # ── Extract real candidate email ────────────────────────────────────────
    # SJB relays applications from wfw@mysmartjobboard.com — the actual
    # candidate email is in the Reply-To header or body text.
    candidate_email_addr = ""
    # 1. Check Reply-To header (most reliable for SJB)
    reply_to_list = msg.get("replyTo", [])
    if reply_to_list:
        rt_addr = reply_to_list[0].get("emailAddress", {}).get("address", "")
        if rt_addr and "mysmartjobboard" not in rt_addr.lower():
            candidate_email_addr = rt_addr
    # 2. Parse from email body (SJB typically embeds "Email: user@domain.com")
    if not candidate_email_addr and body_text:
        m = re.search(
            r'(?:E-?mail|Email address)[:\s]+([a-zA-Z0-9._%+\-]+@(?!mysmartjobboard)[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})',
            body_text, re.IGNORECASE
        )
        if m:
            candidate_email_addr = m.group(1).strip()
    # 3. Fallback: any non-SJB email address in the body
    if not candidate_email_addr and body_text:
        for match in re.finditer(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', body_text):
            addr = match.group(0)
            if "mysmartjobboard" not in addr and "workforwarriors" not in addr.lower():
                candidate_email_addr = addr
                break
    # 4. Last resort: use the raw from address (likely SJB relay, better than empty)
    if not candidate_email_addr:
        candidate_email_addr = raw_from_addr

    print(f"  Email resolved: {candidate_email_addr or '(none)'}")

    # Get the resume from attachment
    resume_text = _fetch_attachment_text(token, msg_id)

    # Combine: resume + email body so the AI sees everything the candidate sent
    if resume_text and body_text:
        resume_text = f"{resume_text}\n\n--- CANDIDATE APPLICATION MESSAGE ---\n{body_text}"
        print(f"  Resume + email body combined ({len(resume_text)} chars)")
    elif body_text:
        resume_text = body_text
        print(f"  No attachment — using email body only ({len(resume_text)} chars)")

    if not resume_text:
        print(f"Inbox poll: no resume content in '{subject[:50]}' — skipping")
        _graph_mark_read(token, msg_id)
        return

    # Extract all contact fields (phone, address, city, state) across every
    # available source in a single unified pass — body and resume both searched.
    contact = _extract_contact_fields(body_text, resume_text)
    if contact['phones']:
        print(f"  Phone(s): {contact['phones']}")
    if contact['address']:
        print(f"  Address:  {contact['address']}")

    payload = {
        "subject_line":      subject,
        "resume_text":       resume_text,
        "candidate_name":    candidate_name,
        "candidate_email":   candidate_email_addr,
        "candidate_phone":   contact['phones'],
        "candidate_address": contact['address'],
        "candidate_city":    contact['city'],
        "candidate_state":   contact['state'],
        "vsc_name":          VSC_DISPLAY_NAME,
        "vsc_email":         OUTLOOK_USER,
    }
    success = False
    try:
        r = requests.post(
            "http://127.0.0.1:5001/api/intake/process",
            json=payload,
            timeout=300,
        )
        if r.ok:
            success = True
            print(f"Inbox poll: processed '{subject[:50]}' — grade {r.json().get('grade','?')}")
        else:
            print(f"Inbox poll: pipeline error ({r.status_code}): {r.text[:200]}")
    except Exception as e:
        print(f"Inbox poll: pipeline POST error: {e}")

    if success:
        _graph_mark_read(token, msg_id)
    else:
        _failed_msg_ids.add(msg_id)
        print(f"Inbox poll: '{subject[:50]}' failed — marked to skip until restart")


def run_inbox_poll():
    """Poll O365 inbox via Graph API for unread application emails."""
    if not all([GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, OUTLOOK_USER]):
        print("Inbox poll: Graph credentials not set — skipping")
        threading.Timer(INBOX_POLL_INTERVAL, run_inbox_poll).start()
        return

    try:
        token = _graph_get_token()

        url = (
            f"https://graph.microsoft.com/v1.0/users/{OUTLOOK_USER}/messages"
            f"?$filter=isRead eq false and contains(subject,'Application for')"
            f"&$select=id,subject,from,body,hasAttachments"
            f"&$top=25"
        )
        resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
        resp.raise_for_status()
        messages = resp.json().get("value", [])

        if not messages:
            threading.Timer(INBOX_POLL_INTERVAL, run_inbox_poll).start()
            return

        print(f"Inbox poll: {len(messages)} unread application email(s) — processing in parallel")

        threads = [
            threading.Thread(target=_process_inbox_message, args=(token, msg), daemon=True)
            for msg in messages
        ]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        print(f"Inbox poll: batch complete ({len(messages)} emails)")

    except Exception as e:
        print(f"Inbox poll: Graph API error: {e}")

    threading.Timer(INBOX_POLL_INTERVAL, run_inbox_poll).start()


@app.route('/api/inbox/poll', methods=['POST'])
@api_login_required
def inbox_poll_manual():
    """Manually trigger an inbox poll (for testing)."""
    threading.Thread(target=run_inbox_poll, daemon=True).start()
    return jsonify({'success': True, 'message': 'Inbox poll triggered'})


if __name__ == '__main__':
    _graph_ok = all([GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET])
    print(
        f"\nRESUME AI  Work for Warriors"
        f" | Jobs: {len(JOBS_DB)}"
        f" | API: {'OK' if DEEPSEEK_API_KEY else 'MISSING'}"
        f" | Graph: {'OK' if _graph_ok else 'NOT SET'}"
        f" | http://localhost:5001\n"
    )
    threading.Timer(15,  run_folder_watcher).start()  # 15s — fast, just watches a directory
    threading.Timer(60,  run_inbox_poll).start()       # 60s — give Flask time to be fully ready
    threading.Timer(120, run_ghost_check).start()      # 2min — send after inbox poll, not before

    app.run(host='0.0.0.0', port=5001, debug=True, use_reloader=False)
