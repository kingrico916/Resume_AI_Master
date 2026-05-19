# Resume AI Master — Claude Handoff Document
**Project path:** `C:\Users\antho\OneDrive\Desktop\resume_ai_master`  
**Last updated:** 2026-04-02  
**Owner:** Anthony Antonucci (IP is 100% his, personally funded)

---

## One-Line Summary
Flask-based automated veteran resume analysis and placement platform for Work for Warriors California. 7 VSCs across CA. San Diego demo upcoming to pitch to other state contractors.

---

## Core Workflow (DO NOT DEVIATE FROM THIS)
```
Candidate applies on SmartJobBoard (SJB)
  → Email lands in VSC's O365 inbox with "Application for [Job Title]" in subject
  → Power Automate monitors inbox, POSTs to /api/powerautomate
  → Flask parses job title from subject, looks it up in DB
  → Extracts candidate location from resume text (regex, first 600 chars)
  → Geocodes candidate location → finds 3 nearby jobs within 50 miles
  → AI (OpenRouter → claude-3.5-sonnet) grades resume per master prompt
  → Parses structured AI output into fields
  → Creates CRM record at stage RESUME_ANALYZED
  → Auto-logs INTAKE engagement
  → Emails VSC the analysis report (Email 1 — only automated email)
  → VSC uses CRM to manage candidate, sends emails using templates
```

**Key rules:**
- VSC never sees two automated emails — only one (the AI analysis report)
- Alternatives (3 nearby jobs) are ALWAYS sent regardless of grade A–F
- Power Automate (not Graph API) — each VSC sets up their own flow
- Salesforce may come organizationally — build AI engine as standalone IP
- Ghost detection fires at 7 days of no engagement (not 14)

---

## Architecture

### Files That Matter
| File | Purpose |
|------|---------|
| `app.py` | Flask app, all routes, AI functions, ghost detection |
| `data/database.py` | SQLite schema, all DB functions, constants |
| `data/email_templates.py` | 10 email templates (9 VSC-to-candidate + 1 VSC analysis report) |
| `templates/crm.html` | CRM UI — candidate pipeline, engagements, emails, submissions |
| `templates/email_manager.html` | Email compose UI with templates + custom compose |
| `templates/index.html` | Manual mode — VSC-operated resume AI + CSV job loader |

### Database Tables (SQLite at `data/awis.db`)
```
email_logs        — every email sent (vsc_name, to_name, to_email, template_key, subject, body, status)
candidates        — CRM records (name, contact, military info, stage, notes, source, vsc_name)
submissions       — employer submissions per candidate (job_title, company_name, outcome)
job_listings      — master job DB loaded from SJB CSV (geocoded lat/lon)
engagements       — every VSC–candidate contact (type, subtype, notes, resets ghost clock)
custom_templates  — VSC-saved custom email templates (persisted across sessions)
```

### Key Constants (data/database.py)
```python
PIPELINE_STAGES = ['NEW', 'CONTACTED', 'MEETING_SCHEDULED', 'MEETING_COMPLETE',
                   'RESUME_ANALYZED', 'SUBMITTED', 'HIRED', 'REJECTED', 'WITHDRAWN']

ENGAGEMENT_TYPES = ['INTAKE', 'COUNSELING', 'MEETING', 'COACHING', 'SUBMITTED']

ENGAGEMENT_SUBTYPES = {
    'INTAKE':     [],
    'COUNSELING': ['PHONE', 'EMAIL', 'IN_PERSON'],
    'MEETING':    ['VIRTUAL', 'IN_PERSON', 'PHONE', 'MISSED'],
    'COACHING':   ['PHONE', 'EMAIL', 'IN_PERSON', 'MOCK_INTERVIEW', 'RESUME_REVIEW', 'INTERVIEW_PREP'],
    'SUBMITTED':  [],
}
```

WITHDRAWN is reachable from any pipeline stage — all stage buttons are always clickable.

---

## AI Analysis (app.py)

### `analyze_for_vsc(resume_text, candidate_name, target_job, alternative_jobs)`
Sends to OpenRouter → `anthropic/claude-3.5-sonnet`. Returns structured text with these labeled sections:
```
LOCATION: city, state
GRADE: A/A-/B/C/D/F
GRADE_SUMMARY: one-line
JUSTIFICATION: paragraph
MISSING_REQUIREMENTS: bullet list
RECOMMENDATION: VSC action guidance
RESUME_REWRITE: full ATS-optimized resume
ALTERNATIVE_1: Title | Company | Distance | Score | Why
ALTERNATIVE_2: ...
ALTERNATIVE_3: ...
```

### `parse_vsc_analysis(raw)`
Regex parser that converts the above structured text to a Python dict.

### `quick_extract_location(resume_text)`
Regex scan of first 600 chars of resume. Looks for "City, ST" patterns. Returns `(city, state)`.

---

## Email Templates (data/email_templates.py)
10 templates total. All in `TEMPLATES` list, keyed in `TEMPLATES_BY_KEY`.

| Key | Purpose |
|-----|---------|
| `initial_contact` | First outreach, requests military info |
| `mission_followup` | Pre-meeting engagement |
| `interview_invitation` | Has job_title, company, date, time, location fields |
| `post_meeting` | Post-meeting action plan |
| `no_show` | Reschedule after missed meeting |
| `submission_confirmation` | Confirms employer submission |
| `post_submission_checkin` | Sets communication expectations |
| `rejection_support` | Candidate not selected |
| `closed_position` | Role no longer open |
| `vsc_analysis_report` | **Internal only** — AI output emailed to VSC. Excluded from Email Manager template list. |

Templates use `{placeholder}` syntax. `render_template(key, variables)` in email_templates.py handles substitution.

---

## API Routes (app.py)

### Power Automate Intake
```
POST /api/powerautomate
Body: { subject, body (resume text), from_email, from_name, vsc_name }
```

### CRM
```
GET  /api/candidates                    — list all (with submission_count, last_email_at)
POST /api/candidates                    — create new
GET  /api/candidates/<id>               — detail + submissions + email history
PUT  /api/candidates/<id>               — update any fields (stage, notes, profile)
DELETE /api/candidates/<id>             — soft delete (is_deleted=1)
GET  /api/candidates/<id>/engagements   — engagement history
POST /api/candidates/<id>/engagements   — log engagement { type, subtype, notes, vsc_name }
POST /api/candidates/<id>/submissions   — add employer submission
PUT  /api/submissions/<id>              — update submission outcome
```

### Email
```
GET  /api/templates                     — all templates (built-in + custom, excludes vsc_analysis_report)
POST /api/templates                     — save a VSC custom template
DELETE /api/templates/<key>             — delete a custom template (key must start with "custom_")
POST /api/preview_email                 — render template with variables, return subject+body
POST /api/send_email                    — send email, log to email_logs
GET  /api/email_history                 — recent email log
```

### Jobs
```
POST /load_jobs                         — upload CSV, starts background geocode thread, returns immediately
GET  /api/jobs/status                   — poll geocoding progress { state, progress, total, geocoded, in_memory }
GET  /api/jobs/search?q=                — search job titles
```

### Ghost Detection
```
POST /api/ghost_check                   — manual trigger
```
Ghost check also runs automatically every 24 hours via `threading.Timer(5, run_ghost_check)` at startup.

---

## CRM UI (templates/crm.html)

### Detail Panel Tabs
- **Profile** — name, contact, military background, VSC assigned. Save button PUTs to API.
- **Engagements** — log type/subtype/notes. History table. Every log resets the 7-day ghost clock.
- **Notes** — internal free-text, not visible to candidate.
- **Emails** — history of all emails sent to this candidate.
- **Submissions** — employer submissions with outcome dropdown (PENDING/INTERVIEW/HIRED/REJECTED).

### JS Constants (injected by Flask via Jinja2)
```javascript
const STAGES = ...          // from pipeline_stages
const STAGE_LABELS = ...    // from stage_labels
const STAGE_COLORS = ...    // from stage_colors
const ENG_TYPES = ...       // from engagement_types
const ENG_SUBTYPES = ...    // from engagement_subtypes
const ENG_TYPE_LABELS = ... // from engagement_type_labels
const ENG_SUBTYPE_LABELS = ...
```
Engagement type dropdown is dynamically built from `ENG_TYPES` on `DOMContentLoaded`.

---

## Email Manager (templates/email_manager.html)

Three modes:
1. **Template** — select a built-in template, fill optional fields (job title, company, interview details), preview populates subject/body, VSC can edit freely before sending.
2. **Custom Saved** — VSC-created templates stored in `custom_templates` DB table. Pre-fill subject/body on select. Orange "Saved" badge. Has ✕ delete button.
3. **Custom Email (blank)** — "+ Custom Email" button opens blank compose. No template required.

**Save as Template** button in compose section — prompts for name + when-to-use, POSTs to `/api/templates`, appears immediately in template list.

---

## Job Loading (large CSV)
Previous issue: Nominatim geocoding blocked the browser request (10k rows × 1 sec/req = hours).

Current fix:
- `/load_jobs` accepts the CSV, reads it, starts a **background thread** for geocoding, returns immediately (202)
- Background thread geocodes unique city/state combos only (100–200 locations vs 10k rows)
- `save_jobs()` uses `executemany()` for batch DB insert
- Frontend polls `/api/jobs/status` every 2 seconds and shows live progress ("Geocoding 47/112: San Diego, CA...")

---

## Environment Variables (Windows System Environment — NOT .env file)
```
OPENROUTER_API_KEY   — OpenRouter API key
OUTLOOK_USER         — anthony.antonucci@workforwarriors.org
OUTLOOK_PASSWORD     — O365 app password for SMTP
```
These are set as Windows system env vars. No `.env` file. Do not add `load_dotenv()`.

---

## Run the App
```
cd C:\Users\antho\OneDrive\Desktop\resume_ai_master
python app.py
```
Opens at **http://localhost:5000**

Routes:
- `/` — home/mode selector
- `/manual` — manual VSC resume AI + job CSV loader
- `/crm` — candidate pipeline CRM
- `/email-manager` — email compose

---

## Pending / Not Yet Built
- **VSC identity/login** — currently VSC types their name manually per session. No auth. Future: dropdown of 7 VSC names persisted in localStorage or a simple login.
- **Analytics dashboard** — Ryan's management view (pipeline counts, placement rates, VSC activity). Route `/dashboard` exists as placeholder.
- **Power Automate setup docs** — step-by-step for each VSC to configure their own flow (monitor inbox → POST to `/api/powerautomate`).
- **Live deployment** — Railway or Render for San Diego demo. App is ready to deploy, needs `requirements.txt` verification and a production WSGI config.
- **Salesforce integration** — not started. AI engine is designed as standalone so it can feed into Salesforce later without a rewrite.

---

## Style / Language Rules (enforce these)
- Templates use direct, professional military-adjacent language. No passive voice ("Please reply" → "Reply with", "Feel free to" → direct action).
- You're running a pipeline, not asking for favors.
- No "job search journey." Just "job search."
- Closings: "Respectfully" for formal, "Best" for casual. Never "Warm regards."
