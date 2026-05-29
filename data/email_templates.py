"""
data.email_templates
====================
All nine Work for Warriors VSC email templates.

Each template is a dict with:
    key         str   — machine identifier
    label       str   — display name
    stage       str   — pipeline stage this template maps to
    when        str   — one-line guidance on when to send
    subject     str   — email subject (may contain placeholders)
    body        str   — email body (may contain placeholders)
    fields      list  — which optional fields the template uses
                        possible values: job_title, company_name,
                        interview_date, interview_time, interview_location

Placeholders replaced at send time:
    {candidate_name}        Veteran's name
    {vsc_name}              VSC full name
    {job_title}             Job title (optional)
    {company_name}          Company name (optional)
    {interview_date}        Interview date (optional)
    {interview_time}        Interview time (optional)
    {interview_location}    Interview location / format (optional)
"""

TEMPLATES = [
    {
        "key": "initial_contact",
        "label": "Initial Contact & Information Request",
        "stage": "NEW LEAD",
        "when": "Beginning the intake process.",
        "fields": ["job_title"],
        "subject": "Next Steps – Work For Warriors",
        "body": """Dear {candidate_name},

Thank you for connecting with Work For Warriors and your interest in the {job_title} opportunity.

I will be your point of contact throughout this process. Reply with the following information prior to our meeting:

  • Full Name
  • Phone Number
  • Email Address
  • Full Home Address
  • Are you currently serving, a Veteran, or a family member of a Veteran?
  • Branch of service and years served
  • Military occupational specialty (MOS/AFSC/NEC/Rate)
  • Highest rank achieved
  • City and state you are currently located in
  • Any certifications, licenses, or degrees you hold

You will find my scheduling link in my email signature below.

Respectfully,

{vsc_name}
Veteran Services Coordinator
Work for Warriors"""
    },
    {
        "key": "mission_followup",
        "label": "Mission-Aligned Pre-Meeting Engagement",
        "stage": "CONTACTED",
        "when": "After interest is expressed or application received — before a meeting is scheduled.",
        "fields": ["job_title", "scheduling_link"],
        "subject": "Next Steps with Work For Warriors",
        "body": """Dear {candidate_name},

Thank you for connecting with Work For Warriors and for your interest in the {job_title} opportunity.

Our next step is to schedule a one-on-one session to review your résumé, discuss your career goals, and align you with the right opportunities.

Browse the Smart Jobs Board and bring any roles you'd like to discuss to your session.

Schedule here: {scheduling_link}

Respectfully,

{vsc_name}
Veteran Services Coordinator
Work for Warriors"""
    },
    {
        "key": "interview_invitation",
        "label": "Military-Style Interview Invitation",
        "stage": "INTERVIEW SCHEDULED",
        "when": "When inviting a candidate to interview for a specific position.",
        "fields": ["job_title", "company_name", "interview_date", "interview_time", "interview_location"],
        "subject": "Interview Invitation – {job_title}",
        "body": """Dear {candidate_name},

Thank you for connecting with Work For Warriors and your interest in the {job_title} position. After reviewing everything, I encourage you to move forward in the application process and would like to schedule a time to communicate the options available to you.

Interview Details:
  Date:          {interview_date}
  Time:          {interview_time}
  Location:      {interview_location}
  Company:       {company_name}

Confirm your availability or advise if an alternate time is needed.

Respectfully,

{vsc_name}
Veteran Services Coordinator"""
    },
    {
        "key": "post_meeting",
        "label": "Post-Meeting Action Plan",
        "stage": "MEETING COMPLETE",
        "when": "After completing a one-on-one meeting with the candidate.",
        "fields": [],
        "subject": "Next Steps After Our Meeting",
        "body": """Dear {candidate_name},

Thank you for connecting with Work For Warriors. A few things to keep in mind as you move forward:

Our Smart Jobs Board is your primary resource. Check it regularly and continue applying to positions that interest you. Once you apply, the assigned Veteran Staffing Consultant will reach out to guide you through the next steps.

When you identify a role you want to pursue, I can help tailor your application and advocate with the employer on your behalf. Keep me informed of any applications you submit, interviews you complete, or offers you receive — positive or negative.

If you get stuck or need support at any point, reach out.

Respectfully,

{vsc_name}
Veteran Services Coordinator
Work for Warriors"""
    },
    {
        "key": "no_show",
        "label": "No-Show Reschedule Invitation",
        "stage": "MEETING SCHEDULED",
        "when": "When a candidate misses their scheduled meeting.",
        "fields": [],
        "subject": "Rescheduling Our Meeting",
        "body": """Dear {candidate_name},

Thank you for allowing me to assist in your employment search. I noticed we missed our scheduled meeting. I understand things come up and would still like to connect.

Use the scheduling link in my signature to select a new time. If I don't hear back after a couple of follow-ups, I'll assume you're no longer pursuing support at this time — but you're always welcome to reconnect with Work For Warriors should the need arise.

Respectfully,

{vsc_name}
Veteran Services Coordinator
Work for Warriors"""
    },
    {
        "key": "submission_confirmation",
        "label": "Submission Confirmation & Communication Expectations",
        "stage": "SUBMITTED",
        "when": "Confirming submission to an employer.",
        "fields": ["job_title", "company_name"],
        "subject": "Submission Confirmation – {job_title}",
        "body": """Hi {candidate_name},

I've officially submitted you for the {job_title} position with {company_name}.

At this stage, the employer will take over the process. If you're contacted for an interview or receive any updates, keep Work For Warriors informed.

If any feedback is provided by the employer, please share it so we can apply it to future candidates moving forward.

Best,

{vsc_name}
Veteran Services Coordinator
Work for Warriors"""
    },
    {
        "key": "post_submission_checkin",
        "label": "Candidate Submission & Communication Expectations",
        "stage": "SUBMITTED",
        "when": "Immediately after submitting a candidate — sets expectations on communication.",
        "fields": ["job_title", "company_name"],
        "subject": "Your Submission – {job_title}",
        "body": """Hey {candidate_name},

I've submitted you for the {job_title} opportunity and advocated on your behalf with the employer.

Due to union or confidentiality policies, Work For Warriors may not receive decision updates. If you hear anything — interview, offer, or start date — let us know via text or call.

In the event you're not selected, please share any significant feedback so we can strengthen future submissions.

I'm here if you need anything along the way.

{vsc_name}
Veteran Services Coordinator
Work for Warriors"""
    },
    {
        "key": "rejection_support",
        "label": "Rejection Support & Encouragement",
        "stage": "REJECTED",
        "when": "When a candidate was not selected for a role.",
        "fields": ["job_title", "company_name"],
        "subject": "Update on Your Application – {job_title}",
        "body": """Hi {candidate_name},

I wanted to follow up regarding the {job_title} position. The employer has chosen to move forward with other candidates.

If any significant feedback was shared, send it my way so we can strengthen future opportunities. Don't get discouraged — this is just one step in the process.

I'm here to support you however I can.

Best,

{vsc_name}
Veteran Services Coordinator
Work for Warriors"""
    },
    {
        "key": "closed_position",
        "label": "Closed Position – Encouragement Response",
        "stage": "GENERAL",
        "when": "When a candidate applies to or asks about a position that is no longer open.",
        "fields": ["job_title", "company_name"],
        "subject": "Update on Your Application – {job_title}",
        "body": """Dear {candidate_name},

Thank you for submitting your application for the {job_title} position through the Work For Warriors platform.

At this time, the specific role you applied for is no longer open. However, we truly appreciate your interest and encourage you to continue exploring other opportunities available through our Smart Jobs Board.

Please don't hesitate to reach out if you need assistance navigating the platform or identifying roles that may be a good fit. We're here to support you throughout your job search.

Respectfully,

{vsc_name}
Veteran Services Coordinator
Work for Warriors"""
    },
    {
        "key": "vsc_analysis_report",
        "label": "VSC Resume Analysis Report (AWIS v4.0)",
        "stage": "RESUME_ANALYZED",
        "when": "Sent automatically to the VSC when a new application is processed by the AI.",
        "fields": [
            "candidate_name", "candidate_email", "job_title", "company_name",
            "eligibility", "requirements",
            "primary_grade", "true_fit_grade", "shadow_grade", "classification",
            "justification", "missing_requirements", "verification_required",
            "improvements", "ats_keywords", "ats_resume", "candidate_city",
            "alt_1_title", "alt_1_company", "alt_1_distance", "alt_1_score", "alt_1_why",
            "alt_2_title", "alt_2_company", "alt_2_distance", "alt_2_score", "alt_2_why",
            "alt_3_title", "alt_3_company", "alt_3_distance", "alt_3_score", "alt_3_why",
            "eligibility_model", "development_model", "verification_model", "opportunities_model",
        ],
        "subject": "AWIS Report — {candidate_name} | {job_title} | {eligibility} | Grade: {primary_grade}",
        "body": """AWIS RESUME ANALYSIS REPORT — v4.0
Work for Warriors — Automated AI Review

CANDIDATE:    {candidate_name}
APPLIED FOR:  {job_title} at {company_name}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ELIGIBILITY DECISION                    [{eligibility_model}]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{eligibility}

CLASSIFICATION:    {classification}
PRIMARY GRADE:     {primary_grade}
TRUE FIT GRADE:    {true_fit_grade}
SHADOW GRADE:      {shadow_grade}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REQUIREMENT VALIDATION                  [{eligibility_model}]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{requirements}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
JUSTIFICATION                           [{development_model}]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{justification}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
MISSING REQUIREMENTS                    [{verification_model}]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{missing_requirements}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
VERIFICATION REQUIRED — VSC OUTREACH QUESTIONS      [{verification_model}]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{verification_required}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
KEY TERMS — FROM JOB POSTING            [{development_model}]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{ats_keywords}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RESUME IMPROVEMENTS                     [{development_model}]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{improvements}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ATS-OPTIMIZED RESUME REWRITE            [{development_model}]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{ats_resume}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ADDITIONAL OPPORTUNITIES — {candidate_city} AREA    [{opportunities_model}]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. {alt_1_title} — {alt_1_company}
   Distance: {alt_1_distance} | Fit Score: {alt_1_score}
   {alt_1_why}

2. {alt_2_title} — {alt_2_company}
   Distance: {alt_2_distance} | Fit Score: {alt_2_score}
   {alt_2_why}

3. {alt_3_title} — {alt_3_company}
   Distance: {alt_3_distance} | Fit Score: {alt_3_score}
   {alt_3_why}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SUITABLE → Submit  |  PENDING REVIEW → Verify first  |  NOT SUITABLE — FINAL → Stop
Candidate added to CRM. Use standard outreach templates to respond.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OPTIONAL — DRAFT CANDIDATE OUTREACH (use at your discretion)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
To: {candidate_name} <{candidate_email}>
Subject: Action Required — {job_title} Application

Dear {candidate_name},

Your application for the {job_title} position at {company_name} is under review. Before I submit you to the employer, I need to flag something important.

Most employers run résumés through an Applicant Tracking System (ATS) before a human ever sees the file. These systems scan for exact keywords. If a qualification is not explicitly stated on your résumé — even one you clearly possess — the system will automatically reject your application. You will never know it happened.

To prevent that, it's ocassionally helpful to ask additional questions. When this is the case, please reply promptly and confirm even if you have mentioned it previously.

{verification_required}

One line per item is enough. Once I have your response, I can update your materials and submit you with confidence.

Respectfully,
{vsc_name}
Veteran Services Coordinator
Work for Warriors"""
    },
]

# Quick lookup by key
TEMPLATES_BY_KEY = {t["key"]: t for t in TEMPLATES}


def get_template(key: str) -> dict:
    """Return template dict by key, or None if not found."""
    return TEMPLATES_BY_KEY.get(key)


def render_template(key: str, variables: dict) -> dict:
    """
    Return subject and body with all {placeholders} substituted.
    Missing variables are left as-is (not replaced).
    """
    tmpl = get_template(key)
    if not tmpl:
        raise ValueError(f"Unknown template key: {key!r}")

    subject = tmpl["subject"]
    body = tmpl["body"]

    for var, value in variables.items():
        placeholder = "{" + var + "}"
        subject = subject.replace(placeholder, str(value) if value else "")
        body    = body.replace(placeholder, str(value) if value else "")

    return {"subject": subject, "body": body}
