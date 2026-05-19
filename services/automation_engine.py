"""
services.automation_engine
===========================
Full automation chain triggered by candidate intake form submission.

Sequence:
  1. Get target job from DB
  2. Geocode candidate location
  3. Find nearby jobs from DB
  4. Run AI analysis
  5. Create candidate record in CRM (source='intake')
  6. Send initial contact email to candidate
  7. Generate CalCareers package if job is from calcareers board
  8. Send VSC notification email with full analysis
  9. Return result dict

Usage:
    from services.automation_engine import run_automation
    result = run_automation(intake_data, analyze_fn, geocode_fn, generate_package_fn)
"""

import re
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from data.database import (
    get_job_by_id,
    get_jobs_near,
    create_candidate,
    update_candidate,
    log_email,
)
from data.email_templates import render_template


def run_automation(
    intake_data: dict,
    analyze_fn,
    geocode_fn,
    generate_package_fn=None,
) -> dict:
    """
    Execute the full automation chain for a candidate intake submission.

    Parameters
    ----------
    intake_data : dict
        Keys: vsc_name, vsc_email, first_name, last_name, email, phone,
              city, state, branch, rank, mos, years_served, resume_text,
              target_job_id, education, work_experience, veterans_preference,
              template_track, outlook_user, outlook_password, ryan_email
    analyze_fn : callable
        analyze_with_ai(resume_text, name, location, target_job, alt_jobs) -> str
    geocode_fn : callable
        geocode_location(city, state) -> (lat, lon)
    generate_package_fn : callable or None
        generate_calcareers_package_bytes(intake_data) -> bytes or None

    Returns
    -------
    dict with keys:
        success          bool
        candidate_id     int or None
        grade            str or None
        analysis         str or None
        email_sent       bool
        package_generated bool
        vsc_notified     bool
        errors           list[str]
    """
    result = {
        'success': False,
        'candidate_id': None,
        'grade': None,
        'analysis': None,
        'email_sent': False,
        'package_generated': False,
        'vsc_notified': False,
        'errors': [],
    }

    # ── Convenience helpers ───────────────────────────────────────────────────
    first_name     = intake_data.get('first_name', '')
    last_name      = intake_data.get('last_name', '')
    candidate_name = f"{first_name} {last_name}".strip()
    candidate_email = intake_data.get('email', '')
    candidate_phone = intake_data.get('phone', '')
    city           = intake_data.get('city', '')
    state          = intake_data.get('state', '')
    candidate_location = f"{city}, {state}".strip(', ')
    resume_text    = intake_data.get('resume_text', '')
    target_job_id  = intake_data.get('target_job_id', '')
    vsc_name       = intake_data.get('vsc_name', 'Your VSC')
    vsc_email      = intake_data.get('vsc_email', '')
    outlook_user   = intake_data.get('outlook_user', '')
    outlook_password = intake_data.get('outlook_password', '')

    # ── Step 1: Get target job ────────────────────────────────────────────────
    target_job = None
    if target_job_id:
        try:
            target_job = get_job_by_id(target_job_id)
            if not target_job:
                result['errors'].append(f"Target job ID '{target_job_id}' not found in database")
        except Exception as e:
            result['errors'].append(f"Step 1 (get target job) failed: {e}")
    else:
        result['errors'].append("No target_job_id provided")

    # ── Step 2: Geocode candidate location ───────────────────────────────────
    cand_lat, cand_lon = None, None
    try:
        if city and state:
            cand_lat, cand_lon = geocode_fn(city, state)
            if cand_lat is None:
                result['errors'].append(f"Could not geocode location: {city}, {state}")
    except Exception as e:
        result['errors'].append(f"Step 2 (geocode) failed: {e}")

    # ── Step 3: Find nearby jobs ──────────────────────────────────────────────
    nearby_jobs = []
    try:
        if cand_lat is not None and cand_lon is not None:
            db_nearby = get_jobs_near(cand_lat, cand_lon, radius_miles=50)
            # Convert DB row format to the format analyze_fn expects
            for j in db_nearby:
                nearby_jobs.append({
                    'Job Title':     j.get('job_title', ''),
                    'Company Name':  j.get('company_name', ''),
                    'City':          j.get('city', ''),
                    'State':         j.get('state', ''),
                    'Job Description': j.get('job_description', ''),
                    'Qualifications':  j.get('qualifications', ''),
                    'Salary From':   j.get('salary_from', ''),
                    'Salary To':     j.get('salary_to', ''),
                    'distance':      round(j.get('distance_miles', 0), 1),
                })
    except Exception as e:
        result['errors'].append(f"Step 3 (nearby jobs) failed: {e}")

    # ── Step 4: AI analysis ───────────────────────────────────────────────────
    analysis = None
    grade = None
    if resume_text and target_job:
        try:
            # Build target_job in the format analyze_fn expects
            target_job_fmt = {
                'Job Title':       target_job.get('job_title', ''),
                'Company Name':    target_job.get('company_name', ''),
                'City':            target_job.get('city', ''),
                'State':           target_job.get('state', ''),
                'Job Description': target_job.get('job_description', ''),
                'Qualifications':  target_job.get('qualifications', ''),
                'Salary From':     target_job.get('salary_from', ''),
                'Salary To':       target_job.get('salary_to', ''),
            }
            analysis = analyze_fn(
                resume_text,
                candidate_name,
                candidate_location,
                target_job_fmt,
                nearby_jobs,
            )
            result['analysis'] = analysis

            # Extract grade from analysis text
            if analysis:
                grade_match = re.search(r'\*\*Grade:\*\*\s*([A-F][+-]?)', analysis)
                if grade_match:
                    grade = grade_match.group(1)
            result['grade'] = grade

        except Exception as e:
            result['errors'].append(f"Step 4 (AI analysis) failed: {e}")
    else:
        if not resume_text:
            result['errors'].append("No resume text provided for analysis")
        if not target_job:
            result['errors'].append("No target job available for analysis")

    # ── Step 5: Create candidate record in CRM ────────────────────────────────
    candidate_id = None
    try:
        target_title   = target_job.get('job_title', '') if target_job else ''
        target_company = target_job.get('company_name', '') if target_job else ''

        notes_parts = []
        if target_title:
            notes_parts.append(f"Target job: {target_title} at {target_company}")
        if grade:
            notes_parts.append(f"AI Grade: {grade}")
        if analysis:
            notes_parts.append("\n--- AI ANALYSIS ---\n" + analysis[:3000])

        candidate_id = create_candidate({
            'vsc_name':    vsc_name,
            'first_name':  first_name,
            'last_name':   last_name,
            'email':       candidate_email,
            'phone':       candidate_phone,
            'city':        city,
            'state':       state,
            'branch':      intake_data.get('branch', ''),
            'rank':        intake_data.get('rank', ''),
            'mos':         intake_data.get('mos', ''),
            'years_served': intake_data.get('years_served', ''),
            'stage':       'CONTACTED',
            'notes':       '\n'.join(notes_parts),
            'resume_text': resume_text,
            'source':      'intake',
        })
        result['candidate_id'] = candidate_id
    except Exception as e:
        result['errors'].append(f"Step 5 (create candidate) failed: {e}")

    # ── Step 6: Send initial contact email to candidate ───────────────────────
    if candidate_email:
        try:
            rendered = render_template('initial_contact', {
                'candidate_name': first_name or candidate_name,
                'vsc_name':       vsc_name,
            })
            _send_smtp_email(
                from_addr=outlook_user,
                from_label='Work for Warriors',
                to_addr=candidate_email,
                to_label=candidate_name,
                reply_to=vsc_email or outlook_user,
                subject=rendered['subject'],
                body=rendered['body'],
                password=outlook_password,
            )
            # Log the email
            log_email(
                vsc_name=vsc_name,
                vsc_email=vsc_email,
                to_name=candidate_name,
                to_email=candidate_email,
                template_key='initial_contact',
                template_label='Initial Contact + Info Request',
                subject=rendered['subject'],
                body=rendered['body'],
                status='sent',
                candidate_id=candidate_id,
            )
            result['email_sent'] = True
        except Exception as e:
            result['errors'].append(f"Step 6 (initial contact email) failed: {e}")
            # Log the failure
            try:
                rendered_fallback = render_template('initial_contact', {
                    'candidate_name': first_name or candidate_name,
                    'vsc_name': vsc_name,
                })
                log_email(
                    vsc_name=vsc_name,
                    vsc_email=vsc_email,
                    to_name=candidate_name,
                    to_email=candidate_email,
                    template_key='initial_contact',
                    template_label='Initial Contact + Info Request',
                    subject=rendered_fallback['subject'],
                    body=rendered_fallback['body'],
                    status='failed',
                    error_message=str(e),
                    candidate_id=candidate_id,
                )
            except Exception:
                pass

    # ── Step 7: Generate CalCareers package if applicable ─────────────────────
    if generate_package_fn is not None and target_job:
        try:
            board = (target_job.get('board') or '').lower()
            company = (target_job.get('company_name') or '').lower()
            is_calcareers = (
                board == 'calcareers'
                or 'california' in company
            )
            if is_calcareers:
                pkg_bytes = generate_package_fn(intake_data)
                if pkg_bytes:
                    result['package_generated'] = True
        except Exception as e:
            result['errors'].append(f"Step 7 (CalCareers package) failed: {e}")

    # ── Step 8: Send VSC notification email ───────────────────────────────────
    notify_addr = vsc_email or outlook_user
    if notify_addr:
        try:
            target_title   = target_job.get('job_title', 'N/A') if target_job else 'N/A'
            target_company = target_job.get('company_name', 'N/A') if target_job else 'N/A'
            grade_display  = grade or 'N/A'
            analysis_snippet = (analysis or '')[:1500]

            subject = f"New Candidate Ready — {candidate_name} | Grade: {grade_display}"
            body = f"""NEW AUTOMATED INTAKE SUBMISSION
Work for Warriors Resume AI — Automated Mode

CANDIDATE INFORMATION
---------------------
Name:         {candidate_name}
Email:        {candidate_email}
Phone:        {candidate_phone}
Location:     {candidate_location}
Branch:       {intake_data.get('branch', 'N/A')}
Rank:         {intake_data.get('rank', 'N/A')}
MOS/Rate:     {intake_data.get('mos', 'N/A')}
Years Served: {intake_data.get('years_served', 'N/A')}

TARGET JOB
----------
Title:   {target_title}
Company: {target_company}

AI ANALYSIS GRADE: {grade_display}

ANALYSIS SUMMARY (first 1500 characters)
-----------------------------------------
{analysis_snippet}

---
The candidate has been added to your CRM (source: intake).
Log in to Work for Warriors Resume AI to view the full record and analysis.
"""

            _send_smtp_email(
                from_addr=outlook_user,
                from_label='Work for Warriors AI',
                to_addr=notify_addr,
                to_label=vsc_name,
                reply_to=outlook_user,
                subject=subject,
                body=body,
                password=outlook_password,
            )
            result['vsc_notified'] = True
        except Exception as e:
            result['errors'].append(f"Step 8 (VSC notification) failed: {e}")

    # ── Final result ──────────────────────────────────────────────────────────
    # Success if candidate was created (core requirement)
    result['success'] = candidate_id is not None
    return result


def _send_smtp_email(
    from_addr: str,
    from_label: str,
    to_addr: str,
    to_label: str,
    reply_to: str,
    subject: str,
    body: str,
    password: str,
):
    """
    Send a plain-text email via Office 365 SMTP.
    Raises on failure.
    """
    if not from_addr or not password:
        raise ValueError("SMTP credentials not configured (OUTLOOK_USER / OUTLOOK_PASSWORD)")

    msg = MIMEMultipart()
    msg['From']     = f"{from_label} <{from_addr}>"
    msg['To']       = f"{to_label} <{to_addr}>"
    msg['Reply-To'] = reply_to
    msg['Subject']  = subject
    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP('smtp.office365.com', 587) as server:
        server.ehlo()
        server.starttls()
        server.login(from_addr, password)
        server.send_message(msg)
