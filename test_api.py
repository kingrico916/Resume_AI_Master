"""Test OpenRouter with realistic prompt size and 4 concurrent threads."""
import requests, os, threading, time
from dotenv import load_dotenv
load_dotenv()

OR_KEY = os.getenv("OPENROUTER_API_KEY", "")
print(f"Key: {'OK (' + OR_KEY[:12] + '...)' if OR_KEY else 'MISSING'}")

FAKE_RESUME = """
John Smith | john@email.com | 555-123-4567 | Sacramento, CA
MILITARY SERVICE: U.S. Army, Sergeant E-5, 2015–2022 (7 years)
MOS: 25U Signal Support Systems Specialist
Deployments: Afghanistan (2017), Kuwait (2019)

EDUCATION:
Bachelor of Science, Information Technology — Sacramento State University, 2014
GPA: 3.4

WORK EXPERIENCE:
IT Support Specialist — Sacramento County, Jan 2023–Present
- Provide Tier 2 technical support for 500+ end users
- Manage Active Directory, Office 365, ServiceNow ticketing
- Reduced average ticket resolution time by 35%

Network Technician — AT&T, Jun 2022–Dec 2022
- Installed and maintained fiber optic and copper infrastructure
- Troubleshot network outages and performance issues

SKILLS: Windows Server, Linux, Python, SQL, Cisco networking, CompTIA Security+

AWARDS: Army Achievement Medal (x3), Good Conduct Medal
""" * 3  # Make it ~600 tokens

FAKE_JOB = """
IT Analyst I — California Department of Technology
Salary: $5,500–$7,500/month
Location: Sacramento, CA

DESCRIPTION:
Under the general supervision of the IT Manager, the IT Analyst I will provide technical support
and analysis for enterprise information systems. Responsibilities include analyzing user requirements,
evaluating IT systems, providing recommendations, and assisting with implementation of technology solutions.

QUALIFICATIONS:
- Bachelor's degree in Computer Science, Information Technology, or related field
- 1 year of IT support or systems analysis experience
- Knowledge of Windows Server environments
- Experience with ITSM tools (ServiceNow preferred)
- Strong written and verbal communication skills
- California driver's license preferred
""" * 2

PROMPT = f"""You are a veteran employment eligibility analyst.

CANDIDATE: John Smith
JOB: IT Analyst I at California Dept of Technology (Sacramento, CA)
JOB DESCRIPTION: {FAKE_JOB}
RESUME:
{FAKE_RESUME}

OUTPUT — EXACT FORMAT ONLY:
ELIGIBILITY: [SUITABLE | PENDING REVIEW | NOT SUITABLE — FINAL]

REQUIREMENTS:
[Requirement] — [CONFIRMED | NOT_PROVIDED | NOT_MET] — [brief note]

PRIMARY_GRADE: [grade]
TRUE_FIT_GRADE: [grade]
SHADOW_GRADE: [grade or N/A]

CLASSIFICATION: [Direct Fit | Redirect | Develop | Non-viable]
END"""

print(f"\nPrompt length: ~{len(PROMPT.split())} words / ~{len(PROMPT)//4} tokens")
print(f"\nTesting 4 concurrent threads (simulating 1 candidate's council)...\n")

results = {}
errors = {}

def call_track(track_id, max_tokens):
    t0 = time.time()
    try:
        resp = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={"Authorization": f"Bearer {OR_KEY}", "Content-Type": "application/json"},
            json={
                "model": "anthropic/claude-4.5-haiku-20251001",
                "messages": [{"role": "user", "content": PROMPT}],
                "max_tokens": max_tokens,
                "provider": {"order": ["Anthropic"], "allow_fallbacks": True},
            },
            timeout=90,
        )
        elapsed = round(time.time() - t0, 1)
        if resp.status_code == 200:
            choices = resp.json().get('choices', [])
            if choices:
                content = choices[0].get('message', {}).get('content', '')
                results[track_id] = f"OK ({len(content)} chars, {elapsed}s)"
            else:
                results[track_id] = f"200 but no choices ({elapsed}s) — {resp.text[:200]}"
        else:
            results[track_id] = f"HTTP {resp.status_code} ({elapsed}s) — {resp.text[:300]}"
    except Exception as e:
        results[track_id] = f"EXCEPTION: {e}"

threads = [
    threading.Thread(target=call_track, args=('T1-eligibility', 800)),
    threading.Thread(target=call_track, args=('T2-development', 3200)),
    threading.Thread(target=call_track, args=('T3-verification', 600)),
    threading.Thread(target=call_track, args=('T4-opportunities', 900)),
]
t_start = time.time()
for t in threads: t.start()
for t in threads: t.join()
total = round(time.time() - t_start, 1)

print(f"Completed in {total}s:\n")
for k, v in sorted(results.items()):
    print(f"  {k}: {v}")
