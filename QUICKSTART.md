# QUICK START - Resume AI + CalCareers Integrated System

## ONE WEB PORTAL FOR EVERYTHING

**Includes:**
- Resume AI analysis
- Job matching
- CalCareers package generation
- All in one interface

---

## Windows Setup (5 Minutes)

### Prerequisites
- Docker Desktop for Windows
- OpenRouter API key

---

### Step 1: Install Docker Desktop

Download: https://www.docker.com/products/docker-desktop/

Install and start Docker Desktop.

---

### Step 2: Extract Files

Unzip `resume_ai_master.zip` to:
```
C:\resume_ai_master\
```

---

### Step 3: Configure Environment

1. Copy `.env.example` to `.env`
2. Edit `.env` in Notepad:

```
OPENROUTER_API_KEY=sk-or-v1-YOUR_KEY_HERE
OUTLOOK_USER=anthony.antonucci@workforwarriors.org
OUTLOOK_PASSWORD=your_password_here
```

Save and close.

---

### Step 4: Start Container

Open Command Prompt:

```cmd
cd C:\resume_ai_master
docker-compose up -d
```

Wait 30 seconds for startup.

---

### Step 5: Open Web Interface

Browser: http://localhost:5000

---

## VSC Complete Workflow

### 1. Load Jobs
- Upload master SmartJobBoard CSV
- Wait for geocoding (shows status)

### 2. Enter Candidate
- Name, City, State
- Click "Search Jobs by Location"

### 3. Select Job
- Pick target job from list
- Shows distance in miles

### 4. Paste Resume
- Copy/paste resume text

### 5. Analyze
- Click "Analyze Resume"
- Wait 60 seconds
- Review AI analysis

### 6. Generate CalCareers Package (NEW!)
- Form appears after analysis
- Fill candidate details
- Add education entries
- Add work experience
- Enter job target info
- Select Veterans' Preference (if applicable)
- Choose template track
- Click "Generate CalCareers Package"
- **Downloads ZIP file** with complete application folders

---

## What You Get

**After Step 5 (Analysis):**
- Grade + justification
- Missing requirements
- Resume rewrite
- 3 alternative jobs

**After Step 6 (Package):**
- ZIP download with:
  - 7 organized folders
  - VSC checklist
  - CalHR 1093 instructions
  - Audit log

---

## Stopping/Starting

**Stop:**
```cmd
docker-compose down
```

**Start:**
```cmd
docker-compose up -d
```

**View Logs:**
```cmd
docker-compose logs -f
```

---

## Common Issues

**"API Key not set"**
- Check .env file has correct OPENROUTER_API_KEY

**"No jobs loaded"**
- Upload CSV first
- Check CSV has City and State columns

**"Docker not running"**
- Start Docker Desktop
- Wait for it to fully start

---

## Files Overview

```
resume_ai_master/
├── app.py                  # Flask server
├── templates/
│   └── index.html          # Web interface
├── Dockerfile              # Docker image
├── docker-compose.yml      # Docker config
├── requirements.txt        # Python dependencies
├── .env.example            # Environment template
└── README.md               # Full documentation
```

---

## Need Help?

1. Check logs: `docker-compose logs`
2. Test health: http://localhost:5000/health
3. Read full README.md

---

**Quick Start Version:** 1.0.0  
**Date:** 2026-01-27
