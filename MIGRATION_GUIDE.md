# AgentCare - MacBook Migration Guide

Complete guide for seamless migration of AgentCare system to a new MacBook device.

## Table of Contents

1. [Pre-Migration Preparation](#pre-migration-preparation)
2. [Migration Day Steps](#migration-day-steps)
3. [Automated Setup](#automated-setup)
4. [Manual Setup (Alternative)](#manual-setup-alternative)
5. [Health Checks & Verification](#health-checks--verification)
6. [Troubleshooting](#troubleshooting)
7. [Backup & Recovery](#backup--recovery)

---

## Pre-Migration Preparation

### On Your Current MacBook (Before Migration)

#### Step 1: Prepare Credentials for Transfer

**Important: NEVER share actual credentials files via unencrypted channels**

Option A: **Using Secure Cloud Storage** (Recommended)
```bash
# On your current device, upload credentials to secure storage:
# 1. Open config/credentials.json (Google Service Account)
# 2. Upload to Google Drive (in a private folder)
# 3. Create shareable link with restricted access
# 4. Note the file ID for easy download later
```

Option B: **Using AirDrop**
```bash
# Prepare credentials on current device:
mkdir -p ~/Desktop/AgentCare_Migration
cp config/credentials.json ~/Desktop/AgentCare_Migration/
cp .env ~/Desktop/AgentCare_Migration/.env.backup

# Keep this folder ready for AirDrop transfer
```

Option C: **Using SSH Key-Protected Repository**
```bash
# If credentials are stored in a private Git repo branch:
git branch credentials-backup
git checkout credentials-backup
# Only push to private repo with SSH authentication
```

#### Step 2: Verify Current Installation

```bash
# Check current system status
cd ~/Desktop/AgentCare
python scripts/health_check.py

# Export sync logs for reference
docker-compose exec app python -c "
import json
from pathlib import Path
logs = list(Path('data/sync_logs').glob('*.json'))[:10]
for log in logs:
    print(log.name)
"

# Document any active processes
docker-compose ps
```

#### Step 3: Create Backup Archives

```bash
# Backup critical data
cd ~/Desktop/AgentCare

# Backup configuration
tar -czf ~/Desktop/agentcare_config_backup.tar.gz config/

# Backup application logs and sync records
tar -czf ~/Desktop/agentcare_data_backup.tar.gz data/ logs/

# Store these securely - you may want to keep them on a USB drive or cloud storage
```

#### Step 4: Document Your Setup

Create a `MIGRATION_NOTES.txt` file:
```
System Migration Notes
======================
Current Date: [date]
Previous Device: [MacBook model, OS version]
New Device: [MacBook model, OS version]

Installed Components:
- Python version: [output of python3 --version]
- Docker version: [output of docker --version]
- Poetry version: [output of poetry --version]

Configuration:
- Project location: ~/Desktop/AgentCare
- Redis URL: [from .env]
- Gemini API Key: [KEY EXISTS - don't share actual value]
- Google Credentials: [FILE EXISTS - credentials.json]

External APIs Status:
- Google Sheets API: [Connected/Not tested]
- Supabase URL: [configured/not configured]
- Telegram Bot: [yes/no]

Critical URLs:
- Production server: [IP/hostname]
- GitHub repo: [URL]
- Documentation: [where to find it]

Notes:
[Any specific setup notes]
```

---

## Migration Day Steps

### Step 1: Initial Setup on New MacBook

When you have your new MacBook ready:

```bash
# 1. Open Terminal
# 2. Create workspace directory
mkdir -p ~/Desktop
cd ~/Desktop

# 3. Clone the repository (assuming you have Git access)
git clone https://github.com/yourusername/AgentCare.git
# OR restore from backup/migration folder
```

### Step 2: Transfer Credentials

Choose ONE of these methods:

**Method A: From Google Drive**
```bash
# Download credentials from shared link
# Place in: ~/Desktop/AgentCare/config/credentials.json
```

**Method B: From AirDrop**
```bash
# Accept AirDrop transfer of AgentCare_Migration folder
# Copy files to ~/Desktop/AgentCare/
cp ~/Desktop/AgentCare_Migration/credentials.json \
   ~/Desktop/AgentCare/config/

cp ~/Desktop/AgentCare_Migration/.env.backup \
   ~/Desktop/AgentCare/.env
```

**Method C: From Migration USB**
```bash
# Insert USB drive
# Copy backup archives
tar -xzf /Volumes/[USB_NAME]/agentcare_config_backup.tar.gz \
    -C ~/Desktop/AgentCare/

tar -xzf /Volumes/[USB_NAME]/agentcare_data_backup.tar.gz \
    -C ~/Desktop/AgentCare/
```

### Step 3: Verify File Transfer

```bash
# Check that credentials are in place
ls -la ~/Desktop/AgentCare/config/credentials.json
ls -la ~/Desktop/AgentCare/.env

# Check file sizes match (should be similar to original)
stat ~/Desktop/AgentCare/config/credentials.json
```

---

## Automated Setup

This is the easiest and recommended approach. The script automates nearly everything.

### Full Automated Setup (Recommended)

```bash
# Navigate to project directory
cd ~/Desktop/AgentCare

# Make setup script executable
chmod +x setup_new_mac.sh

# Run the complete setup
# Option 1: Interactive mode (asks for confirmations)
bash setup_new_mac.sh

# Option 2: Assume yes to all prompts
bash setup_new_mac.sh --yes

# Option 3: Skip Docker (if installing manually)
bash setup_new_mac.sh --skip-docker

# Option 4: Copy credentials from old MacBook (if accessible via network)
bash setup_new_mac.sh --source-dir /Volumes/OldMacBook/Users/yourname/Desktop/AgentCare
```

### What the Setup Script Does

The `setup_new_mac.sh` script automates these phases:

1. **Prerequisites Check**
   - Verifies macOS environment
   - Installs Xcode Command Line Tools if needed
   - Installs/verifies Homebrew
   - Installs Git, Python 3.11+, Poetry

2. **Repository Setup**
   - Clones or updates from Git
   - Checks out latest main branch

3. **Python Dependencies**
   - Creates virtual environment
   - Installs all Poetry dependencies

4. **Credentials & Configuration**
   - Verifies Google credentials
   - Sets up .env file from template
   - Loads project configurations

5. **Docker Setup**
   - Installs Docker Desktop if needed
   - Waits for Docker daemon to start
   - Installs Docker Compose

6. **Services Startup**
   - Starts Redis container
   - Starts FastAPI application
   - Waits for services to become healthy

7. **Health Checks**
   - Runs comprehensive system checks
   - Tests all API connections
   - Verifies database connectivity

8. **Verification**
   - Tests API endpoints
   - Checks Redis connection
   - Verifies container status

### Monitoring Setup Progress

While the setup runs, monitor the log:

```bash
# In another terminal window
tail -f ~/Desktop/AgentCare/setup.log
```

---

## Manual Setup (Alternative)

If you prefer to set up manually or the automated script encounters issues:

### Phase 1: Prerequisites

```bash
# Install Homebrew (if not installed)
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Install core tools
brew install git python@3.11 poetry

# Verify installations
git --version
python3 --version
poetry --version
```

### Phase 2: Project Setup

```bash
cd ~/Desktop/AgentCare

# Create and activate virtual environment
python3 -m venv .venv
source .venv/bin/activate

# Install dependencies
poetry install
```

### Phase 3: Configuration

```bash
# Copy .env file
cp .env.example .env

# Edit .env with your values
nano .env

# Verify credentials
ls -la config/credentials.json
```

### Phase 4: Docker Setup

```bash
# Install Docker Desktop
brew install --cask docker

# Start Docker Desktop (manually or via:)
open /Applications/Docker.app

# Wait for Docker daemon to start
docker info

# Start services
docker-compose up -d
```

### Phase 5: Verification

```bash
# Check services
docker-compose ps

# Run health checks
python scripts/health_check.py

# Test API
curl http://localhost:8000/health
```

---

## Health Checks & Verification

### Comprehensive Health Check

```bash
# Run the health check script
cd ~/Desktop/AgentCare
source .venv/bin/activate
python scripts/health_check.py
```

The script checks:

✓ **Environment & Configuration**
- .env file presence
- settings.yaml configuration
- Project configs (MT, SK, SS)

✓ **Credentials & Security**
- Google Service Account
- API keys configured
- Sensitive files protected

✓ **Google APIs**
- Service account authentication
- Sheets API connectivity
- Drive API access

✓ **Redis Cache**
- Connection status
- Memory usage
- Data integrity

✓ **FastAPI Server**
- Health endpoint responding
- Swagger UI accessible
- Server version info

✓ **Docker Services**
- Docker daemon running
- docker-compose configuration
- Container status

✓ **Supabase Database**
- PostgreSQL connection
- Table accessibility
- Data sync status

✓ **Data Directories**
- Logs directory
- Sync records
- Storage capacity

✓ **Python Dependencies**
- Critical packages installed
- Version compatibility

### Manual Verification Tests

```bash
# 1. Test API
curl -X GET http://localhost:8000/health | python -m json.tool

# 2. Test Redis
docker exec agentcare-redis redis-cli ping

# 3. Check Google Credentials
python -c "
from google.oauth2.service_account import Credentials
creds = Credentials.from_service_account_file('config/credentials.json')
print(f'✓ Credentials loaded: {creds.service_account_email}')
"

# 4. View running containers
docker-compose ps

# 5. Check API logs
docker-compose logs agentcare | tail -20

# 6. Test database connection (if Supabase configured)
python -c "
from supabase import create_client
import os
from dotenv import load_dotenv
load_dotenv()
supabase = create_client(
    os.getenv('SUPABASE_URL'),
    os.getenv('SUPABASE_SERVICE_ROLE_KEY')
)
print('✓ Supabase connected')
"
```

---

## Troubleshooting

### Docker Issues

**Problem: "Docker daemon is not running"**
```bash
# Solution 1: Start Docker manually
open /Applications/Docker.app

# Solution 2: Check Docker installation
docker info

# Solution 3: Reinstall Docker
brew uninstall --cask docker
brew install --cask docker
```

**Problem: "Cannot connect to Docker daemon"**
```bash
# Solution: Wait for Docker to fully start
sleep 10
docker-compose up -d --build
```

### Python/Dependencies Issues

**Problem: "Module not found" errors**
```bash
# Solution: Activate virtual environment
source .venv/bin/activate

# Install missing dependencies
poetry install

# Or update all
poetry update
```

**Problem: "Poetry not found"**
```bash
# Reinstall Poetry
curl -sSL https://install.python-poetry.org | python3 -

# Add to PATH
export PATH="$HOME/.local/bin:$PATH"
poetry --version
```

### Credentials Issues

**Problem: "credentials.json not found"**
```bash
# Solution: Verify file exists
ls -la config/credentials.json

# If missing: Copy from backup
cp ~/Desktop/AgentCare_Migration/credentials.json config/

# If still missing: Download from secure storage
# (Google Drive, USB, etc.)
```

**Problem: "Invalid credentials" errors**
```bash
# Check credentials format
python -c "import json; json.load(open('config/credentials.json'))"

# Verify file permissions (should be readable)
chmod 600 config/credentials.json

# Test credentials directly
python -c "
from google.oauth2.service_account import Credentials
try:
    Credentials.from_service_account_file('config/credentials.json')
    print('✓ Credentials valid')
except Exception as e:
    print(f'✗ Error: {e}')
"
```

### API/Redis Connection Issues

**Problem: "Cannot connect to http://localhost:8000"**
```bash
# Check if service is running
docker-compose ps

# View logs
docker-compose logs agentcare

# Restart services
docker-compose restart agentcare

# Check port is available
lsof -i :8000
```

**Problem: "Cannot connect to Redis"**
```bash
# Check Redis container
docker-compose ps agentcare-redis

# Test Redis
docker exec agentcare-redis redis-cli ping

# View Redis logs
docker-compose logs agentcare-redis

# Restart Redis
docker-compose restart agentcare-redis
```

### Google Sheets API Issues

**Problem: "Permission denied" errors**
```bash
# Verify credentials have correct scopes
python -c "
import json
with open('config/credentials.json') as f:
    data = json.load(f)
    print(f\"Service Account: {data.get('client_email')}\")
    print(f\"Project: {data.get('project_id')}\")
"

# Test API connection
python -c "
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

creds = Credentials.from_service_account_file(
    'config/credentials.json',
    scopes=['https://www.googleapis.com/auth/spreadsheets']
)

service = build('sheets', 'v4', credentials=creds)
print('✓ API connection successful')
"
```

### Supabase Issues

**Problem: "Cannot connect to Supabase"**
```bash
# Check configuration
cat .env | grep SUPABASE

# Test connection
python -c "
import os
from dotenv import load_dotenv
from supabase import create_client

load_dotenv()
url = os.getenv('SUPABASE_URL')
key = os.getenv('SUPABASE_SERVICE_ROLE_KEY')

if not url or not key:
    print('✗ Supabase credentials not configured')
else:
    client = create_client(url, key)
    print('✓ Connected')
"
```

### Port Conflicts

**Problem: "Port 8000 already in use"**
```bash
# Find process using port
lsof -i :8000

# Kill process (replace PID)
kill -9 <PID>

# Or use different port in .env
# SERVER_PORT=8001
# Then restart
docker-compose down
docker-compose up -d
```

---

## Backup & Recovery

### Creating Backups

```bash
# Full system backup
tar -czf ~/backup_agentcare_full_$(date +%Y%m%d).tar.gz \
    ~/Desktop/AgentCare

# Configuration-only backup
tar -czf ~/backup_agentcare_config_$(date +%Y%m%d).tar.gz \
    ~/Desktop/AgentCare/config \
    ~/Desktop/AgentCare/.env

# Data backup (logs and sync records)
tar -czf ~/backup_agentcare_data_$(date +%Y%m%d).tar.gz \
    ~/Desktop/AgentCare/data \
    ~/Desktop/AgentCare/logs
```

### Restoring from Backup

```bash
# Restore full system
cd ~/Desktop
tar -xzf ~/backup_agentcare_full_20260114.tar.gz

# Restore configuration only
tar -xzf ~/backup_agentcare_config_20260114.tar.gz \
    -C ~/Desktop/AgentCare/

# Restore data only
tar -xzf ~/backup_agentcare_data_20260114.tar.gz \
    -C ~/Desktop/AgentCare/
```

### Recovery Scenarios

**Scenario 1: Corrupted Installation**
```bash
# Start fresh
rm -rf ~/Desktop/AgentCare
git clone <repo-url> ~/Desktop/AgentCare
cd ~/Desktop/AgentCare

# Restore configuration
tar -xzf ~/backup_agentcare_config_latest.tar.gz -C .

# Run setup
bash setup_new_mac.sh
```

**Scenario 2: Failed Migration (Keep Old System)**
```bash
# If migration fails, you can go back to old MacBook
# Data is still there, just start services as before

# On old device
cd ~/Desktop/AgentCare
docker-compose up -d
```

---

## Quick Reference

### Essential Commands

```bash
# Navigate to project
cd ~/Desktop/AgentCare

# Activate Python environment
source .venv/bin/activate

# Start services
bash scripts/startup.sh

# Run health check
python scripts/health_check.py

# View logs
docker-compose logs -f

# View specific service logs
docker-compose logs -f agentcare

# Stop services
docker-compose down

# Restart services
docker-compose restart

# View API documentation
# Open: http://localhost:8000/docs

# View system health
# Open: http://localhost:8000/health

# View Swagger UI
curl http://localhost:8000/openapi.json | jq .
```

### Environment Variables

Key variables in `.env`:

```bash
# Server
SERVER_HOST=0.0.0.0
SERVER_PORT=8000
SERVER_WORKERS=4
DEBUG=false

# Cache
REDIS_URL=redis://localhost:6379/0
CACHE_TTL=3600

# APIs
GEMINI_API_KEY=<your-key>
GEMINI_MODEL=gemini-2.5-flash

# Google
GOOGLE_CREDENTIALS_FILE=config/credentials.json

# Supabase (optional)
SUPABASE_URL=https://xxxx.supabase.co
SUPABASE_SERVICE_ROLE_KEY=<your-key>

# Webhook
WEBHOOK_SECRET=<random-secret>
```

---

## Support & Additional Resources

- **GitHub Issues**: Report problems at https://github.com/yourusername/AgentCare/issues
- **Documentation**: Check `docs/` directory
- **Logs**: All logs are saved in `logs/` and `data/sync_logs/`
- **Health Check**: Run `python scripts/health_check.py` to diagnose issues

---

## Checklist for Successful Migration

- [ ] Pre-migration backup created
- [ ] Credentials securely transferred
- [ ] Repository cloned on new device
- [ ] `.env` file configured
- [ ] `config/credentials.json` in place
- [ ] `setup_new_mac.sh` completed successfully
- [ ] Health checks all passing
- [ ] API responding on localhost:8000
- [ ] Redis connected
- [ ] Google Sheets API authenticated
- [ ] Docker containers running
- [ ] All previous data and configurations accessible
- [ ] Services auto-start on system restart (optional - can configure)

---

**Last Updated**: 2026-01-14
**Version**: 1.0
**For Latest Updates**: Check the main repository
