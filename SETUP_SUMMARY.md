# AgentCare - MacBook Setup System - Summary

Complete automation system for seamless migration to a new MacBook device. Everything you need is ready to go!

## 🚀 Quick Start (5 Minutes)

### Fastest Way to Set Up on New MacBook

```bash
# 1. Clone repository on new MacBook
git clone <your-repo-url> ~/Desktop/AgentCare
cd ~/Desktop/AgentCare

# 2. Transfer credentials (choose one method):
# Option A: From Google Drive
# Option B: From AirDrop
# Option C: From secure backup
# → Place credentials.json in config/ and .env in root

# 3. Run the setup (fully automated!)
bash setup_new_mac.sh

# That's it! Services will start automatically.
```

---

## 📦 What's Included

### 🔧 Automation Scripts

| Script | Purpose | When to Use |
|--------|---------|------------|
| **`setup_new_mac.sh`** | Complete system setup (9 phases) | First time on new MacBook |
| **`scripts/startup.sh`** | Start all services with checks | Daily server startup |
| **`scripts/health_check.py`** | Comprehensive health verification | After setup or troubleshooting |
| **`.env.template`** | Configuration template | Create your `.env` file |

### 📋 Documentation

| Document | Purpose |
|----------|---------|
| **`MIGRATION_GUIDE.md`** | Complete step-by-step migration guide (40+ sections) |
| **`SETUP_SUMMARY.md`** | This file - quick overview |
| **`docs/`** | Existing project documentation |

---

## 🎯 The Setup Process (Automated)

The `setup_new_mac.sh` script handles these 9 phases automatically:

### Phase 1: Prerequisites ✓
- Checks macOS environment
- Installs Homebrew, Xcode tools, Git
- Installs Python 3.11+ and Poetry

### Phase 2: Repository ✓
- Clones or updates from GitHub
- Ensures latest code

### Phase 3: Python Dependencies ✓
- Creates virtual environment
- Installs all Poetry packages (FastAPI, Redis, Google APIs, etc.)

### Phase 4: Credentials ✓
- Verifies credentials.json
- Sets up .env from template
- Loads project configurations

### Phase 5: Docker ✓
- Installs Docker Desktop if needed
- Waits for daemon to start

### Phase 6: Services ✓
- Starts FastAPI (port 8000)
- Starts Redis cache (port 6379)

### Phase 7: Health Checks ✓
- Tests all API connections
- Verifies databases
- Checks Google Sheets API
- Validates Supabase connection

### Phase 8: Verification ✓
- Tests API endpoints
- Confirms services running
- Checks Docker containers

### Phase 9: Summary ✓
- Shows access URLs
- Provides next steps

---

## ⚙️ Services Being Automated

### Servers Launched
✓ **FastAPI** - Main API server (localhost:8000)
✓ **Redis** - Caching layer (localhost:6379)
✓ **Google Sheets Sync** - Automatic sync service
✓ **Supabase PostgreSQL** - Database (cloud)

### Connections Verified
✓ Google Sheets API
✓ Google Drive API
✓ Gemini AI API
✓ Supabase PostgreSQL
✓ Redis Cache
✓ Local API endpoints

### Settings Transferred
✓ Google credentials
✓ API keys
✓ Project configurations (MT, SK, SS)
✓ Sync rules
✓ Agent settings
✓ Environment variables

---

## 🔑 Required Credentials

### Must Have (Critical)
```
config/credentials.json   → Google Service Account key
.env file with:
  - GEMINI_API_KEY        → Gemini AI API key
  - REDIS_URL             → Redis connection string
```

### Optional (For Full Features)
```
.env file with:
  - SUPABASE_URL          → Database URL
  - SUPABASE_SERVICE_ROLE_KEY → Database auth
  - TELEGRAM_BOT_TOKEN    → For notifications
```

See `.env.template` for all options.

---

## 📖 Usage Scenarios

### Scenario 1: Complete Fresh Start (Recommended)

```bash
cd ~/Desktop/AgentCare

# Run full setup
bash setup_new_mac.sh

# That's it! Everything runs automatically.
```

**Timeline**: ~15-20 minutes
**Automation Level**: ⭐⭐⭐⭐⭐ (Fully automated)

### Scenario 2: Manual Setup Step by Step

```bash
cd ~/Desktop/AgentCare

# Follow detailed steps in MIGRATION_GUIDE.md
# Each phase can be done separately

# Phase 1: Install prerequisites
# Phase 2: Clone repo
# Phase 3: Install Python deps
# ... etc
```

**Timeline**: ~30-45 minutes
**Automation Level**: ⭐ (Manual control)

### Scenario 3: Skip Docker (For Testing)

```bash
# Run setup without Docker
bash setup_new_mac.sh --skip-docker

# Start services manually later
docker-compose up -d
```

**Use When**: Testing before full Docker install

### Scenario 4: Transfer from Old MacBook

```bash
# If old MacBook accessible on network
bash setup_new_mac.sh --source-dir /Volumes/OldMacBook/Users/username/Desktop/AgentCare

# Automatically copies credentials and configs!
```

**Use When**: Old device is accessible via network

---

## 🧪 Health Checks

### Check System Health Anytime

```bash
# Full health report
python scripts/health_check.py

# Shows:
# ✓ Environment & configuration files
# ✓ Credentials & API keys
# ✓ Google API authentication
# ✓ Redis connection
# ✓ API server status
# ✓ Docker services
# ✓ Supabase database
# ✓ Data directories
# ✓ Python dependencies
```

### Quick Manual Checks

```bash
# API responding?
curl http://localhost:8000/health

# Redis connected?
docker exec agentcare-redis redis-cli ping

# Containers running?
docker-compose ps

# View logs?
docker-compose logs -f
```

---

## 🛠️ Troubleshooting Quick Links

| Problem | Solution |
|---------|----------|
| Docker won't start | See "Docker Issues" in MIGRATION_GUIDE.md |
| "Module not found" | Activate venv: `source .venv/bin/activate` |
| Credentials missing | Transfer from old device or secure storage |
| API not responding | Check logs: `docker-compose logs agentcare` |
| Redis connection failed | Restart Redis: `docker-compose restart agentcare-redis` |
| Port already in use | Change port in .env: `SERVER_PORT=8001` |

**Full troubleshooting**: See MIGRATION_GUIDE.md (Troubleshooting section)

---

## 🔄 Daily Startup

After initial setup, starting services is simple:

```bash
cd ~/Desktop/AgentCare

# Option 1: Use startup script (recommended)
bash scripts/startup.sh

# Option 2: Start manually
docker-compose up -d
python scripts/health_check.py

# Option 3: Start with logs visible
docker-compose up
```

---

## 📊 What Gets Set Up

### Code
- Latest code from Git repository
- All project files and configurations
- Google Apps Script deployments (27 files)
- Sync rules and automation logic

### Services
- FastAPI server (Python) - REST API
- Redis cache - Session/data caching
- Google Sheets connectors - Data sync
- Supabase client - PostgreSQL database

### Configurations
- Project settings (MT, SK, SS)
- Sync rules for automation
- Google credentials for APIs
- Environment variables
- Cache settings
- Logging configuration

### Data
- Application logs
- Sync operation history
- Function execution records
- Cached metadata

---

## 🔐 Security Notes

### Safe Credentials Transfer
✓ Use Google Drive with restricted sharing
✓ Use AirDrop between personal devices
✓ Use encrypted USB for backups
✗ Never email credentials
✗ Never share via Slack/Teams
✗ Never commit to Git

### File Permissions
```bash
# Keep credentials.json restricted
chmod 600 config/credentials.json

# Keep .env restricted
chmod 600 .env

# Add to .gitignore (already done)
cat .gitignore | grep -E "^\.env|^config/credentials"
```

### During Migration
1. Credentials only on secure channels
2. Use HTTPS for file transfers
3. Delete old credentials from new device after migration
4. Backup credentials in secure location

---

## 📞 Support & Debugging

### Check Setup Log
```bash
# View setup log
tail -f ~/Desktop/AgentCare/setup.log

# Or open in editor
nano ~/Desktop/AgentCare/setup.log
```

### Enable Verbose Output
```bash
# Run with bash verbose mode
bash -x setup_new_mac.sh

# Runs each command before executing it
```

### View Service Logs
```bash
# All services
docker-compose logs -f

# Specific service
docker-compose logs -f agentcare

# Last N lines
docker-compose logs agentcare --tail 50
```

### Get Help
- **This file**: SETUP_SUMMARY.md (overview)
- **Detailed guide**: MIGRATION_GUIDE.md (40+ sections)
- **Issues**: Check project GitHub issues
- **Logs**: Check setup.log and Docker logs

---

## ✅ Success Checklist

After setup completes, verify:

- [ ] No errors in setup.log
- [ ] Health checks show all "✓ OK"
- [ ] API responding: http://localhost:8000/docs
- [ ] Redis connected: `docker exec agentcare-redis redis-cli ping`
- [ ] All containers running: `docker-compose ps`
- [ ] credentials.json exists: `ls config/credentials.json`
- [ ] .env configured: `cat .env | head -20`
- [ ] Can access Swagger UI: http://localhost:8000/docs
- [ ] Health endpoint responds: `curl http://localhost:8000/health`

---

## 🎓 Learning Resources

### File Locations
- **Setup scripts**: `scripts/` directory
- **Configuration**: `config/` directory
- **Source code**: `src/` directory
- **Documentation**: `docs/` directory
- **Tests**: `tests/` directory
- **Deployment**: `.github/workflows/` directory

### Key Commands Reference
```bash
# Setup
bash setup_new_mac.sh              # Full automation
bash setup_new_mac.sh --yes        # Auto-confirm
bash setup_new_mac.sh --skip-docker # Skip Docker phase

# Services
bash scripts/startup.sh            # Start all services
docker-compose down                # Stop services
docker-compose restart agentcare   # Restart API

# Health & Monitoring
python scripts/health_check.py     # Full health report
docker-compose ps                  # Show container status
docker-compose logs -f             # View live logs

# Development
source .venv/bin/activate          # Activate Python env
poetry install                     # Install dependencies
poetry update                      # Update dependencies
python -m pytest tests/            # Run tests
```

---

## 🚦 Startup Phases Detailed

### When Running `bash setup_new_mac.sh`:

```
╔════════════════════════════════════════╗
║  AgentCare - New MacBook Setup        ║
╚════════════════════════════════════════╝

ℹ Phase 1: Checking prerequisites...
✓ Running on macOS
✓ Homebrew is installed
✓ Git is installed: git version 2.x.x
✓ Python is installed: 3.11.x
✓ Poetry is installed: Poetry 1.x.x
✓ All prerequisites are met

ℹ Phase 2: Repository setup...
✓ Repository cloned to ~/Desktop/AgentCare
✓ Current branch: main
✓ Latest commit: abc1234 - Recent commit message

ℹ Phase 3: Installing Python dependencies...
✓ Virtual environment created
✓ All Python dependencies installed

ℹ Phase 4: Setting up credentials...
✓ Google credentials found
✓ .env file loaded
✓ Credentials and configuration ready

ℹ Phase 5: Setting up Docker...
✓ Docker is running
✓ Docker Compose is installed
✓ Services started

ℹ Phase 6: Waiting for services...
✓ Redis is ready
✓ FastAPI Server is ready

ℹ Phase 7: Running health checks...
✓ All health checks passed

ℹ Phase 8: Service information...
✓ API is responding
✓ Redis is running
✓ Docker containers running

ℹ Phase 9: Startup complete!

========================================
✓ All systems operational!
========================================

Next steps:
1. Dashboard: http://localhost:8000/docs
2. Health check: http://localhost:8000/health
3. View logs: tail -f ~/Desktop/AgentCare/logs/app.log
```

---

## 🎯 Next Steps After Setup

1. **Verify Everything Works**
   ```bash
   python scripts/health_check.py
   ```

2. **Access API Dashboard**
   - Open browser: http://localhost:8000/docs

3. **Check Service Status**
   ```bash
   docker-compose ps
   ```

4. **View Application Logs**
   ```bash
   tail -f logs/app.log
   ```

5. **Run Any Final Configuration**
   - Edit `.env` if additional settings needed
   - Update project configs if project settings changed

6. **Create Backup** (optional but recommended)
   ```bash
   tar -czf ~/backup_agentcare_$(date +%Y%m%d).tar.gz \
       ~/Desktop/AgentCare
   ```

---

## 📈 Performance Notes

**Typical Setup Time**: 15-25 minutes
**Breakdown**:
- Prerequisites: 2-5 min (Homebrew, Git, Python)
- Repository: 1-2 min
- Python deps: 3-5 min
- Docker: 5-10 min
- Services startup: 2-3 min
- Health checks: 1-2 min

**System Requirements**:
- macOS 11+ (Big Sur or later)
- 10 GB free disk space
- 4 GB RAM minimum (8+ recommended)
- Stable internet connection

---

## 🔄 Auto-Restart & Monitoring (Optional)

### Enable Service Auto-Start on Reboot

```bash
# Create launchd plist for auto-start
cat > ~/Library/LaunchAgents/com.agentcare.startup.plist <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.agentcare.startup</string>
  <key>ProgramArguments</key>
  <array>
    <string>bash</string>
    <string>/Users/username/Desktop/AgentCare/scripts/startup.sh</string>
  </array>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <false/>
</dict>
</plist>
EOF

# Load the startup agent
launchctl load ~/Library/LaunchAgents/com.agentcare.startup.plist
```

---

## 📝 Files Created/Modified

### New Files Created
```
✓ setup_new_mac.sh           (Main automation script)
✓ scripts/health_check.py    (Comprehensive health checker)
✓ scripts/startup.sh         (Service startup automation)
✓ .env.template              (Environment configuration template)
✓ MIGRATION_GUIDE.md         (Detailed migration guide)
✓ SETUP_SUMMARY.md          (This file)
```

### Using Existing Files
```
✓ docker-compose.yml         (Service orchestration)
✓ Dockerfile                 (Container definition)
✓ pyproject.toml             (Python dependencies)
✓ .gitignore                 (.env, credentials excluded)
✓ config/                    (Project configurations)
```

---

## 🎉 You're All Set!

Everything is configured for a **seamless transition** to your new MacBook:

1. ✅ **Fully Automated** - 9-phase setup runs completely automatically
2. ✅ **Comprehensive** - Handles all dependencies and services
3. ✅ **Verified** - Built-in health checks ensure everything works
4. ✅ **Documented** - Detailed guide for any custom needs
5. ✅ **Secure** - Credentials handled safely

**Ready?** Run: `bash setup_new_mac.sh`

**Questions?** Check `MIGRATION_GUIDE.md` for 40+ detailed sections.

**Issues?** Run `python scripts/health_check.py` to diagnose.

---

**Last Updated**: January 14, 2026
**Version**: 1.0
**Status**: ✓ Production Ready
