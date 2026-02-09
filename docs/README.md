# AgentCare

Google Sheets automation with AI - Python replacement for MyGoogleScripts.

## 🎯 Main Goals
1. **Stable Buttons**: Reliable execution from Google Sheets.
2. **"Journal" Workflow**: Debug via logs (Supabase/Files) before fixing code.
3. **Traceable Sync**: Data consistency via Supabase.
4. **Smart Agent**: Proactive AI assistance.

## 🤖 For AI Agents
**Start Here:** 👉 **[AI HANDBOOK (Правила)](./AI_HANDBOOK.md)** 👈
*This is the single source of truth for all rules, workflows (Beads), and architectural principles.*

## 📚 Documentation
- **[Architecture](ARCHITECTURE.md)**: System overview, components, and data flow.
- **[API Documentation](docs/API.md)**: REST API reference.
- **[Deployment](DEPLOY.md)**: Production deployment guide.
- **[Legacy Docs](docs/archive/)**: Archived documentation.

## 🚀 Quick Start

```bash
# Clone
git clone https://github.com/your-repo/agentcare.git
cd agentcare

# Setup
cp .env.example .env
poetry install

# Run Server
poetry run uvicorn src.main:app --reload
```

## 🏗 System Components
- **Server**: Python 3.12 (FastAPI)
- **Database**: Supabase (PostgreSQL)
- **Client**: Google Apps Script (Thin Client)
- **AI**: Gemini 1.5 Pro

## 📦 Projects
- **MT**: CosmeticaBar
- **SK**: Carmado
- **SS**: San

---
*License: MIT*
