# AgentCare

Google Sheets automation with AI - Python replacement for MyGoogleScripts.

## Features

- Synchronization between 15+ Google Sheets
- AI analysis (INCI, ТН ВЭД classification) via Gemini
- Price processing (9 phases)
- Invoice generation
- Auto-order calculation
- Cascade rules engine
- Transaction support with rollback
- Live server logging to sheet `Логи` (всегда первая вкладка)
- Server-side sync rules support (`config/rules/<spreadsheet_id>.yaml`) to avoid лишних чтений листа правил
- HTML-форма управления правилами: `docs/rules-form.html` (работает через API `/rules/{spreadsheet_id}`)

## Documentation

### For Developers
- [Quick Start Guide](QUICKSTART.md)
- [Production Deployment](PROD_DEPLOY.md)
- [Logging System](LOGGING_README.md)
- [API Documentation](docs/)

### For AI Agents 🤖
- **[AI Central Command](AI_CENTRAL_COMMAND.md)** - START HERE!
- [AI Rules (Detailed)](AI_RULES.md)
- [Agent Procedures](AGENTS.md)
- [Beads Quickstart](BEADS_QUICKSTART.md)

### System Documentation 📚
- **[Architecture](ARCHITECTURE.md)** - Полная архитектура системы
- **[Function Map](FUNCTION_MAP.md)** - Карта всех функций с зависимостями
- **[API Reference](API_REFERENCE.md)** - Справочник по всем endpoints

## Quick Start

```bash
# Clone
git clone https://github.com/your-repo/agentcare.git
cd agentcare

# Setup
cp .env.example .env
# Edit .env with your credentials

# Install dependencies
pip install poetry
poetry install

# Run
poetry run uvicorn src.main:app --reload
```

## Docker

```bash
docker compose up -d
```

## Documentation

- [Architecture](docs/ARCHITECTURE.md)
- [Modules](docs/MODULES.md)
- [API](docs/API.md)
- [Rules](docs/RULES.md)
- [Dependencies](docs/DEPENDENCIES.md)
- [Deployment](docs/DEPLOYMENT.md)
- [GAS & CI: clasp credentials and GitHub secrets](docs/GAS_CI_SECRETS.md)

## Projects

- **MT** - CosmeticaBar (testers, samples)
- **SK** - Carmado (samples, RRP, discounts)
- **SS** - San (base price)

## License

MIT

# CareAgent
