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
- [Dependencies](docs/DEPENDENCIES.md)
- [Deployment](docs/DEPLOYMENT.md)

## Projects

- **MT** - CosmeticaBar (testers, samples)
- **SK** - Carmado (samples, RRP, discounts)
- **SS** - San (base price)

## License

MIT
# CareAgent
