# Deploy to VPS (Remote Shell Script)

# 1. Upload Code (exclude .git, venv, pycache)
echo "🚀 Syncing files..."
rsync -avz --exclude '.git' --exclude '__pycache__' --exclude '.venv' \
    ./ app@46.226.167.153:~/AgentCare/

# 2. Add Credentials manually
if [ -f "config/credentials.json" ]; then
    echo "🔑 Uploading credentials.json..."
    scp config/credentials.json app@46.226.167.153:~/AgentCare/config/
else
    echo "⚠️  WARNING: config/credentials.json not found locally."
    echo "    Please make sure to upload it manually to ~/AgentCare/config/credentials.json on the server."
fi

# 3. Create .env on server
echo "⚙️  Setting up .env..."
ssh app@46.226.167.153 "cp -n ~/AgentCare/.env.example ~/AgentCare/.env || true"

# 4. Start Docker Compose
echo "🐳 Starting Docker..."
ssh app@46.226.167.153 "cd ~/AgentCare && \
    if docker compose version >/dev/null 2>&1; then \
        docker compose down && docker compose build && docker compose up -d; \
    elif docker-compose version >/dev/null 2>&1; then \
        docker-compose down && docker-compose build && docker-compose up -d; \
    else \
        echo '❌ Error: Neither \"docker compose\" nor \"docker-compose\" found.'; \
        exit 1; \
    fi"
