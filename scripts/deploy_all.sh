#!/bin/bash
set -e

# ============================================================
# deploy_all.sh — Полный деплой: сервер + GAS
# ============================================================

# Configuration
SERVER_IP="46.226.167.153"
SERVER_USER="root"
GAS_DIR="./gas"

echo "🚀 ======================================"
echo "   AgentCare Full Deploy"
echo "========================================"

# ============================================================
# 1. Deploy Python Server
# ============================================================
echo ""
echo "📦 [1/3] Syncing Python code to server..."

rsync -avz --exclude '.git' --exclude '__pycache__' --exclude '.venv' --exclude '.env' \
    -e "ssh -o StrictHostKeyChecking=no" \
    ./src/ ${SERVER_USER}@${SERVER_IP}:~/AgentCare/src/

echo "✅ Python code synced"

# ============================================================
# 2. Restart Docker
# ============================================================
echo ""
echo "🐳 [2/3] Updating Docker containers..."

ssh -o StrictHostKeyChecking=no ${SERVER_USER}@${SERVER_IP} "cd ~/AgentCare && docker-compose down && docker-compose up -d"

echo "✅ Docker updated"

# ============================================================
# 3. Deploy Google Apps Script
# ============================================================
echo ""
echo "📜 [3/3] Deploying Google Apps Script..."

if ! command -v clasp &> /dev/null; then
    echo "⚠️  clasp not found. Skipping GAS deploy."
    echo "   Install: npm install -g @google/clasp"
    exit 0
fi

# Check for .clasp.json in gas folder
if [ ! -f "$GAS_DIR/.clasp.json" ]; then
    echo "⚠️  $GAS_DIR/.clasp.json not found. Creating..."
    echo '{
  "scriptId": "199Np7xsBiBRQih5_tlUdpt6EmkfRGjZAhTvKm4Ua0Q6XEaMtvAmQUn0g",
  "rootDir": "."
}' > "$GAS_DIR/.clasp.json"
fi

cd "$GAS_DIR"
clasp push

echo "✅ GAS code pushed"

# ============================================================
# Done
# ============================================================
echo ""
echo "🎉 ======================================"
echo "   Deploy Complete!"
echo "========================================"
echo ""
echo "Server: http://${SERVER_IP}:8000"
echo "Health: http://${SERVER_IP}:8000/health"
echo ""
