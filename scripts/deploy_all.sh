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
echo "📦 [1/3] Syncing project files to server..."

# Clean up local socket if exists (prevents rsync socket errors)
if [ -e ".beads/bd.sock" ]; then
    rm -f ".beads/bd.sock"
fi

# Sync entire project except git, caches, venv and .env (to preserve server-side .env)
rsync -avz --exclude '.git' --exclude '__pycache__' --exclude '.venv' --exclude '.env' --exclude '*.sock' \
    -e "ssh -o StrictHostKeyChecking=no" \
    ./ ${SERVER_USER}@${SERVER_IP}:~/AgentCare/

# Ensure correct permissions on the server (appuser needs to read config files)
ssh -o StrictHostKeyChecking=no ${SERVER_USER}@${SERVER_IP} "chmod -R 755 ~/AgentCare && chown -R root:root ~/AgentCare"

echo "✅ Project files synced"

# ============================================================
# 2. Restart Docker with Rebuild
# ============================================================
echo ""
echo "🐳 [2/3] Rebuilding and updating Docker containers..."

ssh -o StrictHostKeyChecking=no ${SERVER_USER}@${SERVER_IP} "cd ~/AgentCare && docker-compose down && docker-compose up -d --build"

# Wait for server to start before reloading rules
echo "⏳ Waiting for server to start..."
sleep 10

echo "✅ Docker updated (rebuilt)"

# Reload Rules on Server for all projects
echo ""
echo "⚙️  Reloading synchronization rules on server..."
# MT
curl -s -X POST "http://${SERVER_IP}:8000/api/v1/rules/13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ/reload?include_disabled=true" > /dev/null || echo "⚠️  Failed to reload MT rules"
# SS
curl -s -X POST "http://${SERVER_IP}:8000/api/v1/rules/1Bq2Pq0P1SQZfJNBZC3yduYCJmnyc4L4vmbLtvsVUkcg/reload?include_disabled=true" > /dev/null || echo "⚠️  Failed to reload SS rules"
# SK
curl -s -X POST "http://${SERVER_IP}:8000/api/v1/rules/1hSsS9_Iu_MgKWsoE19hAMouQInLGVFaBF6ZFG4Bsm1s/reload?include_disabled=true" > /dev/null || echo "⚠️  Failed to reload SK rules"

echo "✅ Rules reloaded"

# ============================================================
# 4. Deploy Google Apps Script
# ============================================================
echo ""
echo "📜 [4/4] Deploying Google Apps Script..."

if ! command -v clasp &> /dev/null; then
    echo "⚠️  clasp not found. Skipping GAS deploy."
    echo "   Install: npm install -g @google/clasp"
    exit 0
fi

# Deploy all GAS projects (MT, SK, SS) using the push_all_gas.sh script
bash ./scripts/push_all_gas.sh

echo "✅ All GAS projects pushed"

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
echo "🔔 ВНИМАНИЕ: Если вы не видите новые файлы в редакторе Apps Script,"
echo "   просто ОБНОВИТЕ СТРАНИЦУ браузера (F5)."
echo ""
