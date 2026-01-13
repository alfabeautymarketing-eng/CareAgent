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

# Sync entire project except git, caches, venv and .env (to preserve server-side .env)
# Set permissions explicitly to allow container to read files
rsync -avz --exclude '.git' --exclude '__pycache__' --exclude '.venv' --exclude '.env' \
    --chmod=Du=rwx,Dg=rx,Do=rx,Fu=rw,Fg=r,Fo=r \
    -e "ssh -o StrictHostKeyChecking=no" \
    ./ ${SERVER_USER}@${SERVER_IP}:~/AgentCare/

echo "✅ Project files synced"

# ============================================================
# 2. Restart Docker with Rebuild
# ============================================================
echo ""
echo "🐳 [2/3] Rebuilding and updating Docker containers..."

ssh -o StrictHostKeyChecking=no ${SERVER_USER}@${SERVER_IP} "cd ~/AgentCare && docker-compose down && docker-compose up -d --build"

echo "✅ Docker updated (rebuilt)"

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
