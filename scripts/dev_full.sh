#!/bin/bash
set -e

# Configuration
PORT=8000
GAS_CLIENT_FILE="gas/00_Client.gs"
PROD_URL="http://46.226.167.153:8000"

echo "🚀 Starting Full Local Development Environment..."

# 1. Clean up existing processes
echo "🧹 Cleaning up old processes..."
pkill -f "uvicorn src.main:app" || true
pkill -f "ngrok http $PORT" || true

# 2. Start Local Server
echo "💻 Starting FastAPI server..."
.venv/bin/uvicorn src.main:app --host 0.0.0.0 --port $PORT --reload > /tmp/uvicorn.log 2>&1 &
SERVER_PID=$!

# 3. Start ngrok Tunnel
echo "🔗 Starting ngrok tunnel..."
ngrok http $PORT --log=stdout > /tmp/ngrok.log 2>&1 &
NGROK_PID=$!

# 4. Wait for Tunnel URL
echo "⏳ Waiting for tunnel URL..."
sleep 5
TUNNEL_URL=$(curl -s http://localhost:4040/api/tunnels | python3 -c "import sys, json; print(json.load(sys.stdin)['tunnels'][0]['public_url'])" 2>/dev/null)

if [ -z "$TUNNEL_URL" ]; then
    echo "❌ Failed to get ngrok URL. Please check if ngrok is installed and authenticated."
    kill $SERVER_PID || true
    kill $NGROK_PID || true
    exit 1
fi

echo "✅ Tunnel URL: $TUNNEL_URL"

# 5. Temporarily update GAS config and push
echo "📝 Updating Google Apps Script projects to use tunnel..."
original_content=$(cat "$GAS_CLIENT_FILE")

# Backup original file
cp "$GAS_CLIENT_FILE" "$GAS_CLIENT_FILE.bak"

# Replace PROD_URL with TUNNEL_URL in the gas file
# Note: Using | as delimiter in sed because URLs contain slashes
sed -i '' "s|const SERVER_URL = '$PROD_URL'|const SERVER_URL = '$TUNNEL_URL'|g" "$GAS_CLIENT_FILE"

# Push to all GAS projects
bash ./scripts/push_all_gas.sh

echo ""
echo "✨ ALL DONE! ✨"
echo "--------------------------------------------------------"
echo "Локальный сервер: http://localhost:$PORT"
echo "Публичный туннель: $TUNNEL_URL"
echo "Google Таблицы теперь подключены к вашему компьютеру!"
echo "--------------------------------------------------------"
echo "Логи сервера: tail -f /tmp/uvicorn.log"
echo "Нажмите Ctrl+C, чтобы остановить всё и вернуть Production URL"

# Cleanup function on exit
cleanup() {
    echo ""
    echo "🛑 Stopping development environment..."
    kill $SERVER_PID || true
    kill $NGROK_PID || true
    
    echo "⏪ Restoring Production URL in GAS..."
    mv "$GAS_CLIENT_FILE.bak" "$GAS_CLIENT_FILE"
    bash ./scripts/push_all_gas.sh
    
    echo "✅ Done. Production restored."
    exit 0
}

trap cleanup SIGINT SIGTERM

# Keep script running and show server logs
tail -f /tmp/uvicorn.log
