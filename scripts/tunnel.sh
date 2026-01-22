#!/bin/bash

# Port to tunnel
PORT=8000

echo "🚀 Starting ngrok tunnel on port $PORT..."

# Start ngrok in background
# We use --log=stdout to capture the URL later if needed, 
# but for now we just start it.
ngrok http $PORT --log=stdout > /tmp/ngrok.log &
NGROK_PID=$!

# Wait for ngrok to start
sleep 3

# Get the public URL using ngrok API
TUNNEL_URL=$(curl -s http://localhost:4040/api/tunnels | python3 -c "import sys, json; print(json.load(sys.stdin)['tunnels'][0]['public_url'])" 2>/dev/null)

if [ -z "$TUNNEL_URL" ]; then
    echo "❌ Failed to get ngrok URL. Is ngrok running?"
    kill $NGROK_PID
    exit 1
fi

echo ""
echo "✅ Tunnel started!"
echo "Public URL: $TUNNEL_URL"
echo ""
echo "--------------------------------------------------------"
echo "Чтобы переключить Google Таблицу на локальный сервер:"
echo "1. В таблице откройте меню: Настройки -> Установить SERVER_URL"
echo "2. Введите этот URL: $TUNNEL_URL"
echo "--------------------------------------------------------"
echo ""
echo "Press Ctrl+C to stop tunnel"

# Keep script running
wait $NGROK_PID
