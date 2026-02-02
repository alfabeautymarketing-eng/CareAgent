#!/bin/bash
# Script to load latest changes when starting work on a device

echo "🔄 Pulling latest changes from GitHub..."
git pull

echo "📦 Updating Python dependencies (Poetry)..."
poetry install

echo "🐝 Syncing Beads tasks..."
# Try to find bd in common locations if not in PATH
BD_PATH=$(which bd || echo "/Users/aleksandr/.local/bin/bd")
if [ -f "$BD_PATH" ]; then
    "$BD_PATH" sync
else
    echo "⚠️ Warning: Beads (bd) not found. Please ensure it's installed."
fi

echo "✅ Project is up to date and ready for work!"
