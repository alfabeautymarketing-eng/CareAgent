#!/bin/bash
set -e

# Configuration
GAS_DIR="./gas"
PROJECTS_FILE="./config/projects.json"

# 1. Deploy Server Side (Python/Docker)
echo "🚀 Deploying Python Server..."
./deploy.sh

# 2. Deploy Google Apps Script
echo "📜 Deploying Google Apps Script..."

if ! command -v clasp &> /dev/null; then
    echo "⚠️  clasp not found. Skipping GAS deploy."
    echo "   Please install: npm install -g @google/clasp"
    echo "   And login: clasp login"
    exit 0
fi

# Check if projects.json exists
if [ ! -f "$PROJECTS_FILE" ]; then
    echo "⚠️  config/projects.json not found."
    echo "   Creating template..."
    mkdir -p config
    echo '{
    "SS": "SCRIPT_ID_HERE",
    "SK": "SCRIPT_ID_HERE",
    "MT": "SCRIPT_ID_HERE"
}' > "$PROJECTS_FILE"
    echo "   Please edit config/projects.json and add your Script IDs."
    exit 0
fi

# Read projects and deploy
# Uses python to parse json because jq might not be installed
cd "$GAS_DIR"

# Loop through keys in json
for project in $(python3 -c "import json; print(' '.join(json.load(open('../$PROJECTS_FILE')).keys()))"); do
    script_id=$(python3 -c "import json; print(json.load(open('../$PROJECTS_FILE'))['$project'])")
    
    if [ "$script_id" == "SCRIPT_ID_HERE" ]; then
        echo "⚠️  Skipping $project (Script ID not set)"
        continue
    fi

    echo "   Pushing to $project ($script_id)..."
    
    # Create temp .clasp.json
    echo "{
  \"scriptId\": \"$script_id\",
  \"rootDir\": \".\"
}" > .clasp.json

    # Push
    clasp push -f
    
    echo "   ✅ $project updated."
done

rm -f .clasp.json
echo "🎉 All Done!"
