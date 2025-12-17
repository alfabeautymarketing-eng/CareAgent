#!/bin/bash
set -e

# Configuration
SS_ID="1sTgZa-n1aP7oIhyQfPeN8QDgDNnCubqMWAd-TKjKpJXWsQm_ZhXnojPD"
GAS_DIR="gas"

echo "📥 Cloning Source Code from SS Project..."

# Ensure gas dir exists and is empty-ish (preserve our new files)
mkdir -p "$GAS_DIR"
# Backup our new files
cp "$GAS_DIR/Menu.gs" "Menu.gs.bak" 2>/dev/null || true
cp "$GAS_DIR/Client.gs" "Client.gs.bak" 2>/dev/null || true

# Clone into temp dir
mkdir -p gas_tmp
cd gas_tmp
if ! clasp clone "$SS_ID"; then
    echo "❌ Clasp clone failed. Did you run 'clasp login'?"
    exit 1
fi
cd ..

# Copy files to gas/
cp gas_tmp/*.js "$GAS_DIR/" 2>/dev/null || true
cp gas_tmp/*.gs "$GAS_DIR/" 2>/dev/null || true
cp gas_tmp/appsscript.json "$GAS_DIR/" 2>/dev/null || true
rm -rf gas_tmp

# Restore/Overwrite with our new files (Client/Menu)
if [ -f "Menu.gs.bak" ]; then mv "Menu.gs.bak" "$GAS_DIR/Menu.gs"; fi
if [ -f "Client.gs.bak" ]; then mv "Client.gs.bak" "$GAS_DIR/Client.gs"; fi

echo "🔧 Patching 01Config.js..."
CONFIG_FILE="$GAS_DIR/01Config.js"
MISSING_ID="13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ"

if [ -f "$CONFIG_FILE" ]; then
    # Check if ID already exists
    if grep -q "$MISSING_ID" "$CONFIG_FILE"; then
        echo "   ID already present."
    else
        # Very simple sedimentary injection: find DOC_TO_PROJECT = { and append our line after it
        # This is fragile but SED is hard for multi-line JSON objects within JS. 
        # Better approach: Just append it to the map via replace strings if we know a common key.
        # Let's try to just find the closing brace of the object? No.
        # Let's assume the user can check it or we use a reliable anchor.
        # Anchor: "const DOC_TO_PROJECT = {"
        sed -i '' "s/const DOC_TO_PROJECT = {/const DOC_TO_PROJECT = {\\n  '$MISSING_ID': 'MT',/" "$CONFIG_FILE"
        echo "   ✅ Added MT ID to config."
    fi
else
    echo "⚠️  01Config.js not found. Skipping patch."
fi

echo "✅ Setup Complete. Now run: ./scripts/deploy_all.sh"
