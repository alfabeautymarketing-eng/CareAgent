#!/bin/bash
set -e

GAS_DIR="gas"
CONFIG_FILE="$GAS_DIR/01Config.js"
MISSING_ID="13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ"

echo "🔧 Patching 01Config.js..."

if [ -f "$CONFIG_FILE" ]; then
    # Check if ID already exists
    if grep -q "$MISSING_ID" "$CONFIG_FILE"; then
        echo "   ID already present."
    else
        # Append the ID to the specific map. Identifying unique string.
        # "const DOC_TO_PROJECT = {" is a good anchor.
        # We will replace "const DOC_TO_PROJECT = {" with "const DOC_TO_PROJECT = {\n  'ID': 'MT',"
        # Using python for safer file editing than sed across platforms might be overkill but safe.
        # Let's use perl -i which works well on mac.
        perl -i -pe "s/const DOC_TO_PROJECT = \{/const DOC_TO_PROJECT = \{\n  '$MISSING_ID': 'MT',/" "$CONFIG_FILE"
        echo "   ✅ Added MT ID to config."
    fi
else
    echo "⚠️  01Config.js not found in gas/. Check copy."
    exit 1
fi
