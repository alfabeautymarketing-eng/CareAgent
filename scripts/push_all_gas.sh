#!/bin/bash
set -e

# ============================================================
# push_all_gas.sh — Массовый деплой GAS во все проекты
# ============================================================

GAS_DIR="./gas"

# Projects and their IDs
ID_MT="199Np7xsBiBRQih5_tlUdpt6EmkfRGjZAhTvKm4Ua0Q6XEaMtvAmQUn0g"
ID_SS="1sTgZa-n1aP7oIhyQfPeN8QDgDNnCubqMWAd-TKjKpJXWsQm_ZhXnojPD"
ID_SK="1DJvK1vUT2OTubN0TLdZvsgYMSYByLHl8xTsus3K-KJ-VtJxgGnSw5Ih8"

echo "🚀 Starting GAS mass deploy..."

for PRJ in "MT" "SS" "SK"; do
    case $PRJ in
        MT) ID=$ID_MT ;;
        SS) ID=$ID_SS ;;
        SK) ID=$ID_SK ;;
    esac
    
    echo ""
    echo "----------------------------------------"
    echo "📦 Deploying to $PRJ (ID: $ID)..."
    echo "----------------------------------------"
    
    # Update .clasp.json with the project ID
    printf '{"scriptId":"%s","rootDir":"."}' "$ID" > "$GAS_DIR/.clasp.json"
    
    # Push to GAS
    cd "$GAS_DIR"
    clasp push -f
    cd ..
    
    echo "✅ Success for $PRJ"
done

echo ""
echo "🎉 Mass deploy complete!"
