# Autosave script
bd sync
git add .
timestamp=$(date "+%Y-%m-%d %H:%M:%S")
git commit -m "Autosave: $timestamp"
bash scripts/push_all_gas.sh
git push
echo "✅ Changes executed and pushed to GitHub."
