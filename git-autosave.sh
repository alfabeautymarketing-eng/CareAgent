#!/bin/bash
# Autosave script
git add .
timestamp=$(date "+%Y-%m-%d %H:%M:%S")
git commit -m "Autosave: $timestamp"
git push
echo "✅ Changes executed and pushed to GitHub."
