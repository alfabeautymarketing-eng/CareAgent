#!/usr/bin/env bash
# Setup Google Sheets MCP server (NPM: mcp-gsheets)
# Requires GOOGLE_APPLICATION_CREDENTIALS environment variable

export GOOGLE_APPLICATION_CREDENTIALS="/Users/aleksandr/Desktop/AgentCare/config/credentials.json"
export GOOGLE_PROJECT_ID="gen-lang-client-0520143643"
npx -y mcp-gsheets
