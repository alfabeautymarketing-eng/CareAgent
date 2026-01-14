#!/usr/bin/env bash
# Setup Google Drive MCP server (Docker container)
# Mount credentials for service account

# Note: This requires browser authentication on first run.
# Run explicitly in terminal if it hangs.
export GOOGLE_APPLICATION_CREDENTIALS="/Users/aleksandr/Desktop/AgentCare/config/credentials.json"
npx -y @modelcontextprotocol/server-gdrive
