#!/usr/bin/env bash
# Setup Google AppScript MCP server (Docker container)
# Mount credentials for service account

docker run -i --rm \
  -v /Users/aleksandr/Desktop/AgentCare/config/credentials.json:/app/credentials.json:ro \
  -e GOOGLE_APPLICATION_CREDENTIALS=/app/credentials.json \
  ghcr.io/mcp-server/google-appsscript
