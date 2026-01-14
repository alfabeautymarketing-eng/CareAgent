#!/usr/bin/env bash
# Setup GitHub MCP server (Docker container)
# Check for GITHUB_PERSONAL_ACCESS_TOKEN environment variable
if [ -z "$GITHUB_PERSONAL_ACCESS_TOKEN" ]; then
    echo "Error: GITHUB_PERSONAL_ACCESS_TOKEN is not set."
    echo "Please set it: export GITHUB_PERSONAL_ACCESS_TOKEN=your_token_here"
    exit 1
fi
GITHUB_TOKEN="$GITHUB_PERSONAL_ACCESS_TOKEN"

docker run -i --rm \
  -e GITHUB_PERSONAL_ACCESS_TOKEN=$GITHUB_TOKEN \
  ghcr.io/github/github-mcp-server
