#!/usr/bin/env bash
# Setup Redis MCP server (connect to local Redis instance)
# Ensure Redis is running on localhost:6379
npx -y @modelcontextprotocol/server-redis redis://localhost:6379
