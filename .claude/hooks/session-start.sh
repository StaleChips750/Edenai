#!/bin/bash
set -euo pipefail

# Only run in remote Claude Code on the web environments
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

# Add official Anthropic marketplace and install plugins
claude plugin marketplace add https://github.com/anthropics/claude-plugins-official 2>/dev/null || true
claude plugin install superpowers@superpowers-dev 2>/dev/null || \
  (claude plugin marketplace add https://github.com/obra/superpowers 2>/dev/null || true && \
   claude plugin install superpowers@superpowers-dev 2>/dev/null || true)
claude plugin install frontend-design@claude-plugins-official 2>/dev/null || true
claude plugin install code-review@claude-plugins-official 2>/dev/null || true

# Install claude-mem
npx -y claude-mem install --non-interactive 2>/dev/null || true
