#!/usr/bin/env bash
set -euo pipefail

# Defaults
BRANCH="ci/add-deploy-workflow"
PR_TITLE="chore(ci): add GAS deploy workflow and docs"
PR_BODY=$'Добавляет workflow для автодеплоя и документацию по CLASP.\n\nChecklist:\n- [ ] Проверить secrets\n- [ ] Тесты CI\n'

# Options
USE_CWD=false
TOKEN_FROM_CLIP=false
DRY_RUN=false
TOKEN_ARG=""

# Parse flags
while [[ $# -gt 0 ]]; do
  case "$1" in
    --use-cwd)
      USE_CWD=true; shift;;
    --token-from-clipboard)
      TOKEN_FROM_CLIP=true; shift;;
    --token)
      TOKEN_ARG="$2"; shift 2;;
    --dry-run)
      DRY_RUN=true; shift;;
    -h|--help)
      echo "Usage: $0 [--use-cwd] [--token-from-clipboard] [--token <token>] [--dry-run]"; exit 0;;
    *)
      echo "Unknown option: $1"; exit 1;;
  esac
done

# Safety checks
if ! command -v gh >/dev/null 2>&1; then
  echo "Error: gh CLI is not installed. Install it and retry: https://cli.github.com/"
  exit 1
fi

# Ensure we are in repository root unless --use-cwd is set
REPO_ROOT=$(git rev-parse --show-toplevel 2>/dev/null || true)
if [ -z "$REPO_ROOT" ]; then
  echo "Error: run this script from inside a git repository." >&2
  exit 1
fi
if [ "$USE_CWD" = false ]; then
  cd "$REPO_ROOT"
else
  echo "Using current working directory; not switching to repo root (use --use-cwd to override)."
fi

# Make sure branch exists locally
if ! git show-ref --verify --quiet "refs/heads/$BRANCH"; then
  echo "Branch '$BRANCH' not found locally. Create or switch to it and push first." >&2
  echo "Example: git switch -c $BRANCH && git push -u origin $BRANCH"
  exit 1
fi

# Make sure branch is pushed
if ! git ls-remote --exit-code --heads origin "$BRANCH" >/dev/null 2>&1; then
  echo "Branch '$BRANCH' has not been pushed to origin. Pushing now..."
  git push -u origin "$BRANCH"
fi

# Prompt user before requesting token
read -p "Ready to create PR from '$BRANCH' with title '$PR_TITLE'? (type 'yes' to continue): " confirm
if [ "$confirm" != "yes" ]; then
  echo "Aborted by user."; exit 0
fi

# If dry-run requested, exit before prompting for token
if [ "$DRY_RUN" = true ]; then
  echo "DRY RUN: would run 'gh auth login' and create PR from $BRANCH with title: $PR_TITLE"
  exit 0
fi

# Determine token (from arg / env / clipboard / prompt)
if [ -n "$TOKEN_ARG" ]; then
  GHTOKEN="$TOKEN_ARG"
elif [ -n "${GHTOKEN:-}" ]; then
  echo "Using GHTOKEN from environment"
  # use existing GHTOKEN exported in env
elif [ "$TOKEN_FROM_CLIP" = true ]; then
  # try macOS pbpaste then xclip
  if command -v pbpaste >/dev/null 2>&1; then
    GHTOKEN=$(pbpaste)
  elif command -v xclip >/dev/null 2>&1; then
    GHTOKEN=$(xclip -o -selection clipboard)
  else
    echo "Clipboard tool not found (pbpaste/xclip)."; exit 1
  fi
  echo "Using token from clipboard"
else
  read -s -p "Paste GitHub Personal Access Token (it will not be echoed) and press Enter (or press Enter to try clipboard): " GHTOKEN
  echo
  if [ -z "$GHTOKEN" ]; then
    if command -v pbpaste >/dev/null 2>&1; then
      GHTOKEN=$(pbpaste)
      echo "Using clipboard content as token"
    elif command -v xclip >/dev/null 2>&1; then
      GHTOKEN=$(xclip -o -selection clipboard)
      echo "Using clipboard content as token"
    else
      echo "No token provided and clipboard is unavailable. Aborting."; exit 1
    fi
  fi
fi

# If still empty, error
if [ -z "${GHTOKEN:-}" ]; then
  echo "No token provided. Aborting."; exit 1
fi

if [ "$DRY_RUN" = true ]; then
  echo "DRY RUN: would run 'gh auth login' and create PR from $BRANCH with title: $PR_TITLE"; exit 0
fi

# Login with token (provided on stdin)
echo "$GHTOKEN" | gh auth login --with-token || { echo "gh auth login failed"; exit 1; }

# Create PR
echo "Creating PR..."
PR_URL=$(gh pr create --base main --head "$BRANCH" --title "$PR_TITLE" --body "$PR_BODY" --web 2>/dev/null || true)

# gh pr create may open web or return data; try to show it
if [ -z "$PR_URL" ]; then
  echo "Attempting to fetch the created PR in the browser..."
  gh pr view --web
else
  echo "PR created: $PR_URL"
fi

# Offer to logout (revoke local auth)
read -p "Do you want to logout from gh on this machine now? (y/N): " logout
if [ "$logout" = "y" ] || [ "$logout" = "Y" ]; then
  gh auth logout --hostname github.com --confirm || true
  echo "Logged out locally. Remember to revoke the PAT on GitHub if it was temporary."
else
  echo "You remain logged in. Revoke the PAT on GitHub when you no longer need it."
fi

# Cleanup
unset GHTOKEN

echo "Done."