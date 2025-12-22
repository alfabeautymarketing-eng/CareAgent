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
REVIEWERS=""
LABELS=""

# Parse flags
while [[ $# -gt 0 ]]; do
  case "$1" in
    --use-cwd)
      USE_CWD=true; shift;;
    --token-from-clipboard)
      TOKEN_FROM_CLIP=true; shift;;
    --token)
      TOKEN_ARG="$2"; shift 2;;
    --reviewers)
      REVIEWERS="$2"; shift 2;;
    --labels)
      LABELS="$2"; shift 2;;
    --dry-run)
      DRY_RUN=true; shift;;
    -h|--help)
      echo "Usage: $0 [--use-cwd] [--token-from-clipboard] [--token <token>] [--dry-run] [--reviewers \"user1,user2\"] [--labels \"label1,label2\"]"; exit 0;;
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

# Resolve owner/repo from git remote
REMOTE_URL=$(git remote get-url origin 2>/dev/null || true)
if [[ -z "$REMOTE_URL" ]]; then
  echo "Error: cannot determine git remote 'origin' URL." >&2
  exit 1
fi
if [[ "$REMOTE_URL" =~ github.com[:/]+([^/]+)/([^/.]+)(\.git)?$ ]]; then
  OWNER="${BASH_REMATCH[1]}"
  REPO="${BASH_REMATCH[2]}"
else
  echo "Could not parse origin remote URL: $REMOTE_URL" >&2
  exit 1
fi
FULL_REPO="$OWNER/$REPO"

# Check for existing PR for this head
EXISTING_PR_JSON=$(gh api "repos/$FULL_REPO/pulls?head=$OWNER:$BRANCH" 2>/dev/null || true)
EXISTING_URL=""
if [ -n "$EXISTING_PR_JSON" ]; then
  if printf '%s' "$EXISTING_PR_JSON" | grep -q '"html_url"'; then
    if command -v jq >/dev/null 2>&1; then
      EXISTING_URL=$(printf '%s' "$EXISTING_PR_JSON" | jq -r '.[0].html_url')
    else
      EXISTING_URL=$(printf '%s' "$EXISTING_PR_JSON" | python -c "import sys,json;arr=json.load(sys.stdin);print(arr[0].get('html_url',''))")
    fi
  fi
fi

if [ -n "${EXISTING_URL:-}" ]; then
  echo "A PR already exists: $EXISTING_URL"
  PR_URL="$EXISTING_URL"
else
  # Create PR via gh api (works regardless of gh version flags)
  CREATED_JSON=$(gh api -X POST "repos/$FULL_REPO/pulls" -f title="$PR_TITLE" -f head="$BRANCH" -f base="main" -f body="$PR_BODY" 2>/dev/null || true)
  if [ -z "$CREATED_JSON" ]; then
    echo "gh api: failed to create PR or returned empty response; falling back to interactive 'gh pr create --web'." >&2
    gh pr create --base main --head "$BRANCH" --title "$PR_TITLE" --body "$PR_BODY" --web || true
    echo "Opened web browser to create PR interactively.";
    PR_URL=""
  else
    if command -v jq >/dev/null 2>&1; then
      PR_URL=$(printf '%s' "$CREATED_JSON" | jq -r .html_url)
    else
      PR_URL=$(printf '%s' "$CREATED_JSON" | python -c "import sys,json;print(json.load(sys.stdin).get('html_url',''))")
    fi
    if [ -n "$PR_URL" ]; then
      echo "PR created: $PR_URL"
    else
      echo "PR created but URL could not be extracted. Response:";
      printf '%s
' "$CREATED_JSON"
      echo "You can view PRs with: gh pr list --head \"$BRANCH\""
    fi
  fi
fi

# Determine PR number (from existing or newly created response)
PR_NUMBER=""
if [ -n "${EXISTING_PR_JSON:-}" ]; then
  if command -v jq >/dev/null 2>&1; then
    PR_NUMBER=$(printf '%s' "$EXISTING_PR_JSON" | jq -r '.[0].number // empty')
  else
    PR_NUMBER=$(printf '%s' "$EXISTING_PR_JSON" | python -c "import sys,json;arr=json.load(sys.stdin);print(arr[0].get('number',''))")
  fi
fi
if [ -z "$PR_NUMBER" ] && [ -n "${CREATED_JSON:-}" ]; then
  if command -v jq >/dev/null 2>&1; then
    PR_NUMBER=$(printf '%s' "$CREATED_JSON" | jq -r '.number // empty')
  else
    PR_NUMBER=$(printf '%s' "$CREATED_JSON" | python -c "import sys,json;print(json.load(sys.stdin).get('number',''))")
  fi
fi

# Add reviewers if requested
if [ -n "${REVIEWERS:-}" ]; then
  if [ -z "${PR_NUMBER:-}" ]; then
    echo "Cannot add reviewers: PR number unknown (PR may have been created via web). Please add reviewers manually." >&2
  else
    echo "Adding reviewers: $REVIEWERS"
    REVIEWERS_JSON="[\"${REVIEWERS//,/,\"\"}\"]"
    REQ_REV=$(gh api -X POST "repos/$FULL_REPO/pulls/$PR_NUMBER/requested_reviewers" -f reviewers="$REVIEWERS_JSON" 2>/dev/null || true)
    if [ -n "$REQ_REV" ]; then
      echo "Requested reviewers successfully."
    else
      echo "Failed to request reviewers (check permissions)." >&2
    fi
  fi
fi

# Add labels if requested
if [ -n "${LABELS:-}" ]; then
  if [ -z "${PR_NUMBER:-}" ]; then
    echo "Cannot add labels: PR number unknown (PR may have been created via web). Please add labels manually." >&2
  else
    echo "Adding labels: $LABELS"
    LABELS_JSON="[\"${LABELS//,/,\"\"}\"]"
    LABELS_RESP=$(gh api -X POST "repos/$FULL_REPO/issues/$PR_NUMBER/labels" -f labels="$LABELS_JSON" 2>/dev/null || true)
    if [ -n "$LABELS_RESP" ]; then
      echo "Labels added: $LABELS"
    else
      echo "Failed to add labels (check permissions)." >&2
    fi
  fi
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