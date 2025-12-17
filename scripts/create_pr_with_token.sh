#!/usr/bin/env bash
set -euo pipefail

BRANCH="ci/add-deploy-workflow"
PR_TITLE="chore(ci): add GAS deploy workflow and docs"
PR_BODY=$'Добавляет workflow для автодеплоя и документацию по CLASP.\n\nChecklist:\n- [ ] Проверить secrets\n- [ ] Тесты CI\n'

# Safety checks
if ! command -v gh >/dev/null 2>&1; then
  echo "Error: gh CLI is not installed. Install it and retry: https://cli.github.com/"
  exit 1
fi

# Ensure we are in repository root
REPO_ROOT=$(git rev-parse --show-toplevel 2>/dev/null || true)
if [ -z "$REPO_ROOT" ]; then
  echo "Error: run this script from inside a git repository." >&2
  exit 1
fi
cd "$REPO_ROOT"

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

# Prompt for PAT (hidden input)
read -s -p "Paste GitHub Personal Access Token (it will not be echoed) and press Enter: " GHTOKEN
echo
if [ -z "$GHTOKEN" ]; then
  echo "No token provided. Aborting."; exit 1
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