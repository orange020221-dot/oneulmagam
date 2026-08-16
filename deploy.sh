#!/bin/zsh
set -euo pipefail
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
cd "$(dirname "$0")"

if ! gh auth status >/dev/null 2>&1; then
  echo "GitHub 로그인이 필요합니다."
  echo "https://github.com/login/device"
  gh auth login --hostname github.com --git-protocol https --web
fi

USER_NAME="$(gh api user --jq .login)"
if ! gh repo view "$USER_NAME/oneulmagam" >/dev/null 2>&1; then
  gh repo create oneulmagam --public --source=. --remote=origin --push
else
  git remote get-url origin >/dev/null 2>&1 || git remote add origin "https://github.com/$USER_NAME/oneulmagam.git"
  git push -u origin main
fi

echo "GitHub: https://github.com/$USER_NAME/oneulmagam"
npx --yes vercel --yes --prod
