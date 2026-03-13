#!/bin/bash
# main ブランチでのファイル編集をブロックする
branch=$(git branch --show-current 2>/dev/null)
if [ "$branch" = "main" ]; then
  echo "main ブランチでの直接編集はブロックされています。feature ブランチで作業してください。" >&2
  exit 2
fi
exit 0
