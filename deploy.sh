#!/usr/bin/env bash
set -e

cd "$(dirname "$0")"

git add .

if git diff --cached --quiet; then
  echo "변경사항 없음 — push만 진행합니다."
else
  git commit -m "deploy: $(date '+%Y-%m-%d %H:%M')"
fi

git push origin main

echo "배포 완료 — https://bongwoo-gif.github.io/baemin-id-photo/"
