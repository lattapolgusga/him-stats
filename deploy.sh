#!/usr/bin/env bash
# อัปเดตระบบหลังบ้านจาก Terminal: ./deploy.sh
# ครั้งแรก: npx @google/clasp login  และเปิด https://script.google.com/home/usersettings → เปิด Google Apps Script API
set -e
cd "$(dirname "$0")/apps-script"
npx -y @google/clasp push -f
if [ -f ../.deployment-id ]; then
  npx -y @google/clasp deploy -i "$(cat ../.deployment-id)" -d "update $(date +%F_%H%M)"
else
  out=$(npx -y @google/clasp deploy -d "first deploy")
  echo "$out"
  echo "$out" | grep -oE 'AKfy[A-Za-z0-9_-]+' | head -1 > ../.deployment-id
  echo "API URL: https://script.google.com/macros/s/$(cat ../.deployment-id)/exec"
fi
