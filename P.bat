@echo off
cd /d "d:/W/Iveco/PowerShell"
git add -A
git commit -m "%~2"
git tag -a %1 -m "%~2"
git push origin main
git push origin %1