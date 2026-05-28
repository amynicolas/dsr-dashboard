@echo off
REM ─────────────────────────────────────────────────────────────────────────────
REM  DSR Dashboard — GitHub Pages Deploy Script
REM  Run this after processing each day's DSR Excel to push to your live site.
REM  One-time setup required: see GITHUB_PAGES_SETUP.md
REM ─────────────────────────────────────────────────────────────────────────────

cd /d "%~dp0"

REM Copy dashboard as index.html (GitHub Pages serves index.html as root page)
copy /Y DSR_Dashboard.html index.html

REM Stage all changes
git add DSR_Dashboard.html index.html knowledge\dsr_history.json knowledge\targets.json

REM Commit with today's date
for /f "tokens=2 delims==" %%I in ('"wmic os get localdatetime /value"') do set datetime=%%I
set DATE_SLUG=%datetime:~0,8%
git commit -m "DSR update %DATE_SLUG%"

REM Push to GitHub (main branch → GitHub Pages auto-deploys)
git push origin main

echo.
echo Done! Your dashboard will be live in ~30 seconds at:
echo   https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/
echo.
pause
