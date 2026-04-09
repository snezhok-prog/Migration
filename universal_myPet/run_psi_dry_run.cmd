@echo off
setlocal
cd /d "%~dp0"
python migration.py --profile psi --dry-run --skip-auth --no-interactive --limit 1
exit /b %ERRORLEVEL%
