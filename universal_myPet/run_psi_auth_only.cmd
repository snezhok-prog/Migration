@echo off
setlocal
cd /d "%~dp0"
python migration.py --profile psi --auth-only --no-prompt --no-interactive
exit /b %ERRORLEVEL%
