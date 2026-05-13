@echo off
cd /d "%~dp0"
echo Running DyeFlow RS regression smoke test...
python regression_smoke_test.py
pause
