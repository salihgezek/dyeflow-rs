@echo off
cd /d "%~dp0"
py -m uvicorn main:app --host 127.0.0.1 --port 8002 --reload
pause
