@echo off
cd /d "%~dp0"
if not exist venv (
    python -m venv venv
)
call venv\Scripts\activate.bat
pip install -r requirements.txt
if not exist .env (
    copy .env.example .env
    echo Created .env - edit it and add your OPENAI_API_KEY, then re-run this script.
    pause
    exit /b
)
cd backend
uvicorn main:app --reload --port 8420
