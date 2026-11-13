@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv (
  echo Dang tao moi truong Python...
  python -m venv .venv || (echo Chua cai Python 3.11+. Tai tai https://www.python.org/downloads/ & pause & exit /b 1)
  .venv\Scripts\python -m pip install --upgrade pip
  .venv\Scripts\pip install -r requirements.txt
)
if not exist .env (
  copy .env.example .env >nul
  echo Da tao file .env - hay dien SUPABASE_URL va SUPABASE_SERVICE_KEY roi chay lai.
  notepad .env
  pause
  exit /b 0
)
start "" http://127.0.0.1:8000
.venv\Scripts\python -m app.main
pause
