@echo off
REM Start Gridlint on http://127.0.0.1:8000
cd /d "%~dp0"
if not exist ".venv" (
  echo Creating a virtual environment...
  python -m venv .venv || goto :fail
)
call .venv\Scripts\activate.bat
python -m pip install -q -r requirements.txt || goto :fail
echo.
echo Gridlint is starting on http://127.0.0.1:8000
echo Press Ctrl+C to stop.
python -m gridlint serve --port 8000
goto :eof
:fail
echo.
echo Could not start. Make sure Python 3.10 or newer is installed and on PATH.
pause
