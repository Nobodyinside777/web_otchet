@echo off
rem Launcher for the "Prognoz vyruchki" app. Keep this file ASCII-only:
rem cmd.exe misreads non-Latin characters in .bat files.
setlocal
cd /d "%~dp0"

set "VENV_PY=.venv\Scripts\python.exe"
set "URL=http://127.0.0.1:8501"

rem --- 1. Virtual environment (created once) ---
if exist "%VENV_PY%" goto deps

echo First run: creating virtual environment in .venv ...
set "PYCMD="
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)" >nul 2>&1 && set "PYCMD=py -3"
if not defined PYCMD python -c "import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)" >nul 2>&1 && set "PYCMD=python"
if not defined PYCMD goto no_python

%PYCMD% -m venv .venv
if errorlevel 1 goto fail

rem --- 2. Libraries (reinstalled only when requirements.txt changes) ---
:deps
fc /b requirements.txt .venv\requirements.installed >nul 2>&1
if not errorlevel 1 goto start

echo Installing libraries, this takes a couple of minutes ...
"%VENV_PY%" -m pip install --disable-pip-version-check -q --upgrade pip
"%VENV_PY%" -m pip install --disable-pip-version-check -q -r requirements.txt
if errorlevel 1 goto fail
copy /y requirements.txt .venv\requirements.installed >nul

rem --- 3. Start: the browser opens as soon as the server is ready ---
:start
netstat -ano | findstr /c:"127.0.0.1:8501" | findstr LISTENING >nul
if errorlevel 1 goto launch
echo.
echo The app is already running in another window (port 8501).
echo Close that window first if you want to start a fresh copy, then run run.bat again.
echo.
pause
exit /b 1

:launch
echo.
echo The app is starting: %URL%
echo To stop it, close this window.
echo.
start "" /b powershell -NoProfile -WindowStyle Hidden -Command "for ($i = 0; $i -lt 60; $i++) { try { (New-Object Net.Sockets.TcpClient('127.0.0.1', 8501)).Close(); break } catch { Start-Sleep 1 } }; Start-Process '%URL%'"
"%VENV_PY%" -m streamlit run app\main.py
goto end

:no_python
echo.
echo Python 3.12 or newer was not found.
echo Install it from https://www.python.org/downloads/ (tick "Add python.exe to PATH") and run run.bat again.
goto fail_pause

:fail
echo.
echo Something went wrong. Copy the text above and send it to the developer.

:fail_pause
pause
exit /b 1

:end
endlocal
