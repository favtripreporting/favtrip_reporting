@echo off
setlocal
REM ---------------------------------------------------------------------------
REM Run Streamlit UI without a persistent console window.
REM Location: __executable\run_web_windows_silent.bat
REM Behavior: brief flash at launch, then only the browser tab remains.
REM ---------------------------------------------------------------------------

REM Move into the folder of this .bat
pushd "%~dp0"

REM Go to the project root (one level up from __executable)
cd ..

REM Choose Python: prefer venv's interpreter if present
set "PY_VENV=.\.venv\Scripts\python.exe"
set "PY="
if exist "%PY_VENV%" (
  set "PY=%PY_VENV%"
) else (
  for %%P in (python.exe py.exe) do (
    where %%P >nul 2>&1 && (set "PY=%%P" & goto :gotpy)
  )
)
:gotpy
if not defined PY (
  echo [Launcher] Python was not found. Install Python or create .\.venv and try again.
  popd
  exit /b 1
)

REM Streamlit prefs: ensure it opens the browser and stays local
set "STREAMLIT_SERVER_HEADLESS=false"
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"

REM Start Streamlit hidden and detach from this console (which then closes)
REM - We invoke PowerShell only to spawn the hidden child process.
start "" /MIN powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command ^
  "Start-Process -FilePath '%PY%' -ArgumentList '-m','streamlit','run','ui_streamlit.py' -WindowStyle Hidden"

popd
exit /b 0