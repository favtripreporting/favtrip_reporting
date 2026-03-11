# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env
.env.example
.gitignore
FavTripPipeline.spec
FavTripPipelineUI.spec
cli.py
credentials.json
last_run.log
launcher_streamlit.py
requirements.txt
setup_py2app.py
ui_streamlit.py
web_url_credentials.json
  dev_input_sales_files/
    Advanded Testing Sales Report - 1 Week.xlsx
    Advanded Testing Sales Report - 2 Weeks.xlsx  [skipped: too large]
    Advanded Testing Sales Report - Bad End.xlsx  [skipped: too large]
    Advanded Testing Sales Report - Bad Start.xlsx  [skipped: too large]
    Testing Sales Report - Week 2.xlsx  [skipped: too large]
    Testing Sales Report - Week 3.xlsx  [skipped: too large]
    Testing Sales Report - Week 4.xlsx  [skipped: too large]
    Testing Sales Report - Week 5.xlsx  [skipped: too large]
  documentation/
    PROJECT_SNAPSHOT_CODEBUNDLE.md
    README.md
    generate_code_bundle.py
    git_workflow.txt
    requirements.txt
  favtrip/
    __init__.py
    config.py
    config_store.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    sheets_utils.py
  __executable/
    run_windows.bat
---
### file: .env

```
CALC_SPREADSHEET_ID=1gJrQ8W8MExFJzriOoqX3Lyb0n-STbhDtuIqAK617J2E
INCOMING_FOLDER_ID=1kXZoqo0Wa6YM8W_kQV08EmZgLhZ1iDdA
MANAGER_REPORT_FOLDER_ID=1LoP6A9RMvahKyp-47kgl4EGGHIsmcn9E
ORDER_REPORT_FOLDER_ID=1obUeSbrypEh8zvdSw87yihHE3rnIvxKK
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations
TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p
TO_RECIPIENTS=ryan-morrow@uiowa.edu
CC_RECIPIENTS=
USE_ALL_REPORT_KEYS=False
REPORT_KEY_RUN_LIST=COFFEE
REPORT_KEY_RECIPIENTS={}
DEFAULT_ORDER_RECIPIENTS=TO_RECIPIENTS
INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=False
SEND_SEPARATE_FULL_ORDER_EMAIL=False
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send
FORCE_REAUTH=False
REDIRECT_PORT=0
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .gitignore

```
*.env
*credentials.json
token.json

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: __executable/run_windows.bat

```bat
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
```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: credentials.json

```json
{"installed":{"client_id":"674901450584-p8kvvj127a5nghs7lohkmn8ifkjuebt6.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
```

---
### file: documentation/PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env
.env.example
.gitignore
FavTripPipeline.spec
FavTripPipelineUI.spec
cli.py
credentials.json
last_run.log
launcher_streamlit.py
requirements.txt
setup_py2app.py
ui_streamlit.py
web_url_credentials.json
  dev_input_sales_files/
    Advanded Testing Sales Report - 1 Week.xlsx
    Advanded Testing Sales Report - 2 Weeks.xlsx  [skipped: too large]
    Advanded Testing Sales Report - Bad End.xlsx  [skipped: too large]
    Advanded Testing Sales Report - Bad Start.xlsx  [skipped: too large]
    Testing Sales Report - Week 2.xlsx  [skipped: too large]
    Testing Sales Report - Week 3.xlsx  [skipped: too large]
    Testing Sales Report - Week 4.xlsx  [skipped: too large]
    Testing Sales Report - Week 5.xlsx  [skipped: too large]
  documentation/
    PROJECT_SNAPSHOT_CODEBUNDLE.md
    README.md
    generate_code_bundle.py
    git_workflow.txt
    requirements.txt
  favtrip/
    __init__.py
    config.py
    config_store.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    sheets_utils.py
  __executable/
    run_windows.bat
---
### file: .env

```
CALC_SPREADSHEET_ID=1gJrQ8W8MExFJzriOoqX3Lyb0n-STbhDtuIqAK617J2E
INCOMING_FOLDER_ID=1kXZoqo0Wa6YM8W_kQV08EmZgLhZ1iDdA
MANAGER_REPORT_FOLDER_ID=1LoP6A9RMvahKyp-47kgl4EGGHIsmcn9E
ORDER_REPORT_FOLDER_ID=1obUeSbrypEh8zvdSw87yihHE3rnIvxKK
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations
TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p
TO_RECIPIENTS=ryan-morrow@uiowa.edu
CC_RECIPIENTS=
USE_ALL_REPORT_KEYS=False
REPORT_KEY_RUN_LIST=COFFEE
REPORT_KEY_RECIPIENTS={}
DEFAULT_ORDER_RECIPIENTS=TO_RECIPIENTS
INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=False
SEND_SEPARATE_FULL_ORDER_EMAIL=False
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send
FORCE_REAUTH=False
REDIRECT_PORT=0
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .gitignore

```
*.env
*credentials.json
token.json

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: __executable/run_windows.bat

```bat
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
```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: credentials.json

```json
{"installed":{"client_id":"674901450584-p8kvvj127a5nghs7lohkmn8ifkjuebt6.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
```

---
### file: documentation/PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env
.env.example
.gitignore
FavTripPipeline.spec
FavTripPipelineUI.spec
cli.py
credentials.json
last_run.log
launcher_streamlit.py
requirements.txt
setup_py2app.py
ui_streamlit.py
web_url_credentials.json
  dev_input_sales_files/
    Testing Sales Report - Week 2.xlsx  [skipped: too large]
    Testing Sales Report - Week 3.xlsx  [skipped: too large]
    Testing Sales Report - Week 4.xlsx  [skipped: too large]
    Testing Sales Report - Week 5.xlsx  [skipped: too large]
  documentation/
    PROJECT_SNAPSHOT_CODEBUNDLE.md
    README.md
    generate_code_bundle.py
    git_workflow.txt
    requirements.txt
  favtrip/
    __init__.py
    config.py
    config_store.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    sheets_utils.py
  __executable/
    run_windows.bat
---
### file: .env

```
CALC_SPREADSHEET_ID=1gJrQ8W8MExFJzriOoqX3Lyb0n-STbhDtuIqAK617J2E
INCOMING_FOLDER_ID=1kXZoqo0Wa6YM8W_kQV08EmZgLhZ1iDdA
MANAGER_REPORT_FOLDER_ID=1LoP6A9RMvahKyp-47kgl4EGGHIsmcn9E
ORDER_REPORT_FOLDER_ID=1obUeSbrypEh8zvdSw87yihHE3rnIvxKK
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations
TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p
TO_RECIPIENTS=ryan-morrow@uiowa.edu
CC_RECIPIENTS=
USE_ALL_REPORT_KEYS=False
REPORT_KEY_RUN_LIST=COFFEE
REPORT_KEY_RECIPIENTS={}
DEFAULT_ORDER_RECIPIENTS=TO_RECIPIENTS
INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=False
SEND_SEPARATE_FULL_ORDER_EMAIL=False
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send
FORCE_REAUTH=False
REDIRECT_PORT=0
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .gitignore

```
*.env
*credentials.json
token.json

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: __executable/run_windows.bat

```bat
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
```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: credentials.json

```json
{"installed":{"client_id":"674901450584-p8kvvj127a5nghs7lohkmn8ifkjuebt6.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
```

---
### file: documentation/PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env
.env.example
.gitignore
FavTripPipeline.spec
FavTripPipelineUI.spec
cli.py
credentials.json
last_run.log
launcher_streamlit.py
requirements.txt
setup_py2app.py
ui_streamlit.py
web_url_credentials.json
  dev_input_sales_files/
    Testing Sales Report - Week 2.xlsx  [skipped: too large]
    Testing Sales Report - Week 3.xlsx  [skipped: too large]
    Testing Sales Report - Week 4.xlsx  [skipped: too large]
    Testing Sales Report - Week 5.xlsx  [skipped: too large]
  documentation/
    PROJECT_SNAPSHOT_CODEBUNDLE.md
    README.md
    generate_code_bundle.py
    git_workflow.txt
    requirements.txt
  favtrip/
    __init__.py
    config.py
    config_store.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    sheets_utils.py
  __executable/
    run_windows.bat
---
### file: .env

```
CALC_SPREADSHEET_ID=1gJrQ8W8MExFJzriOoqX3Lyb0n-STbhDtuIqAK617J2E
INCOMING_FOLDER_ID=1kXZoqo0Wa6YM8W_kQV08EmZgLhZ1iDdA
MANAGER_REPORT_FOLDER_ID=1LoP6A9RMvahKyp-47kgl4EGGHIsmcn9E
ORDER_REPORT_FOLDER_ID=1obUeSbrypEh8zvdSw87yihHE3rnIvxKK
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations
TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p
TO_RECIPIENTS=ryan-morrow@uiowa.edu
CC_RECIPIENTS=
USE_ALL_REPORT_KEYS=False
REPORT_KEY_RUN_LIST=COFFEE
REPORT_KEY_RECIPIENTS={}
DEFAULT_ORDER_RECIPIENTS=TO_RECIPIENTS
INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=False
SEND_SEPARATE_FULL_ORDER_EMAIL=False
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send
FORCE_REAUTH=False
REDIRECT_PORT=0
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .gitignore

```
*.env
*credentials.json
token.json

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: __executable/run_windows.bat

```bat
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
```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: credentials.json

```json
{"installed":{"client_id":"674901450584-p8kvvj127a5nghs7lohkmn8ifkjuebt6.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
```

---
### file: documentation/PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env
.env.example
.gitignore
FavTripPipeline.spec
FavTripPipelineUI.spec
cli.py
credentials.json
last_run.log
launcher_streamlit.py
requirements.txt
setup_py2app.py
ui_streamlit.py
web_url_credentials.json
  dev_input_sales_files/
    Testing Sales Report - Week 2.xlsx  [skipped: too large]
    Testing Sales Report - Week 3.xlsx  [skipped: too large]
    Testing Sales Report - Week 4.xlsx  [skipped: too large]
    Testing Sales Report - Week 5.xlsx  [skipped: too large]
  documentation/
    PROJECT_SNAPSHOT_CODEBUNDLE.md
    README.md
    generate_code_bundle.py
    git_workflow.txt
    requirements.txt
  favtrip/
    __init__.py
    config.py
    config_store.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    sheets_utils.py
  __executable/
    run_windows.bat
---
### file: .env

```
CALC_SPREADSHEET_ID=1gJrQ8W8MExFJzriOoqX3Lyb0n-STbhDtuIqAK617J2E
INCOMING_FOLDER_ID=1kXZoqo0Wa6YM8W_kQV08EmZgLhZ1iDdA
MANAGER_REPORT_FOLDER_ID=1LoP6A9RMvahKyp-47kgl4EGGHIsmcn9E
ORDER_REPORT_FOLDER_ID=1obUeSbrypEh8zvdSw87yihHE3rnIvxKK
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations
TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p
TO_RECIPIENTS=ryan-morrow@uiowa.edu
CC_RECIPIENTS=
USE_ALL_REPORT_KEYS=False
REPORT_KEY_RUN_LIST=COFFEE
REPORT_KEY_RECIPIENTS={}
DEFAULT_ORDER_RECIPIENTS=TO_RECIPIENTS
INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=False
SEND_SEPARATE_FULL_ORDER_EMAIL=False
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send
FORCE_REAUTH=False
REDIRECT_PORT=0
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .gitignore

```
*.env
*credentials.json
token.json

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: __executable/run_windows.bat

```bat
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
```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: credentials.json

```json
{"installed":{"client_id":"674901450584-p8kvvj127a5nghs7lohkmn8ifkjuebt6.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
```

---
### file: documentation/PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env
.env.example
.gitignore
FavTripPipeline.spec
FavTripPipelineUI.spec
cli.py
credentials.json
last_run.log
launcher_streamlit.py
requirements.txt
setup_py2app.py
ui_streamlit.py
web_url_credentials.json
  documentation/
    PROJECT_SNAPSHOT_CODEBUNDLE.md
    README.md
    generate_code_bundle.py
    requirements.txt
  favtrip/
    __init__.py
    config.py
    config_store.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    sheets_utils.py
  __executable/
    run_windows.bat
---
### file: .env

```
CALC_SPREADSHEET_ID=1gJrQ8W8MExFJzriOoqX3Lyb0n-STbhDtuIqAK617J2E
INCOMING_FOLDER_ID=1kXZoqo0Wa6YM8W_kQV08EmZgLhZ1iDdA
MANAGER_REPORT_FOLDER_ID=1LoP6A9RMvahKyp-47kgl4EGGHIsmcn9E
ORDER_REPORT_FOLDER_ID=1obUeSbrypEh8zvdSw87yihHE3rnIvxKK
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations
TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p
TO_RECIPIENTS=ryan-morrow@uiowa.edu
CC_RECIPIENTS=
USE_ALL_REPORT_KEYS=False
REPORT_KEY_RUN_LIST=COFFEE
REPORT_KEY_RECIPIENTS={}
DEFAULT_ORDER_RECIPIENTS=TO_RECIPIENTS
INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=False
SEND_SEPARATE_FULL_ORDER_EMAIL=False
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send
FORCE_REAUTH=False
REDIRECT_PORT=0
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .gitignore

```
*.env
*credentials.json
token.json

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: __executable/run_windows.bat

```bat
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
```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: credentials.json

```json
{"installed":{"client_id":"674901450584-p8kvvj127a5nghs7lohkmn8ifkjuebt6.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
```

---
### file: documentation/PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env
.env.example
FavTripPipeline.spec
FavTripPipelineUI.spec
cli.py
credentials.json
last_run.log
launcher_streamlit.py
setup_py2app.py
ui_streamlit.py
  documentation/
    PROJECT_SNAPSHOT_CODEBUNDLE.md
    README.md
    generate_code_bundle.py
    requirements.txt
  favtrip/
    __init__.py
    config.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    sheets_utils.py
  __executable/
---
### file: .env

```
CALC_SPREADSHEET_ID=1gJrQ8W8MExFJzriOoqX3Lyb0n-STbhDtuIqAK617J2E
INCOMING_FOLDER_ID=1kXZoqo0Wa6YM8W_kQV08EmZgLhZ1iDdA
MANAGER_REPORT_FOLDER_ID=1LoP6A9RMvahKyp-47kgl4EGGHIsmcn9E
ORDER_REPORT_FOLDER_ID=1obUeSbrypEh8zvdSw87yihHE3rnIvxKK
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations
TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p
TO_RECIPIENTS=ryan-morrow@uiowa.edu
CC_RECIPIENTS=
USE_ALL_REPORT_KEYS=False
REPORT_KEY_RUN_LIST=COFFEE
REPORT_KEY_RECIPIENTS={}
DEFAULT_ORDER_RECIPIENTS=TO_RECIPIENTS
INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=False
SEND_SEPARATE_FULL_ORDER_EMAIL=False
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send
FORCE_REAUTH=False
REDIRECT_PORT=0
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: credentials.json

```json
{"installed":{"client_id":"674901450584-p8kvvj127a5nghs7lohkmn8ifkjuebt6.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
```

---
### file: documentation/PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env
.env.example
FavTripPipeline.spec
FavTripPipelineUI.spec
cli.py
credentials.json
last_run.log
launcher_streamlit.py
setup_py2app.py
ui_streamlit.py
  documentation/
    PROJECT_SNAPSHOT_CODEBUNDLE.md
    README.md
    generate_code_bundle.py
    requirements.txt
  favtrip/
    __init__.py
    config.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    sheets_utils.py
  __executable/
    mac_os_read_first.txt
    run_web_mac.command
    run_web_windows.bat
---
### file: .env

```
CALC_SPREADSHEET_ID=1gJrQ8W8MExFJzriOoqX3Lyb0n-STbhDtuIqAK617J2E
INCOMING_FOLDER_ID=1kXZoqo0Wa6YM8W_kQV08EmZgLhZ1iDdA
MANAGER_REPORT_FOLDER_ID=1LoP6A9RMvahKyp-47kgl4EGGHIsmcn9E
ORDER_REPORT_FOLDER_ID=1obUeSbrypEh8zvdSw87yihHE3rnIvxKK
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations
TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p
TO_RECIPIENTS=ryan-morrow@uiowa.edu
CC_RECIPIENTS=
USE_ALL_REPORT_KEYS=False
REPORT_KEY_RUN_LIST=COFFEE
REPORT_KEY_RECIPIENTS={}
DEFAULT_ORDER_RECIPIENTS=TO_RECIPIENTS
INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=False
SEND_SEPARATE_FULL_ORDER_EMAIL=False
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send
FORCE_REAUTH=False
REDIRECT_PORT=0
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: __executable/mac_os_read_first.txt

```text
You must run this on your terminal before it will be runnable

chmod +x "{FULL PATH}/run_web_mac.command"
```

---
### file: __executable/run_web_mac.command

```
#!/bin/bash
cd "$(dirname "$0")"

if [ -f ".venv/bin/activate" ]; then
    source ".venv/bin/activate"
fi

nohup python3 -m streamlit run ui_streamlit.py >/dev/null 2>&1 &
```

---
### file: __executable/run_web_windows.bat

```bat

#!/bin/bash

# Move into the REAL project root (one folder up)
cd "$(dirname "$0")/.."

# Now the working directory is favtrip_pipeline/
# and credentials.json MUST be here

if [ -f ".venv/bin/activate" ]; then
    source ".venv/bin/activate"
fi

# VERY IMPORTANT: prevent Streamlit from resetting CWD
export STREAMLIT_SERVER_HEADLESS=true

# Launch UI
nohup python3 -m streamlit run ui_streamlit.py >/dev/null 2>&1 &

```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: credentials.json

```json
{"installed":{"client_id":"674901450584-p8kvvj127a5nghs7lohkmn8ifkjuebt6.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
```

---
### file: documentation/PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                PROJECT_SNAPSHOT_CODEBUNDLE.md
README.md
generate_code_bundle.py
requirements.txt
---
### file: PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                PROJECT_SNAPSHOT_CODEBUNDLE.md
README.md
generate_code_bundle.py
requirements.txt
---
### file: PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                PROJECT_SNAPSHOT_CODEBUNDLE.md
README.md
generate_code_bundle.py
requirements.txt
---
### file: PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env
.env.example
FavTripPipeline.spec
FavTripPipelineUI.spec
README.md
build_macos.sh
build_windows.bat
cli.py
credentials.json
generate_code_bundle.py
last_run.log
launcher_streamlit.py
requirements.txt
run_web.sh
run_web_windows.bat
setup_py2app.py
ui_streamlit.py
  favtrip/
    __init__.py
    config.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    sheets_utils.py
---
### file: .env

```
CALC_SPREADSHEET_ID=1gJrQ8W8MExFJzriOoqX3Lyb0n-STbhDtuIqAK617J2E
INCOMING_FOLDER_ID=1kXZoqo0Wa6YM8W_kQV08EmZgLhZ1iDdA
MANAGER_REPORT_FOLDER_ID=1LoP6A9RMvahKyp-47kgl4EGGHIsmcn9E
ORDER_REPORT_FOLDER_ID=1obUeSbrypEh8zvdSw87yihHE3rnIvxKK
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations
TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p
TO_RECIPIENTS=ryan-morrow@uiowa.edu
CC_RECIPIENTS=
USE_ALL_REPORT_KEYS=False
REPORT_KEY_RUN_LIST=COFFEE
REPORT_KEY_RECIPIENTS={}
DEFAULT_ORDER_RECIPIENTS=TO_RECIPIENTS
INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=False
SEND_SEPARATE_FULL_ORDER_EMAIL=False
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send
FORCE_REAUTH=False
REDIRECT_PORT=0
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: build_macos.sh

```bash
#!/usr/bin/env bash
set -euo pipefail
python3 -m pip install --upgrade pip
pip3 install -r requirements.txt pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# Optional GUI wrapper for Streamlit
# pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py

```

---
### file: build_windows.bat

```bat
@echo off
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline --add-data "credentials.json;." --add-data ".env;." cli.py

```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: credentials.json

```json
{"installed":{"client_id":"674901450584-p8kvvj127a5nghs7lohkmn8ifkjuebt6.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
```

---
### file: favtrip/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: favtrip/config.py

```python
from __future__ import annotations
import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv

_BOOL_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _b(s: str, default: bool = False) -> bool:
    if s is None:
        return default
    return str(s).strip().lower() in _BOOL_TRUE


def _csv(s: str) -> List[str]:
    if not s:
        return []
    return [p.strip() for p in s.split(",") if p.strip()]


def _json_or_empty(s: str):
    if not s:
        return {}
    try:
        return json.loads(s)
    except Exception:
        return {}


@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"

    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"

    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None

    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None

    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True

    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285

    HTTP_TIMEOUT_SECONDS: int = 300

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)
        return Config(
            CALC_SPREADSHEET_ID=os.getenv("CALC_SPREADSHEET_ID", ""),
            INCOMING_FOLDER_ID=os.getenv("INCOMING_FOLDER_ID", ""),
            MANAGER_REPORT_FOLDER_ID=os.getenv("MANAGER_REPORT_FOLDER_ID", ""),
            ORDER_REPORT_FOLDER_ID=os.getenv("ORDER_REPORT_FOLDER_ID", ""),
            GID_MANAGER_PDF=os.getenv("GID_MANAGER_PDF", "1921812573"),
            GID_ORDER_CSV=os.getenv("GID_ORDER_CSV", "1875928148"),
            LOCATION_SHEET_TITLE=os.getenv("LOCATION_SHEET_TITLE", "REFR: Values"),
            LOCATION_NAMED_RANGE=os.getenv("LOCATION_NAMED_RANGE", "_locations"),
            TIMESTAMP_TZ=os.getenv("TIMESTAMP_TZ", "America/Chicago"),
            TIMESTAMP_FMT=os.getenv("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p"),
            TO_RECIPIENTS=_csv(os.getenv("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_csv(os.getenv("CC_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_b(os.getenv("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=_csv(os.getenv("REPORT_KEY_RUN_LIST", "")),
            REPORT_KEY_RECIPIENTS=_json_or_empty(os.getenv("REPORT_KEY_RECIPIENTS", "")),
            DEFAULT_ORDER_RECIPIENTS=_csv(os.getenv("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_b(
                os.getenv("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_b(
                os.getenv("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),
            SCOPES=_csv(
                os.getenv(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_b(os.getenv("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(os.getenv("REDIRECT_PORT", "58285")),
            HTTP_TIMEOUT_SECONDS=int(os.getenv("HTTP_TIMEOUT_SECONDS", "300")),
        )

    def to_env(self) -> str:
        # Serialize to .env format (simple)
        data = asdict(self)
        # Convert list and dict fields
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")
```

---
### file: favtrip/drive_utils.py

```python
from __future__ import annotations
import io
from googleapiclient.http import MediaIoBaseUpload


def find_latest_sheet(drive_svc, folder_id: str):
    q = (
        f"'{folder_id}' in parents and "
        "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    )
    resp = drive_svc.files().list(
        q=q, orderBy="createdTime desc", pageSize=1,
        fields="files(id,name,createdTime)"
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def upload_to_drive(drive_svc, data: bytes, name: str, mime: str, folder_id: str, to_sheet: bool=False):
    meta = {"name": name, "parents": [folder_id]}
    if to_sheet:
        meta["mimeType"] = "application/vnd.google-apps.spreadsheet"
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
    return drive_svc.files().create(
        body=meta, media_body=media, fields="id,name,mimeType,webViewLink"
    ).execute()

```

---
### file: favtrip/gmail_utils.py

```python
from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {ts} – {location}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Automated")
    msg.add_alternative(
        f"<p>Hi team,</p><p>Manager Report ({location})</p>"
        f"<a href='{pdf_link}'>Backup Link</a>", subtype="html"
    )
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)

```

---
### file: favtrip/google_client.py

```python
from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail
```

---
### file: favtrip/logger.py

```python
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass
```

---
### file: favtrip/pipeline.py

```python
from __future__ import annotations
import csv
import io
import re
import requests
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

from .config import Config
from .google_client import get_credentials, services
from .sheets_utils import (
    delete_sheet, copy_sheet_as, copy_first_sheet_as, refresh_sheets_with_prefix,
    get_value, first_gid
)
from .drive_utils import find_latest_sheet, upload_to_drive
from .gmail_utils import send_email, email_manager_report


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}&gid={gid}"
    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content


def timestamp_now(tz: str, fmt: str) -> str:
    return datetime.now(ZoneInfo(tz)).strftime(fmt)

import re

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _clean_emails(items):
    """
    Accepts a list or a comma-separated string and returns a list of valid emails.
    Trailing commas and blanks are removed. Invalid tokens are dropped silently.
    """
    if items is None:
        return []
    if isinstance(items, str):
        items = [p.strip() for p in items.split(",")]
    return [e for e in (p.strip() for p in items) if e and _EMAIL_RE.match(e)]

def _fallback_recipients(hint, *candidates):
    """
    Return the first non-empty, valid recipient list from the provided candidates.
    If all candidates are empty/invalid, raise a friendly error.
    """
    for c in candidates:
        cleaned = _clean_emails(c)
        if cleaned:
            return cleaned
    # Nothing usable found:
    raise ValueError(
        f"No valid recipients available for: {hint}. "
        f"Please provide at least one email in the UI or .env "
        f"(TO_RECIPIENTS, DEFAULT_ORDER_RECIPIENTS, or per-report-key)."
    )


@dataclass
class RunResult:
    ok: bool
    elapsed_seconds: int
    location: str
    timestamp: str
    manager_pdf_link: str | None
    full_order_link: str | None


def run_pipeline(cfg: Config, logger=None) -> RunResult:
    import time
    start = time.perf_counter()

    if logger:
        logger.info("Authorizing with Google APIs…")
    creds = get_credentials(cfg.SCOPES, cfg.REDIRECT_PORT, cfg.FORCE_REAUTH)
    sheets_svc, drive_svc, gmail_svc = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    if logger:
        logger.info("Google services ready")

    # Step 1: latest incoming
    if logger:
        logger.info("Finding latest incoming spreadsheet…")
    latest = find_latest_sheet(drive_svc, cfg.INCOMING_FOLDER_ID)
    if not latest:
        raise SystemExit("No incoming report found.")
    new_report_id = latest["id"]
    if logger:
        logger.info(f"Latest incoming: {latest['name']} ({new_report_id})")

    # Step 2: prep calculations workbook
    if logger:
        logger.info("Preparing calculations workbook…")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
    try:
        copy_sheet_as(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", "Last Week")
        if logger:
            logger.info("Copied old 'Current Week' to 'Last Week'")
    except Exception:
        if logger:
            logger.warn("No 'Current Week' sheet exists to copy")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
    copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")
    if logger:
        logger.info("Inserted new 'Current Week' from latest incoming report")

    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ')…")
    refresh_sheets_with_prefix(sheets_svc, cfg.CALC_SPREADSHEET_ID, prefix="REFR: ", logger=logger)

    # Step 3: read location code
    location = get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.LOCATION_NAMED_RANGE)
    ts = timestamp_now(cfg.TIMESTAMP_TZ, cfg.TIMESTAMP_FMT)
    if logger:
        logger.info(f"Location: {location}; Timestamp: {ts}")

    # Step 4A: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_MANAGER_PDF, "pdf")
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 4B: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_ORDER_CSV, "csv")

    # Step 4C: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_{ts}_{location}_FULL.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, "text/csv", cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_{ts}_{location}_FULL.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4D: Create per-report-key outputs and email
    text = master_csv_bytes.decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames:
        raise RuntimeError("CSV has no header.")
    report_col = next((h for h in reader.fieldnames if h.lower() == "report_key"), None)
    if not report_col:
        raise RuntimeError("Report_Key column missing.")
    rows = list(reader)
    groups = {}
    for r in rows:
        key = (r.get(report_col) or "").strip() or "UNASSIGNED"
        groups.setdefault(key, []).append(r)

    for key, key_rows in groups.items():
        if not cfg.USE_ALL_REPORT_KEYS and key.upper() not in (cfg.REPORT_KEY_RUN_LIST or []):
            continue
        csv_buf = io.StringIO()
        w = csv.DictWriter(csv_buf, fieldnames=reader.fieldnames, lineterminator="")
        w.writeheader()
        w.writerows(key_rows)
        key_bytes = csv_buf.getvalue().encode("utf-8")
        tag = clean_tag(key.upper())
        filename = f"Order_Report_{ts}_{location}_{tag}.csv"
        created = upload_to_drive(drive_svc, key_bytes, filename, "text/csv", cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{ts}_{location}_{tag}.pdf"

        # Prefer per‑key recipients; else default; else TO
        candidates = None
        if cfg.REPORT_KEY_RECIPIENTS:
            # The keys in UI are upper-cased; normalize here too
            candidates = cfg.REPORT_KEY_RECIPIENTS.get(tag)

        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.DEFAULT_ORDER_RECIPIENTS,
            cfg.TO_RECIPIENTS,
        )

        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – {tag}"
        msg["From"] = "me"
        msg["To"] = ", ".join(recipients)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi {key} team,\nYour order report is ready. \nGoogle Sheet: {created.get('webViewLink')}\nAttached: {pdfname}\n—Automated")
        msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=pdfname)
        if cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL:
            msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info(f"Emailed {tag}")

    # Step 4E: Send Manager Report

    to_list = _fallback_recipients("Manager Report (TO_RECIPIENTS)", cfg.TO_RECIPIENTS)
    cc_list = _clean_emails(cfg.CC_RECIPIENTS)
    email_manager_report(gmail_svc, "me", to_list, cc_list, pdf_name, pdf_bytes, manager_link, ts, location)
    
    if logger:
        logger.info("Manager email sent")

    # Step 4F: Send Full Order if needed
    full_link = full_created.get('webViewLink')
    if cfg.SEND_SEPARATE_FULL_ORDER_EMAIL:
        to_full = _fallback_recipients("FULL order", cfg.DEFAULT_ORDER_RECIPIENTS, cfg.TO_RECIPIENTS)
        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – FULL"
        msg["From"] = "me"
        msg["To"] = ", ".join(to_full)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi team,\nFULL order report is ready.\nSheet: {full_link}\nAttached: {full_pdf_name}\n—Automated")
        msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info("FULL order email sent")
    else:
        if logger:
            logger.info("Separate full order email disabled")

    elapsed = int(time.perf_counter() - start)
    if logger:
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        logger.info(f"Run completed in {h:02d}:{m:02d}:{s:02d}")

    return RunResult(True, elapsed, location, ts, manager_link, full_link)

```

---
### file: favtrip/sheets_utils.py

```python
from __future__ import annotations
import random
import time
from typing import Any, Dict, List


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def delete_sheet(svc, spreadsheet_id: str, title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), title)
    if s:
        body = {"requests": [{"deleteSheet": {"sheetId": s["sheetId"]}}]}
        svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def copy_sheet_as(svc, spreadsheet_id: str, src_title: str, new_title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), src_title)
    if not s:
        return None
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=spreadsheet_id,
        sheetId=s["sheetId"],
        body={"destinationSpreadsheetId": spreadsheet_id}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return new_id


def copy_first_sheet_as(svc, src_spreadsheet: str, dest_spreadsheet: str, new_title: str):
    meta = svc.spreadsheets().get(spreadsheetId=src_spreadsheet).execute()
    first_id = meta["sheets"][0]["properties"]["sheetId"]
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=src_spreadsheet,
        sheetId=first_id,
        body={"destinationSpreadsheetId": dest_spreadsheet}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=dest_spreadsheet, body=body).execute()
    return new_id


def refresh_sheets_with_prefix(svc, spreadsheet_id: str, prefix: str = "REFR: ", retries: int = 5, logger=None):
    sheets = list_sheets(svc, spreadsheet_id)
    targets = [s["properties"] for s in sheets if s["properties"]["title"].startswith(prefix)]
    for idx, t in enumerate(targets, start=1):
        body = {"requests": [{
            "findReplace": {
                "find": "=",
                "replacement": "=",
                "includeFormulas": True,
                "sheetId": t["sheetId"]
            }
        }]}
        attempt = 0
        while True:
            try:
                svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                if logger:
                    logger.info(f"[{idx}/{len(targets)}] Recalc OK: {t['title']}")
                break
            except Exception:
                attempt += 1
                if attempt > retries:
                    if logger:
                        logger.warn(f"FAILED recalc for {t['title']}")
                    break
                time.sleep(1 + random.random())


def get_value(svc, spreadsheet_id: str, sheet_title: str, named_range: str) -> str:
    try:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=named_range
        ).execute().get("values", [])
    except Exception:
        vals = []
    if not vals:
        try:
            vals = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_title}'!A1:A"
            ).execute().get("values", [])
        except Exception:
            vals = []
    return vals[0][0] if vals and vals[0] else "UNKNOWN"


def first_gid(svc, spreadsheet_id: str) -> int:
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return meta["sheets"][0]["properties"]["sheetId"]

```

---
### file: generate_code_bundle.py

```python
#!/usr/bin/env python3
"""
Generate a single-file Markdown snapshot ("CodeBundle") of a project.

Features:
- Directory tree at the top (filtered).
- Each source file is embedded under "### file: <relative path>" with a fenced code block.
- Excludes common noisy/secret paths by default (customizable).
- CLI options for root, output, include/exclude globs, and size limit.

Usage examples:
  python generate_codebundle.py --root ./favtrip_pipeline
  python generate_codebundle.py --root ./favtrip_pipeline --out PROJECT_SNAPSHOT.md
  python generate_codebundle.py --root . --include "**/*.py" --exclude ".venv/**" --max-bytes 2000000
"""

from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    ap.add_argument("--root", default=".", help="Project root folder (default: .)")
    ap.add_argument("--out", default="PROJECT_SNAPSHOT_CODEBUNDLE.md",
                    help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md)")
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print(f"CodeBundle writing starting...")
    args = parse_args()
    root = Path(args.root).resolve()
    out = Path(args.out).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#python generate_codebundle.py --root .\favtrip_pipeline
```

---
### file: last_run.log

```
[2026-03-03 22:23:45] INFO: Authorizing with Google APIs…
[2026-03-03 22:23:45] INFO: Google services ready
[2026-03-03 22:23:45] INFO: Finding latest incoming spreadsheet…
[2026-03-03 22:23:46] INFO: Latest incoming: Testing Sales Report - Week 4 (1RMW3FoLA4PGv_DOAxEK9TcY4PLJvDBzwZcx-GRW-oz4)
[2026-03-03 22:23:46] INFO: Preparing calculations workbook…
[2026-03-03 22:23:55] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 22:24:05] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 22:24:05] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 22:24:41] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 22:24:50] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 22:25:41] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 22:25:43] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-10-25-PM
[2026-03-03 22:25:43] INFO: Exporting Manager Report (PDF)…
[2026-03-03 22:25:46] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/11DFVDMo3RHOq2DpNQ1Vh-asf9kJC8fyF/view?usp=drivesdk
[2026-03-03 22:25:46] INFO: Exporting Master Order (CSV)…
[2026-03-03 22:25:53] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1ZyGyf-bvU7PYeInHiC1NsHZs95a5x7yopURC6FtOnls/edit?usp=drivesdk
[2026-03-03 22:26:00] INFO: Emailed COFFEE
[2026-03-03 22:26:01] INFO: Manager email sent
[2026-03-03 22:26:01] INFO: Separate full order email disabled
[2026-03-03 22:26:01] INFO: Run completed in 00:02:16

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit

```

---
### file: run_web.sh

```bash
#!/usr/bin/env bash
streamlit run ui_streamlit.py

```

---
### file: run_web_windows.bat

```bat
@echo off
streamlit run ui_streamlit.py

```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```

---
### file: ui_streamlit.py

```python
import os
import time
import threading
import streamlit as st

from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.google_client import (
    start_oauth,
    finish_oauth,
    load_valid_token,
    clear_token,
)

def _rerun():
    # Works on Streamlit >= 1.27 (st.rerun) and older (experimental_rerun)
    try:
        import streamlit as st
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

st.set_page_config(page_title="FavTrip Reporting Pipeline", page_icon="🧾", layout="wide")

# --- Larger Run button ---
st.markdown(
    """
    <style>
      div.stButton > button:first-child {
        font-size: 1.1rem;
        padding: 0.6rem 1.2rem;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🧾 FavTrip Reporting Pipeline")

cfg = Config.load()

# ----------------------------
# Session state: auth gating
# ----------------------------
if "auth_checked" not in st.session_state:
    # On first load: if token is missing/invalid/unrefreshable -> require auth
    st.session_state.auth_required = (load_valid_token(cfg.SCOPES) is None)
    st.session_state.oauth_flow = None
    st.session_state.oauth_url = None
    st.session_state.auth_checked = True

# Sidebar controls (always visible)
with st.sidebar:
    st.header("Defaults (.env)")

    if st.button("Force Google Re-Auth", type="secondary", use_container_width=True):
        clear_token()
        try:
            # Prefer auto flow
            from favtrip.google_client import login_via_local_server
            with st.status("Re-auth in progress (browser will open)…", expanded=True):
                creds = login_via_local_server(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.success("✅ Re-auth complete.")
            st.session_state.auth_required = False
            st.rerun()
        except Exception as e:
            # Fallback to manual method
            try:
                flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.session_state.oauth_flow = flow
                st.session_state.oauth_url = url
                st.session_state.auth_required = True  # show the auth panel
                st.info("Auto re-auth failed; showing manual method. Open the URL shown in the Authentication panel.")
                st.rerun()
            except Exception as e2:
                st.error(f"Failed to start re-auth: {e2}")

    # Developer mode (suppresses console prints when unchecked)
    dev_mode = st.checkbox("Run in developer mode", value=False)

    st.markdown("Edit values for this *one run*. Optionally tick **Update .env** to persist.")

# ----------------------------
# Authentication panel (shown only if auth required)
# ----------------------------
if st.session_state.auth_required:
    from favtrip.google_client import login_via_local_server  # import here to avoid circulars

    with st.expander("Google Authentication", expanded=True):
        st.caption(
            "Authentication is required before running. "
            "Click **Start authentication** to open the browser. "
            "We will wait for the redirect automatically."
        )

        # Preferred: automatic (open browser + capture redirect)
        if st.button("Start authentication (auto-open & capture)", type="primary"):
            try:
                with st.status("Waiting for Google authorization in your browser…", expanded=True):
                    creds = login_via_local_server(cfg.SCOPES, cfg.REDIRECT_PORT)
                    st.success("✅ Authentication complete. token.json saved.")
                st.session_state.oauth_flow = None
                st.session_state.oauth_url = None
                st.session_state.auth_required = False
                st.rerun()
            except Exception as e:
                st.error(f"Auto authentication failed: {e}. You can try the manual method below.")

        st.divider()
        st.write("**Manual method (fallback):**")

        # Manual fallback (existing behavior)
        col_a, col_b = st.columns([2, 1])
        with col_a:
            if st.button("Start authentication (get URL)"):
                try:
                    flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                    st.session_state.oauth_flow = flow
                    st.session_state.oauth_url = url
                    st.success("Auth URL generated below. Open it, grant access, and paste the redirect URL or code.")
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")
        with col_b:
            if st.session_state.oauth_url:
                st.link_button("Open Auth URL", st.session_state.oauth_url, use_container_width=True)

        if st.session_state.oauth_url:
            st.code(st.session_state.oauth_url, language="text")

        pasted = st.text_input(
            "Paste full redirect URL or the code here",
            value="",
            placeholder="https://... or the code",
        )

        if st.button("Complete authentication", type="secondary"):
            flow = st.session_state.get("oauth_flow")
            if not flow:
                st.warning("Click 'Start authentication (get URL)' first.")
            else:
                try:
                    creds = finish_oauth(flow, pasted)
                    st.session_state.oauth_flow = None
                    st.session_state.oauth_url = None
                    st.session_state.auth_required = False
                    st.success("✅ Authentication complete. token.json saved.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to finish OAuth: {e}")

# ----------------------------
# Only show Run Options if NOT requiring auth
# ----------------------------
if not st.session_state.auth_required:

    # ---- Run Form (Run button top-right) ----
    with st.form("run_form"):
        tl, tr = st.columns([4, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")
        with tr:
            submitted = st.form_submit_button("▶️ Run Pipeline", use_container_width=True)

        # --- Main options ---
        colA, colB, colC = st.columns(3)
        with colA:
            to = st.text_input("To Recipients (comma)", value=",".join(cfg.TO_RECIPIENTS or []))
            cc = st.text_input("CC Recipients (comma)", value=",".join(cfg.CC_RECIPIENTS or []))
        with colB:
            use_all = st.checkbox("Use all Report Keys in CSV", value=cfg.USE_ALL_REPORT_KEYS)
            report_keys = st.text_input("Report Keys to run (comma)", value=",".join(cfg.REPORT_KEY_RUN_LIST or []))
        with colC:
            include_full = st.checkbox("Attach FULL order in each email", value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL)
            send_full = st.checkbox("Send separate FULL order email", value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL)
            #force_reauth = st.checkbox("Force Google re-auth for this run", value=cfg.FORCE_REAUTH)

        # --- Per-report-key recipients editor (above Advanced) ---
        with st.expander("Per-Report-Key Recipients (optional)"):
            st.caption('Map **REPORT KEY (ALL CAPS)** → **Emails (comma)** (friendly editor for `REPORT_KEY_RECIPIENTS`).')
            rows = []
            if cfg.REPORT_KEY_RECIPIENTS:
                for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({"REPORT KEY (ALL CAPS)": k, "Emails (comma)": ",".join(v or [])})
            else:
                rows = [{"REPORT KEY (ALL CAPS)": "", "Emails (comma)": ""}]
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                use_container_width=True,
                key="rk_editor",
            )

        # --- Advanced (IDs, GIDs, Timezone, Redirect Port) ---
        with st.expander("Advanced (IDs, GIDs, Timezone, Redirect Port)"):
            col1, col2 = st.columns(2)
            with col1:
                calc_id = st.text_input("Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID)
                incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID)
                mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID)
                order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID)
                
                raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                redirect_port = st.number_input(
                    "Redirect Port (0 = auto)",
                    min_value=0,             # allow 0 explicitly
                    max_value=65535,
                    value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                    help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535.",
                )

            with col2:
                gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF))
                gid_order = st.text_input("Order CSV gid", value=str(cfg.GID_ORDER_CSV))
                loc_sheet = st.text_input("Location Sheet Title", value=cfg.LOCATION_SHEET_TITLE)
                loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE)
                tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ)
                tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT)

        save_env = st.checkbox("Update defaults in .env with the edited fields (optional)")

    # ----------------------------
    # Submission handling
    # ----------------------------
    def _split_emails(csv_str: str):
        return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

    if submitted:
        # Apply per-run config
        cfg.TO_RECIPIENTS = _split_emails(to)
        cfg.CC_RECIPIENTS = _split_emails(cc)
        cfg.USE_ALL_REPORT_KEYS = use_all
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

        cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
        cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full

        cfg.CALC_SPREADSHEET_ID = calc_id
        cfg.INCOMING_FOLDER_ID = incoming_id
        cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
        cfg.ORDER_REPORT_FOLDER_ID = order_folder
        cfg.REDIRECT_PORT = int(redirect_port)

        cfg.GID_MANAGER_PDF = gid_mgr
        cfg.GID_ORDER_CSV = gid_order
        cfg.LOCATION_SHEET_TITLE = loc_sheet
        cfg.LOCATION_NAMED_RANGE = loc_range
        cfg.TIMESTAMP_TZ = tz
        cfg.TIMESTAMP_FMT = tfmt

        # Per-key recipients from editor
        rk_map = {}
        for r in edited_rows:
            key = (r.get("REPORT KEY (ALL CAPS)") or "").strip().upper()
            emails_csv = r.get("Emails (comma)") or ""
            emails = _split_emails(emails_csv)
            if key and emails:
                rk_map[key] = emails
        cfg.REPORT_KEY_RECIPIENTS = rk_map

        if save_env:
            cfg.save()
            st.success("Saved updated defaults to .env")

        # If user checked "Force Google re-auth for this run", kick them into auth gating first.
        if cfg.FORCE_REAUTH:
            clear_token()
            try:
                flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.session_state.oauth_flow = flow
                st.session_state.oauth_url = url
                st.session_state.auth_required = True
                st.info("Re-auth required for this run. Open the URL shown in the Authentication panel.")
                _rerun()
            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")
        else:
            # --- Live run with timer + last log (no full log after completion) ---
            # Write all logs to last_run.log; overwrite on each run
            logger = StatusLogger(print_to_console=dev_mode, file_path="last_run.log", overwrite=True)
            result_holder = {"value": None, "error": None}

            def _runner():
                try:
                    result_holder["value"] = run_pipeline(cfg, logger=logger)
                except Exception as e:
                    result_holder["error"] = e

            t0 = time.perf_counter()
            th = threading.Thread(target=_runner, daemon=True)
            th.start()

            with st.status("Running pipeline…", expanded=True) as status:
                timer_ph = st.empty()
                lastlog_ph = st.empty()

                while th.is_alive():
                    elapsed = int(time.perf_counter() - t0)
                    timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                    lastlog_ph.markdown(f"**Last:** {logger.last_line()}")
                    time.sleep(0.5)

                th.join()
                elapsed = int(time.perf_counter() - t0)
                timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                lastlog_ph.markdown(f"**Last:** {logger.last_line()}")

                if result_holder["error"]:
                    st.error(f"Run failed: {result_holder['error']}")
                    status.update(label="❌ Failed", state="error")
                else:
                    result = result_holder["value"]
                    st.write("### Outputs")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Location", result.location)
                    col2.metric("Timestamp", result.timestamp)
                    mm = result.elapsed_seconds
                    col3.metric("Elapsed", f"{mm//3600:02d}:{(mm%3600)//60:02d}:{mm%60:02d}")
                    if result.manager_pdf_link:
                        st.success(f"Manager PDF: {result.manager_pdf_link}")
                    if result.full_order_link:
                        st.success(f"Full Order Sheet: {result.full_order_link}")
                    # Note: intentionally NOT showing the full live log post-run
                    status.update(label="✅ Completed", state="complete")
```

```

---
### file: README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: generate_code_bundle.py

```python
#!/usr/bin/env python3
"""
Generate a single-file Markdown snapshot ("CodeBundle") of a project.

Features:
- Directory tree at the top (filtered).
- Each source file is embedded under "### file: <relative path>" with a fenced code block.
- Excludes common noisy/secret paths by default (customizable).
- CLI options for root, output, include/exclude globs, and size limit.

Usage examples:
  python generate_codebundle.py --root ./favtrip_pipeline
  python generate_codebundle.py --root ./favtrip_pipeline --out PROJECT_SNAPSHOT.md
  python generate_codebundle.py --root . --include "**/*.py" --exclude ".venv/**" --max-bytes 2000000
"""

from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    ap.add_argument("--root", default=".", help="Project root folder (default: .)")
    ap.add_argument("--out", default="PROJECT_SNAPSHOT_CODEBUNDLE.md",
                    help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md)")
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print(f"CodeBundle writing starting...")
    args = parse_args()
    root = Path(args.root).resolve()
    out = Path(args.out).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd C:\Users\rjrul\Downloads\favtrip_pipeline\documentation\
#python generate_codebundle.py --root .\
```

---
### file: requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit

```

```

---
### file: README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )

    # NOTE: We won't hardcode the default here; we'll inject it in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default="PROJECT_SNAPSHOT_CODEBUNDLE.md",
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")

    args = parse_args()

    # NEW: compute project root as the parent folder of this script
    script_dir = Path(__file__).resolve().parent
    project_root = script_dir.parent

    # If --root was not specified, default to the script's parent folder
    if args.root is None:
        root = project_root
    else:
        # If user gave a relative path, resolve it from the *current* working dir
        # (You could also resolve it from project_root; see Option B below)
        root = Path(args.root).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # Default out goes under the root unless an absolute path was given
    out = Path(args.out)
    if not out.is_absolute():
        out = (root / out).resolve()

    # existing ignore logic unchanged …
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd C:\Users\rjrul\Downloads\favtrip_pipeline\documentation\
#python generate_code_bundle.py --root .\
```

---
### file: requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit

```

```

---
### file: README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )

    # NOTE: We won't hardcode the default here; we'll inject it in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default="PROJECT_SNAPSHOT_CODEBUNDLE.md",
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")

    args = parse_args()

    # NEW: compute project root as the parent folder of this script
    script_dir = Path(__file__).resolve().parent
    project_root = script_dir.parent

    # If --root was not specified, default to the script's parent folder
    if args.root is None:
        root = project_root
    else:
        # If user gave a relative path, resolve it from the *current* working dir
        # (You could also resolve it from project_root; see Option B below)
        root = Path(args.root).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # Default out goes under the root unless an absolute path was given
    out = Path(args.out)
    if not out.is_absolute():
        out = (root / out).resolve()

    # existing ignore logic unchanged …
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd C:\Users\rjrul\Downloads\favtrip_pipeline\documentation\
#python generate_code_bundle.py --root .\
```

---
### file: requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit

```

```

---
### file: documentation/README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: documentation/generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd C:\Users\rjrul\Downloads\favtrip_pipeline\documentation\
#python generate_code_bundle.py
```

---
### file: documentation/requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit

```

---
### file: favtrip/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: favtrip/config.py

```python
from __future__ import annotations
import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv

_BOOL_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _b(s: str, default: bool = False) -> bool:
    if s is None:
        return default
    return str(s).strip().lower() in _BOOL_TRUE


def _csv(s: str) -> List[str]:
    if not s:
        return []
    return [p.strip() for p in s.split(",") if p.strip()]


def _json_or_empty(s: str):
    if not s:
        return {}
    try:
        return json.loads(s)
    except Exception:
        return {}


@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"

    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"

    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None

    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None

    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True

    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285

    HTTP_TIMEOUT_SECONDS: int = 300

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)
        return Config(
            CALC_SPREADSHEET_ID=os.getenv("CALC_SPREADSHEET_ID", ""),
            INCOMING_FOLDER_ID=os.getenv("INCOMING_FOLDER_ID", ""),
            MANAGER_REPORT_FOLDER_ID=os.getenv("MANAGER_REPORT_FOLDER_ID", ""),
            ORDER_REPORT_FOLDER_ID=os.getenv("ORDER_REPORT_FOLDER_ID", ""),
            GID_MANAGER_PDF=os.getenv("GID_MANAGER_PDF", "1921812573"),
            GID_ORDER_CSV=os.getenv("GID_ORDER_CSV", "1875928148"),
            LOCATION_SHEET_TITLE=os.getenv("LOCATION_SHEET_TITLE", "REFR: Values"),
            LOCATION_NAMED_RANGE=os.getenv("LOCATION_NAMED_RANGE", "_locations"),
            TIMESTAMP_TZ=os.getenv("TIMESTAMP_TZ", "America/Chicago"),
            TIMESTAMP_FMT=os.getenv("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p"),
            TO_RECIPIENTS=_csv(os.getenv("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_csv(os.getenv("CC_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_b(os.getenv("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=_csv(os.getenv("REPORT_KEY_RUN_LIST", "")),
            REPORT_KEY_RECIPIENTS=_json_or_empty(os.getenv("REPORT_KEY_RECIPIENTS", "")),
            DEFAULT_ORDER_RECIPIENTS=_csv(os.getenv("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_b(
                os.getenv("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_b(
                os.getenv("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),
            SCOPES=_csv(
                os.getenv(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_b(os.getenv("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(os.getenv("REDIRECT_PORT", "58285")),
            HTTP_TIMEOUT_SECONDS=int(os.getenv("HTTP_TIMEOUT_SECONDS", "300")),
        )

    def to_env(self) -> str:
        # Serialize to .env format (simple)
        data = asdict(self)
        # Convert list and dict fields
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")
```

---
### file: favtrip/drive_utils.py

```python
from __future__ import annotations
import io
from googleapiclient.http import MediaIoBaseUpload


def find_latest_sheet(drive_svc, folder_id: str):
    q = (
        f"'{folder_id}' in parents and "
        "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    )
    resp = drive_svc.files().list(
        q=q, orderBy="createdTime desc", pageSize=1,
        fields="files(id,name,createdTime)"
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def upload_to_drive(drive_svc, data: bytes, name: str, mime: str, folder_id: str, to_sheet: bool=False):
    meta = {"name": name, "parents": [folder_id]}
    if to_sheet:
        meta["mimeType"] = "application/vnd.google-apps.spreadsheet"
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
    return drive_svc.files().create(
        body=meta, media_body=media, fields="id,name,mimeType,webViewLink"
    ).execute()

```

---
### file: favtrip/gmail_utils.py

```python
from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {ts} – {location}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Automated")
    msg.add_alternative(
        f"<p>Hi team,</p><p>Manager Report ({location})</p>"
        f"<a href='{pdf_link}'>Backup Link</a>", subtype="html"
    )
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)

```

---
### file: favtrip/google_client.py

```python
from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail
```

---
### file: favtrip/logger.py

```python
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass
```

---
### file: favtrip/pipeline.py

```python
from __future__ import annotations
import csv
import io
import re
import requests
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

from .config import Config
from .google_client import get_credentials, services
from .sheets_utils import (
    delete_sheet, copy_sheet_as, copy_first_sheet_as, refresh_sheets_with_prefix,
    get_value, first_gid
)
from .drive_utils import find_latest_sheet, upload_to_drive
from .gmail_utils import send_email, email_manager_report


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}&gid={gid}"
    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content


def timestamp_now(tz: str, fmt: str) -> str:
    return datetime.now(ZoneInfo(tz)).strftime(fmt)

import re

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _clean_emails(items):
    """
    Accepts a list or a comma-separated string and returns a list of valid emails.
    Trailing commas and blanks are removed. Invalid tokens are dropped silently.
    """
    if items is None:
        return []
    if isinstance(items, str):
        items = [p.strip() for p in items.split(",")]
    return [e for e in (p.strip() for p in items) if e and _EMAIL_RE.match(e)]

def _fallback_recipients(hint, *candidates):
    """
    Return the first non-empty, valid recipient list from the provided candidates.
    If all candidates are empty/invalid, raise a friendly error.
    """
    for c in candidates:
        cleaned = _clean_emails(c)
        if cleaned:
            return cleaned
    # Nothing usable found:
    raise ValueError(
        f"No valid recipients available for: {hint}. "
        f"Please provide at least one email in the UI or .env "
        f"(TO_RECIPIENTS, DEFAULT_ORDER_RECIPIENTS, or per-report-key)."
    )


@dataclass
class RunResult:
    ok: bool
    elapsed_seconds: int
    location: str
    timestamp: str
    manager_pdf_link: str | None
    full_order_link: str | None


def run_pipeline(cfg: Config, logger=None) -> RunResult:
    import time
    start = time.perf_counter()

    if logger:
        logger.info("Authorizing with Google APIs…")
    creds = get_credentials(cfg.SCOPES, cfg.REDIRECT_PORT, cfg.FORCE_REAUTH)
    sheets_svc, drive_svc, gmail_svc = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    if logger:
        logger.info("Google services ready")

    # Step 1: latest incoming
    if logger:
        logger.info("Finding latest incoming spreadsheet…")
    latest = find_latest_sheet(drive_svc, cfg.INCOMING_FOLDER_ID)
    if not latest:
        raise SystemExit("No incoming report found.")
    new_report_id = latest["id"]
    if logger:
        logger.info(f"Latest incoming: {latest['name']} ({new_report_id})")

    # Step 2: prep calculations workbook
    if logger:
        logger.info("Preparing calculations workbook…")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
    try:
        copy_sheet_as(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", "Last Week")
        if logger:
            logger.info("Copied old 'Current Week' to 'Last Week'")
    except Exception:
        if logger:
            logger.warn("No 'Current Week' sheet exists to copy")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
    copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")
    if logger:
        logger.info("Inserted new 'Current Week' from latest incoming report")

    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ')…")
    refresh_sheets_with_prefix(sheets_svc, cfg.CALC_SPREADSHEET_ID, prefix="REFR: ", logger=logger)

    # Step 3: read location code
    location = get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.LOCATION_NAMED_RANGE)
    ts = timestamp_now(cfg.TIMESTAMP_TZ, cfg.TIMESTAMP_FMT)
    if logger:
        logger.info(f"Location: {location}; Timestamp: {ts}")

    # Step 4A: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_MANAGER_PDF, "pdf")
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 4B: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_ORDER_CSV, "csv")

    # Step 4C: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_{ts}_{location}_FULL.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, "text/csv", cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_{ts}_{location}_FULL.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4D: Create per-report-key outputs and email
    text = master_csv_bytes.decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames:
        raise RuntimeError("CSV has no header.")
    report_col = next((h for h in reader.fieldnames if h.lower() == "report_key"), None)
    if not report_col:
        raise RuntimeError("Report_Key column missing.")
    rows = list(reader)
    groups = {}
    for r in rows:
        key = (r.get(report_col) or "").strip() or "UNASSIGNED"
        groups.setdefault(key, []).append(r)

    for key, key_rows in groups.items():
        if not cfg.USE_ALL_REPORT_KEYS and key.upper() not in (cfg.REPORT_KEY_RUN_LIST or []):
            continue
        csv_buf = io.StringIO()
        w = csv.DictWriter(csv_buf, fieldnames=reader.fieldnames, lineterminator="")
        w.writeheader()
        w.writerows(key_rows)
        key_bytes = csv_buf.getvalue().encode("utf-8")
        tag = clean_tag(key.upper())
        filename = f"Order_Report_{ts}_{location}_{tag}.csv"
        created = upload_to_drive(drive_svc, key_bytes, filename, "text/csv", cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{ts}_{location}_{tag}.pdf"

        # Prefer per‑key recipients; else default; else TO
        candidates = None
        if cfg.REPORT_KEY_RECIPIENTS:
            # The keys in UI are upper-cased; normalize here too
            candidates = cfg.REPORT_KEY_RECIPIENTS.get(tag)

        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.DEFAULT_ORDER_RECIPIENTS,
            cfg.TO_RECIPIENTS,
        )

        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – {tag}"
        msg["From"] = "me"
        msg["To"] = ", ".join(recipients)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi {key} team,\nYour order report is ready. \nGoogle Sheet: {created.get('webViewLink')}\nAttached: {pdfname}\n—Automated")
        msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=pdfname)
        if cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL:
            msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info(f"Emailed {tag}")

    # Step 4E: Send Manager Report

    to_list = _fallback_recipients("Manager Report (TO_RECIPIENTS)", cfg.TO_RECIPIENTS)
    cc_list = _clean_emails(cfg.CC_RECIPIENTS)
    email_manager_report(gmail_svc, "me", to_list, cc_list, pdf_name, pdf_bytes, manager_link, ts, location)
    
    if logger:
        logger.info("Manager email sent")

    # Step 4F: Send Full Order if needed
    full_link = full_created.get('webViewLink')
    if cfg.SEND_SEPARATE_FULL_ORDER_EMAIL:
        to_full = _fallback_recipients("FULL order", cfg.DEFAULT_ORDER_RECIPIENTS, cfg.TO_RECIPIENTS)
        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – FULL"
        msg["From"] = "me"
        msg["To"] = ", ".join(to_full)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi team,\nFULL order report is ready.\nSheet: {full_link}\nAttached: {full_pdf_name}\n—Automated")
        msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info("FULL order email sent")
    else:
        if logger:
            logger.info("Separate full order email disabled")

    elapsed = int(time.perf_counter() - start)
    if logger:
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        logger.info(f"Run completed in {h:02d}:{m:02d}:{s:02d}")

    return RunResult(True, elapsed, location, ts, manager_link, full_link)

```

---
### file: favtrip/sheets_utils.py

```python
from __future__ import annotations
import random
import time
from typing import Any, Dict, List


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def delete_sheet(svc, spreadsheet_id: str, title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), title)
    if s:
        body = {"requests": [{"deleteSheet": {"sheetId": s["sheetId"]}}]}
        svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def copy_sheet_as(svc, spreadsheet_id: str, src_title: str, new_title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), src_title)
    if not s:
        return None
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=spreadsheet_id,
        sheetId=s["sheetId"],
        body={"destinationSpreadsheetId": spreadsheet_id}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return new_id


def copy_first_sheet_as(svc, src_spreadsheet: str, dest_spreadsheet: str, new_title: str):
    meta = svc.spreadsheets().get(spreadsheetId=src_spreadsheet).execute()
    first_id = meta["sheets"][0]["properties"]["sheetId"]
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=src_spreadsheet,
        sheetId=first_id,
        body={"destinationSpreadsheetId": dest_spreadsheet}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=dest_spreadsheet, body=body).execute()
    return new_id


def refresh_sheets_with_prefix(svc, spreadsheet_id: str, prefix: str = "REFR: ", retries: int = 5, logger=None):
    sheets = list_sheets(svc, spreadsheet_id)
    targets = [s["properties"] for s in sheets if s["properties"]["title"].startswith(prefix)]
    for idx, t in enumerate(targets, start=1):
        body = {"requests": [{
            "findReplace": {
                "find": "=",
                "replacement": "=",
                "includeFormulas": True,
                "sheetId": t["sheetId"]
            }
        }]}
        attempt = 0
        while True:
            try:
                svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                if logger:
                    logger.info(f"[{idx}/{len(targets)}] Recalc OK: {t['title']}")
                break
            except Exception:
                attempt += 1
                if attempt > retries:
                    if logger:
                        logger.warn(f"FAILED recalc for {t['title']}")
                    break
                time.sleep(1 + random.random())


def get_value(svc, spreadsheet_id: str, sheet_title: str, named_range: str) -> str:
    try:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=named_range
        ).execute().get("values", [])
    except Exception:
        vals = []
    if not vals:
        try:
            vals = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_title}'!A1:A"
            ).execute().get("values", [])
        except Exception:
            vals = []
    return vals[0][0] if vals and vals[0] else "UNKNOWN"


def first_gid(svc, spreadsheet_id: str) -> int:
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return meta["sheets"][0]["properties"]["sheetId"]

```

---
### file: last_run.log

```
[2026-03-03 23:16:33] INFO: Authorizing with Google APIs…
[2026-03-03 23:16:33] INFO: Google services ready
[2026-03-03 23:16:33] INFO: Finding latest incoming spreadsheet…
[2026-03-03 23:16:34] INFO: Latest incoming: Testing Sales Report - Week 5 (1kNjmEbljdUIqJUjwh8e2-plfLWRZSGHxMHxL50-2ce4)
[2026-03-03 23:16:34] INFO: Preparing calculations workbook…
[2026-03-03 23:16:42] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 23:16:50] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 23:16:50] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 23:17:17] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 23:17:22] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 23:18:15] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 23:18:17] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-11-18-PM
[2026-03-03 23:18:17] INFO: Exporting Manager Report (PDF)…
[2026-03-03 23:18:21] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/1UILYk_poamIwe6JrFQUGZIz9az9iBsdU/view?usp=drivesdk
[2026-03-03 23:18:21] INFO: Exporting Master Order (CSV)…
[2026-03-03 23:18:34] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1bOcle5eMseFc43TZMnXLqgHR7_DcPs13ntb6PSBians/edit?usp=drivesdk
[2026-03-03 23:18:42] INFO: Emailed COFFEE
[2026-03-03 23:18:43] INFO: Manager email sent
[2026-03-03 23:18:43] INFO: Separate full order email disabled
[2026-03-03 23:18:43] INFO: Run completed in 00:02:09

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```

---
### file: ui_streamlit.py

```python
import os
import time
import threading
import streamlit as st

from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.google_client import (
    start_oauth,
    finish_oauth,
    load_valid_token,
    clear_token,
)

def _rerun():
    # Works on Streamlit >= 1.27 (st.rerun) and older (experimental_rerun)
    try:
        import streamlit as st
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

st.set_page_config(page_title="FavTrip Reporting Pipeline", page_icon="🧾", layout="wide")

# --- Larger Run button ---
st.markdown(
    """
    <style>
      div.stButton > button:first-child {
        font-size: 1.1rem;
        padding: 0.6rem 1.2rem;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🧾 FavTrip Reporting Pipeline")

cfg = Config.load()

# ----------------------------
# Session state: auth gating
# ----------------------------
if "auth_checked" not in st.session_state:
    # On first load: if token is missing/invalid/unrefreshable -> require auth
    st.session_state.auth_required = (load_valid_token(cfg.SCOPES) is None)
    st.session_state.oauth_flow = None
    st.session_state.oauth_url = None
    st.session_state.auth_checked = True

# Sidebar controls (always visible)
with st.sidebar:
    st.header("Defaults (.env)")

    if st.button("Force Google Re-Auth", type="secondary", use_container_width=True):
        clear_token()
        try:
            # Prefer auto flow
            from favtrip.google_client import login_via_local_server
            with st.status("Re-auth in progress (browser will open)…", expanded=True):
                creds = login_via_local_server(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.success("✅ Re-auth complete.")
            st.session_state.auth_required = False
            st.rerun()
        except Exception as e:
            # Fallback to manual method
            try:
                flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.session_state.oauth_flow = flow
                st.session_state.oauth_url = url
                st.session_state.auth_required = True  # show the auth panel
                st.info("Auto re-auth failed; showing manual method. Open the URL shown in the Authentication panel.")
                st.rerun()
            except Exception as e2:
                st.error(f"Failed to start re-auth: {e2}")

    # Developer mode (suppresses console prints when unchecked)
    dev_mode = st.checkbox("Run in developer mode", value=False)

    st.markdown("Edit values for this *one run*. Optionally tick **Update .env** to persist.")

# ----------------------------
# Authentication panel (shown only if auth required)
# ----------------------------
if st.session_state.auth_required:
    from favtrip.google_client import login_via_local_server  # import here to avoid circulars

    with st.expander("Google Authentication", expanded=True):
        st.caption(
            "Authentication is required before running. "
            "Click **Start authentication** to open the browser. "
            "We will wait for the redirect automatically."
        )

        # Preferred: automatic (open browser + capture redirect)
        if st.button("Start authentication (auto-open & capture)", type="primary"):
            try:
                with st.status("Waiting for Google authorization in your browser…", expanded=True):
                    creds = login_via_local_server(cfg.SCOPES, cfg.REDIRECT_PORT)
                    st.success("✅ Authentication complete. token.json saved.")
                st.session_state.oauth_flow = None
                st.session_state.oauth_url = None
                st.session_state.auth_required = False
                st.rerun()
            except Exception as e:
                st.error(f"Auto authentication failed: {e}. You can try the manual method below.")

        st.divider()
        st.write("**Manual method (fallback):**")

        # Manual fallback (existing behavior)
        col_a, col_b = st.columns([2, 1])
        with col_a:
            if st.button("Start authentication (get URL)"):
                try:
                    flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                    st.session_state.oauth_flow = flow
                    st.session_state.oauth_url = url
                    st.success("Auth URL generated below. Open it, grant access, and paste the redirect URL or code.")
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")
        with col_b:
            if st.session_state.oauth_url:
                st.link_button("Open Auth URL", st.session_state.oauth_url, use_container_width=True)

        if st.session_state.oauth_url:
            st.code(st.session_state.oauth_url, language="text")

        pasted = st.text_input(
            "Paste full redirect URL or the code here",
            value="",
            placeholder="https://... or the code",
        )

        if st.button("Complete authentication", type="secondary"):
            flow = st.session_state.get("oauth_flow")
            if not flow:
                st.warning("Click 'Start authentication (get URL)' first.")
            else:
                try:
                    creds = finish_oauth(flow, pasted)
                    st.session_state.oauth_flow = None
                    st.session_state.oauth_url = None
                    st.session_state.auth_required = False
                    st.success("✅ Authentication complete. token.json saved.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to finish OAuth: {e}")

# ----------------------------
# Only show Run Options if NOT requiring auth
# ----------------------------
if not st.session_state.auth_required:

    # ---- Run Form (Run button top-right) ----
    with st.form("run_form"):
        tl, tr = st.columns([4, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")
        with tr:
            submitted = st.form_submit_button("▶️ Run Pipeline", use_container_width=True)

        # --- Main options ---
        colA, colB, colC = st.columns(3)
        with colA:
            to = st.text_input("To Recipients (comma)", value=",".join(cfg.TO_RECIPIENTS or []))
            cc = st.text_input("CC Recipients (comma)", value=",".join(cfg.CC_RECIPIENTS or []))
        with colB:
            use_all = st.checkbox("Use all Report Keys in CSV", value=cfg.USE_ALL_REPORT_KEYS)
            report_keys = st.text_input("Report Keys to run (comma)", value=",".join(cfg.REPORT_KEY_RUN_LIST or []))
        with colC:
            include_full = st.checkbox("Attach FULL order in each email", value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL)
            send_full = st.checkbox("Send separate FULL order email", value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL)
            #force_reauth = st.checkbox("Force Google re-auth for this run", value=cfg.FORCE_REAUTH)

        # --- Per-report-key recipients editor (above Advanced) ---
        with st.expander("Per-Report-Key Recipients (optional)"):
            st.caption('Map **REPORT KEY (ALL CAPS)** → **Emails (comma)** (friendly editor for `REPORT_KEY_RECIPIENTS`).')
            rows = []
            if cfg.REPORT_KEY_RECIPIENTS:
                for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({"REPORT KEY (ALL CAPS)": k, "Emails (comma)": ",".join(v or [])})
            else:
                rows = [{"REPORT KEY (ALL CAPS)": "", "Emails (comma)": ""}]
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                use_container_width=True,
                key="rk_editor",
            )

        # --- Advanced (IDs, GIDs, Timezone, Redirect Port) ---
        with st.expander("Advanced (IDs, GIDs, Timezone, Redirect Port)"):
            col1, col2 = st.columns(2)
            with col1:
                calc_id = st.text_input("Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID)
                incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID)
                mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID)
                order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID)
                
                raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                redirect_port = st.number_input(
                    "Redirect Port (0 = auto)",
                    min_value=0,             # allow 0 explicitly
                    max_value=65535,
                    value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                    help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535.",
                )

            with col2:
                gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF))
                gid_order = st.text_input("Order CSV gid", value=str(cfg.GID_ORDER_CSV))
                loc_sheet = st.text_input("Location Sheet Title", value=cfg.LOCATION_SHEET_TITLE)
                loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE)
                tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ)
                tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT)

        save_env = st.checkbox("Update defaults in .env with the edited fields (optional)")

    # ----------------------------
    # Submission handling
    # ----------------------------
    def _split_emails(csv_str: str):
        return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

    if submitted:
        # Apply per-run config
        cfg.TO_RECIPIENTS = _split_emails(to)
        cfg.CC_RECIPIENTS = _split_emails(cc)
        cfg.USE_ALL_REPORT_KEYS = use_all
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

        cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
        cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full

        cfg.CALC_SPREADSHEET_ID = calc_id
        cfg.INCOMING_FOLDER_ID = incoming_id
        cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
        cfg.ORDER_REPORT_FOLDER_ID = order_folder
        cfg.REDIRECT_PORT = int(redirect_port)

        cfg.GID_MANAGER_PDF = gid_mgr
        cfg.GID_ORDER_CSV = gid_order
        cfg.LOCATION_SHEET_TITLE = loc_sheet
        cfg.LOCATION_NAMED_RANGE = loc_range
        cfg.TIMESTAMP_TZ = tz
        cfg.TIMESTAMP_FMT = tfmt

        # Per-key recipients from editor
        rk_map = {}
        for r in edited_rows:
            key = (r.get("REPORT KEY (ALL CAPS)") or "").strip().upper()
            emails_csv = r.get("Emails (comma)") or ""
            emails = _split_emails(emails_csv)
            if key and emails:
                rk_map[key] = emails
        cfg.REPORT_KEY_RECIPIENTS = rk_map

        if save_env:
            cfg.save()
            st.success("Saved updated defaults to .env")

        # If user checked "Force Google re-auth for this run", kick them into auth gating first.
        if cfg.FORCE_REAUTH:
            clear_token()
            try:
                flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.session_state.oauth_flow = flow
                st.session_state.oauth_url = url
                st.session_state.auth_required = True
                st.info("Re-auth required for this run. Open the URL shown in the Authentication panel.")
                _rerun()
            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")
        else:
            # --- Live run with timer + last log (no full log after completion) ---
            # Write all logs to last_run.log; overwrite on each run
            logger = StatusLogger(print_to_console=dev_mode, file_path="last_run.log", overwrite=True)
            result_holder = {"value": None, "error": None}

            def _runner():
                try:
                    result_holder["value"] = run_pipeline(cfg, logger=logger)
                except Exception as e:
                    result_holder["error"] = e

            t0 = time.perf_counter()
            th = threading.Thread(target=_runner, daemon=True)
            th.start()

            with st.status("Running pipeline…", expanded=True) as status:
                timer_ph = st.empty()
                lastlog_ph = st.empty()

                while th.is_alive():
                    elapsed = int(time.perf_counter() - t0)
                    timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                    lastlog_ph.markdown(f"**Last:** {logger.last_line()}")
                    time.sleep(0.5)

                th.join()
                elapsed = int(time.perf_counter() - t0)
                timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                lastlog_ph.markdown(f"**Last:** {logger.last_line()}")

                if result_holder["error"]:
                    st.error(f"Run failed: {result_holder['error']}")
                    status.update(label="❌ Failed", state="error")
                else:
                    result = result_holder["value"]
                    st.write("### Outputs")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Location", result.location)
                    col2.metric("Timestamp", result.timestamp)
                    mm = result.elapsed_seconds
                    col3.metric("Elapsed", f"{mm//3600:02d}:{(mm%3600)//60:02d}:{mm%60:02d}")
                    if result.manager_pdf_link:
                        st.success(f"Manager PDF: {result.manager_pdf_link}")
                    if result.full_order_link:
                        st.success(f"Full Order Sheet: {result.full_order_link}")
                    # Note: intentionally NOT showing the full live log post-run
                    status.update(label="✅ Completed", state="complete")
```

```

---
### file: documentation/README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: documentation/generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd C:\Users\rjrul\Downloads\favtrip_pipeline\documentation\
#python generate_code_bundle.py
```

---
### file: documentation/requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit

```

---
### file: favtrip/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: favtrip/config.py

```python
from __future__ import annotations
import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv

_BOOL_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _b(s: str, default: bool = False) -> bool:
    if s is None:
        return default
    return str(s).strip().lower() in _BOOL_TRUE


def _csv(s: str) -> List[str]:
    if not s:
        return []
    return [p.strip() for p in s.split(",") if p.strip()]


def _json_or_empty(s: str):
    if not s:
        return {}
    try:
        return json.loads(s)
    except Exception:
        return {}


@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"

    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"

    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None

    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None

    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True

    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285

    HTTP_TIMEOUT_SECONDS: int = 300

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)
        return Config(
            CALC_SPREADSHEET_ID=os.getenv("CALC_SPREADSHEET_ID", ""),
            INCOMING_FOLDER_ID=os.getenv("INCOMING_FOLDER_ID", ""),
            MANAGER_REPORT_FOLDER_ID=os.getenv("MANAGER_REPORT_FOLDER_ID", ""),
            ORDER_REPORT_FOLDER_ID=os.getenv("ORDER_REPORT_FOLDER_ID", ""),
            GID_MANAGER_PDF=os.getenv("GID_MANAGER_PDF", "1921812573"),
            GID_ORDER_CSV=os.getenv("GID_ORDER_CSV", "1875928148"),
            LOCATION_SHEET_TITLE=os.getenv("LOCATION_SHEET_TITLE", "REFR: Values"),
            LOCATION_NAMED_RANGE=os.getenv("LOCATION_NAMED_RANGE", "_locations"),
            TIMESTAMP_TZ=os.getenv("TIMESTAMP_TZ", "America/Chicago"),
            TIMESTAMP_FMT=os.getenv("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p"),
            TO_RECIPIENTS=_csv(os.getenv("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_csv(os.getenv("CC_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_b(os.getenv("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=_csv(os.getenv("REPORT_KEY_RUN_LIST", "")),
            REPORT_KEY_RECIPIENTS=_json_or_empty(os.getenv("REPORT_KEY_RECIPIENTS", "")),
            DEFAULT_ORDER_RECIPIENTS=_csv(os.getenv("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_b(
                os.getenv("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_b(
                os.getenv("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),
            SCOPES=_csv(
                os.getenv(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_b(os.getenv("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(os.getenv("REDIRECT_PORT", "58285")),
            HTTP_TIMEOUT_SECONDS=int(os.getenv("HTTP_TIMEOUT_SECONDS", "300")),
        )

    def to_env(self) -> str:
        # Serialize to .env format (simple)
        data = asdict(self)
        # Convert list and dict fields
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")
```

---
### file: favtrip/drive_utils.py

```python
from __future__ import annotations
import io
from googleapiclient.http import MediaIoBaseUpload


def find_latest_sheet(drive_svc, folder_id: str):
    q = (
        f"'{folder_id}' in parents and "
        "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    )
    resp = drive_svc.files().list(
        q=q, orderBy="createdTime desc", pageSize=1,
        fields="files(id,name,createdTime)"
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def upload_to_drive(drive_svc, data: bytes, name: str, mime: str, folder_id: str, to_sheet: bool=False):
    meta = {"name": name, "parents": [folder_id]}
    if to_sheet:
        meta["mimeType"] = "application/vnd.google-apps.spreadsheet"
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
    return drive_svc.files().create(
        body=meta, media_body=media, fields="id,name,mimeType,webViewLink"
    ).execute()

```

---
### file: favtrip/gmail_utils.py

```python
from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {ts} – {location}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Automated")
    msg.add_alternative(
        f"<p>Hi team,</p><p>Manager Report ({location})</p>"
        f"<a href='{pdf_link}'>Backup Link</a>", subtype="html"
    )
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)

```

---
### file: favtrip/google_client.py

```python
from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail
```

---
### file: favtrip/logger.py

```python
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass
```

---
### file: favtrip/pipeline.py

```python
from __future__ import annotations
import csv
import io
import re
import requests
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

from .config import Config
from .google_client import get_credentials, services
from .sheets_utils import (
    delete_sheet, copy_sheet_as, copy_first_sheet_as, refresh_sheets_with_prefix,
    get_value, first_gid
)
from .drive_utils import find_latest_sheet, upload_to_drive
from .gmail_utils import send_email, email_manager_report


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}&gid={gid}"
    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content


def timestamp_now(tz: str, fmt: str) -> str:
    return datetime.now(ZoneInfo(tz)).strftime(fmt)

import re

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _clean_emails(items):
    """
    Accepts a list or a comma-separated string and returns a list of valid emails.
    Trailing commas and blanks are removed. Invalid tokens are dropped silently.
    """
    if items is None:
        return []
    if isinstance(items, str):
        items = [p.strip() for p in items.split(",")]
    return [e for e in (p.strip() for p in items) if e and _EMAIL_RE.match(e)]

def _fallback_recipients(hint, *candidates):
    """
    Return the first non-empty, valid recipient list from the provided candidates.
    If all candidates are empty/invalid, raise a friendly error.
    """
    for c in candidates:
        cleaned = _clean_emails(c)
        if cleaned:
            return cleaned
    # Nothing usable found:
    raise ValueError(
        f"No valid recipients available for: {hint}. "
        f"Please provide at least one email in the UI or .env "
        f"(TO_RECIPIENTS, DEFAULT_ORDER_RECIPIENTS, or per-report-key)."
    )


@dataclass
class RunResult:
    ok: bool
    elapsed_seconds: int
    location: str
    timestamp: str
    manager_pdf_link: str | None
    full_order_link: str | None


def run_pipeline(cfg: Config, logger=None) -> RunResult:
    import time
    start = time.perf_counter()

    if logger:
        logger.info("Authorizing with Google APIs…")
    creds = get_credentials(cfg.SCOPES, cfg.REDIRECT_PORT, cfg.FORCE_REAUTH)
    sheets_svc, drive_svc, gmail_svc = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    if logger:
        logger.info("Google services ready")

    # Step 1: latest incoming
    if logger:
        logger.info("Finding latest incoming spreadsheet…")
    latest = find_latest_sheet(drive_svc, cfg.INCOMING_FOLDER_ID)
    if not latest:
        raise SystemExit("No incoming report found.")
    new_report_id = latest["id"]
    if logger:
        logger.info(f"Latest incoming: {latest['name']} ({new_report_id})")

    # Step 2: prep calculations workbook
    if logger:
        logger.info("Preparing calculations workbook…")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
    try:
        copy_sheet_as(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", "Last Week")
        if logger:
            logger.info("Copied old 'Current Week' to 'Last Week'")
    except Exception:
        if logger:
            logger.warn("No 'Current Week' sheet exists to copy")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
    copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")
    if logger:
        logger.info("Inserted new 'Current Week' from latest incoming report")

    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ')…")
    refresh_sheets_with_prefix(sheets_svc, cfg.CALC_SPREADSHEET_ID, prefix="REFR: ", logger=logger)

    # Step 3: read location code
    location = get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.LOCATION_NAMED_RANGE)
    ts = timestamp_now(cfg.TIMESTAMP_TZ, cfg.TIMESTAMP_FMT)
    if logger:
        logger.info(f"Location: {location}; Timestamp: {ts}")

    # Step 4A: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_MANAGER_PDF, "pdf")
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 4B: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_ORDER_CSV, "csv")

    # Step 4C: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_{ts}_{location}_FULL.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, "text/csv", cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_{ts}_{location}_FULL.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4D: Create per-report-key outputs and email
    text = master_csv_bytes.decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames:
        raise RuntimeError("CSV has no header.")
    report_col = next((h for h in reader.fieldnames if h.lower() == "report_key"), None)
    if not report_col:
        raise RuntimeError("Report_Key column missing.")
    rows = list(reader)
    groups = {}
    for r in rows:
        key = (r.get(report_col) or "").strip() or "UNASSIGNED"
        groups.setdefault(key, []).append(r)

    for key, key_rows in groups.items():
        if not cfg.USE_ALL_REPORT_KEYS and key.upper() not in (cfg.REPORT_KEY_RUN_LIST or []):
            continue
        csv_buf = io.StringIO()
        w = csv.DictWriter(csv_buf, fieldnames=reader.fieldnames, lineterminator="")
        w.writeheader()
        w.writerows(key_rows)
        key_bytes = csv_buf.getvalue().encode("utf-8")
        tag = clean_tag(key.upper())
        filename = f"Order_Report_{ts}_{location}_{tag}.csv"
        created = upload_to_drive(drive_svc, key_bytes, filename, "text/csv", cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{ts}_{location}_{tag}.pdf"

        # Prefer per‑key recipients; else default; else TO
        candidates = None
        if cfg.REPORT_KEY_RECIPIENTS:
            # The keys in UI are upper-cased; normalize here too
            candidates = cfg.REPORT_KEY_RECIPIENTS.get(tag)

        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.DEFAULT_ORDER_RECIPIENTS,
            cfg.TO_RECIPIENTS,
        )

        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – {tag}"
        msg["From"] = "me"
        msg["To"] = ", ".join(recipients)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi {key} team,\nYour order report is ready. \nGoogle Sheet: {created.get('webViewLink')}\nAttached: {pdfname}\n—Automated")
        msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=pdfname)
        if cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL:
            msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info(f"Emailed {tag}")

    # Step 4E: Send Manager Report

    to_list = _fallback_recipients("Manager Report (TO_RECIPIENTS)", cfg.TO_RECIPIENTS)
    cc_list = _clean_emails(cfg.CC_RECIPIENTS)
    email_manager_report(gmail_svc, "me", to_list, cc_list, pdf_name, pdf_bytes, manager_link, ts, location)
    
    if logger:
        logger.info("Manager email sent")

    # Step 4F: Send Full Order if needed
    full_link = full_created.get('webViewLink')
    if cfg.SEND_SEPARATE_FULL_ORDER_EMAIL:
        to_full = _fallback_recipients("FULL order", cfg.DEFAULT_ORDER_RECIPIENTS, cfg.TO_RECIPIENTS)
        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – FULL"
        msg["From"] = "me"
        msg["To"] = ", ".join(to_full)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi team,\nFULL order report is ready.\nSheet: {full_link}\nAttached: {full_pdf_name}\n—Automated")
        msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info("FULL order email sent")
    else:
        if logger:
            logger.info("Separate full order email disabled")

    elapsed = int(time.perf_counter() - start)
    if logger:
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        logger.info(f"Run completed in {h:02d}:{m:02d}:{s:02d}")

    return RunResult(True, elapsed, location, ts, manager_link, full_link)

```

---
### file: favtrip/sheets_utils.py

```python
from __future__ import annotations
import random
import time
from typing import Any, Dict, List


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def delete_sheet(svc, spreadsheet_id: str, title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), title)
    if s:
        body = {"requests": [{"deleteSheet": {"sheetId": s["sheetId"]}}]}
        svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def copy_sheet_as(svc, spreadsheet_id: str, src_title: str, new_title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), src_title)
    if not s:
        return None
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=spreadsheet_id,
        sheetId=s["sheetId"],
        body={"destinationSpreadsheetId": spreadsheet_id}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return new_id


def copy_first_sheet_as(svc, src_spreadsheet: str, dest_spreadsheet: str, new_title: str):
    meta = svc.spreadsheets().get(spreadsheetId=src_spreadsheet).execute()
    first_id = meta["sheets"][0]["properties"]["sheetId"]
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=src_spreadsheet,
        sheetId=first_id,
        body={"destinationSpreadsheetId": dest_spreadsheet}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=dest_spreadsheet, body=body).execute()
    return new_id


def refresh_sheets_with_prefix(svc, spreadsheet_id: str, prefix: str = "REFR: ", retries: int = 5, logger=None):
    sheets = list_sheets(svc, spreadsheet_id)
    targets = [s["properties"] for s in sheets if s["properties"]["title"].startswith(prefix)]
    for idx, t in enumerate(targets, start=1):
        body = {"requests": [{
            "findReplace": {
                "find": "=",
                "replacement": "=",
                "includeFormulas": True,
                "sheetId": t["sheetId"]
            }
        }]}
        attempt = 0
        while True:
            try:
                svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                if logger:
                    logger.info(f"[{idx}/{len(targets)}] Recalc OK: {t['title']}")
                break
            except Exception:
                attempt += 1
                if attempt > retries:
                    if logger:
                        logger.warn(f"FAILED recalc for {t['title']}")
                    break
                time.sleep(1 + random.random())


def get_value(svc, spreadsheet_id: str, sheet_title: str, named_range: str) -> str:
    try:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=named_range
        ).execute().get("values", [])
    except Exception:
        vals = []
    if not vals:
        try:
            vals = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_title}'!A1:A"
            ).execute().get("values", [])
        except Exception:
            vals = []
    return vals[0][0] if vals and vals[0] else "UNKNOWN"


def first_gid(svc, spreadsheet_id: str) -> int:
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return meta["sheets"][0]["properties"]["sheetId"]

```

---
### file: last_run.log

```
[2026-03-03 23:16:33] INFO: Authorizing with Google APIs…
[2026-03-03 23:16:33] INFO: Google services ready
[2026-03-03 23:16:33] INFO: Finding latest incoming spreadsheet…
[2026-03-03 23:16:34] INFO: Latest incoming: Testing Sales Report - Week 5 (1kNjmEbljdUIqJUjwh8e2-plfLWRZSGHxMHxL50-2ce4)
[2026-03-03 23:16:34] INFO: Preparing calculations workbook…
[2026-03-03 23:16:42] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 23:16:50] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 23:16:50] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 23:17:17] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 23:17:22] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 23:18:15] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 23:18:17] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-11-18-PM
[2026-03-03 23:18:17] INFO: Exporting Manager Report (PDF)…
[2026-03-03 23:18:21] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/1UILYk_poamIwe6JrFQUGZIz9az9iBsdU/view?usp=drivesdk
[2026-03-03 23:18:21] INFO: Exporting Master Order (CSV)…
[2026-03-03 23:18:34] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1bOcle5eMseFc43TZMnXLqgHR7_DcPs13ntb6PSBians/edit?usp=drivesdk
[2026-03-03 23:18:42] INFO: Emailed COFFEE
[2026-03-03 23:18:43] INFO: Manager email sent
[2026-03-03 23:18:43] INFO: Separate full order email disabled
[2026-03-03 23:18:43] INFO: Run completed in 00:02:09

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```

---
### file: ui_streamlit.py

```python
import os
import time
import threading
import streamlit as st

from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.google_client import (
    start_oauth,
    finish_oauth,
    load_valid_token,
    clear_token,
)

def _rerun():
    # Works on Streamlit >= 1.27 (st.rerun) and older (experimental_rerun)
    try:
        import streamlit as st
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

st.set_page_config(page_title="FavTrip Reporting Pipeline", page_icon="🧾", layout="wide")

# --- Larger Run button ---
st.markdown(
    """
    <style>
      div.stButton > button:first-child {
        font-size: 1.1rem;
        padding: 0.6rem 1.2rem;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🧾 FavTrip Reporting Pipeline")

cfg = Config.load()

# ----------------------------
# Session state: auth gating
# ----------------------------
if "auth_checked" not in st.session_state:
    # On first load: if token is missing/invalid/unrefreshable -> require auth
    st.session_state.auth_required = (load_valid_token(cfg.SCOPES) is None)
    st.session_state.oauth_flow = None
    st.session_state.oauth_url = None
    st.session_state.auth_checked = True

# Sidebar controls (always visible)
with st.sidebar:
    st.header("Defaults (.env)")

    if st.button("Force Google Re-Auth", type="secondary", use_container_width=True):
        clear_token()
        try:
            # Prefer auto flow
            from favtrip.google_client import login_via_local_server
            with st.status("Re-auth in progress (browser will open)…", expanded=True):
                creds = login_via_local_server(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.success("✅ Re-auth complete.")
            st.session_state.auth_required = False
            st.rerun()
        except Exception as e:
            # Fallback to manual method
            try:
                flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.session_state.oauth_flow = flow
                st.session_state.oauth_url = url
                st.session_state.auth_required = True  # show the auth panel
                st.info("Auto re-auth failed; showing manual method. Open the URL shown in the Authentication panel.")
                st.rerun()
            except Exception as e2:
                st.error(f"Failed to start re-auth: {e2}")

    # Developer mode (suppresses console prints when unchecked)
    dev_mode = st.checkbox("Run in developer mode", value=False)

    st.markdown("Edit values for this *one run*. Optionally tick **Update .env** to persist.")

# ----------------------------
# Authentication panel (shown only if auth required)
# ----------------------------
if st.session_state.auth_required:
    from favtrip.google_client import login_via_local_server  # import here to avoid circulars

    with st.expander("Google Authentication", expanded=True):
        st.caption(
            "Authentication is required before running. "
            "Click **Start authentication** to open the browser. "
            "We will wait for the redirect automatically."
        )

        # Preferred: automatic (open browser + capture redirect)
        if st.button("Start authentication (auto-open & capture)", type="primary"):
            try:
                with st.status("Waiting for Google authorization in your browser…", expanded=True):
                    creds = login_via_local_server(cfg.SCOPES, cfg.REDIRECT_PORT)
                    st.success("✅ Authentication complete. token.json saved.")
                st.session_state.oauth_flow = None
                st.session_state.oauth_url = None
                st.session_state.auth_required = False
                st.rerun()
            except Exception as e:
                st.error(f"Auto authentication failed: {e}. You can try the manual method below.")

        st.divider()
        st.write("**Manual method (fallback):**")

        # Manual fallback (existing behavior)
        col_a, col_b = st.columns([2, 1])
        with col_a:
            if st.button("Start authentication (get URL)"):
                try:
                    flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                    st.session_state.oauth_flow = flow
                    st.session_state.oauth_url = url
                    st.success("Auth URL generated below. Open it, grant access, and paste the redirect URL or code.")
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")
        with col_b:
            if st.session_state.oauth_url:
                st.link_button("Open Auth URL", st.session_state.oauth_url, use_container_width=True)

        if st.session_state.oauth_url:
            st.code(st.session_state.oauth_url, language="text")

        pasted = st.text_input(
            "Paste full redirect URL or the code here",
            value="",
            placeholder="https://... or the code",
        )

        if st.button("Complete authentication", type="secondary"):
            flow = st.session_state.get("oauth_flow")
            if not flow:
                st.warning("Click 'Start authentication (get URL)' first.")
            else:
                try:
                    creds = finish_oauth(flow, pasted)
                    st.session_state.oauth_flow = None
                    st.session_state.oauth_url = None
                    st.session_state.auth_required = False
                    st.success("✅ Authentication complete. token.json saved.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to finish OAuth: {e}")

# ----------------------------
# Only show Run Options if NOT requiring auth
# ----------------------------
if not st.session_state.auth_required:

    # ---- Run Form (Run button top-right) ----
    with st.form("run_form"):
        tl, tr = st.columns([4, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")
        with tr:
            submitted = st.form_submit_button("▶️ Run Pipeline", use_container_width=True)

        # --- Main options ---
        colA, colB, colC = st.columns(3)
        with colA:
            to = st.text_input("To Recipients (comma)", value=",".join(cfg.TO_RECIPIENTS or []))
            cc = st.text_input("CC Recipients (comma)", value=",".join(cfg.CC_RECIPIENTS or []))
        with colB:
            use_all = st.checkbox("Use all Report Keys in CSV", value=cfg.USE_ALL_REPORT_KEYS)
            report_keys = st.text_input("Report Keys to run (comma)", value=",".join(cfg.REPORT_KEY_RUN_LIST or []))
        with colC:
            include_full = st.checkbox("Attach FULL order in each email", value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL)
            send_full = st.checkbox("Send separate FULL order email", value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL)
            #force_reauth = st.checkbox("Force Google re-auth for this run", value=cfg.FORCE_REAUTH)

        # --- Per-report-key recipients editor (above Advanced) ---
        with st.expander("Per-Report-Key Recipients (optional)"):
            st.caption('Map **REPORT KEY (ALL CAPS)** → **Emails (comma)** (friendly editor for `REPORT_KEY_RECIPIENTS`).')
            rows = []
            if cfg.REPORT_KEY_RECIPIENTS:
                for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({"REPORT KEY (ALL CAPS)": k, "Emails (comma)": ",".join(v or [])})
            else:
                rows = [{"REPORT KEY (ALL CAPS)": "", "Emails (comma)": ""}]
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                use_container_width=True,
                key="rk_editor",
            )

        # --- Advanced (IDs, GIDs, Timezone, Redirect Port) ---
        with st.expander("Advanced (IDs, GIDs, Timezone, Redirect Port)"):
            col1, col2 = st.columns(2)
            with col1:
                calc_id = st.text_input("Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID)
                incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID)
                mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID)
                order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID)
                
                raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                redirect_port = st.number_input(
                    "Redirect Port (0 = auto)",
                    min_value=0,             # allow 0 explicitly
                    max_value=65535,
                    value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                    help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535.",
                )

            with col2:
                gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF))
                gid_order = st.text_input("Order CSV gid", value=str(cfg.GID_ORDER_CSV))
                loc_sheet = st.text_input("Location Sheet Title", value=cfg.LOCATION_SHEET_TITLE)
                loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE)
                tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ)
                tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT)

        save_env = st.checkbox("Update defaults in .env with the edited fields (optional)")

    # ----------------------------
    # Submission handling
    # ----------------------------
    def _split_emails(csv_str: str):
        return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

    if submitted:
        # Apply per-run config
        cfg.TO_RECIPIENTS = _split_emails(to)
        cfg.CC_RECIPIENTS = _split_emails(cc)
        cfg.USE_ALL_REPORT_KEYS = use_all
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

        cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
        cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full

        cfg.CALC_SPREADSHEET_ID = calc_id
        cfg.INCOMING_FOLDER_ID = incoming_id
        cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
        cfg.ORDER_REPORT_FOLDER_ID = order_folder
        cfg.REDIRECT_PORT = int(redirect_port)

        cfg.GID_MANAGER_PDF = gid_mgr
        cfg.GID_ORDER_CSV = gid_order
        cfg.LOCATION_SHEET_TITLE = loc_sheet
        cfg.LOCATION_NAMED_RANGE = loc_range
        cfg.TIMESTAMP_TZ = tz
        cfg.TIMESTAMP_FMT = tfmt

        # Per-key recipients from editor
        rk_map = {}
        for r in edited_rows:
            key = (r.get("REPORT KEY (ALL CAPS)") or "").strip().upper()
            emails_csv = r.get("Emails (comma)") or ""
            emails = _split_emails(emails_csv)
            if key and emails:
                rk_map[key] = emails
        cfg.REPORT_KEY_RECIPIENTS = rk_map

        if save_env:
            cfg.save()
            st.success("Saved updated defaults to .env")

        # If user checked "Force Google re-auth for this run", kick them into auth gating first.
        if cfg.FORCE_REAUTH:
            clear_token()
            try:
                flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.session_state.oauth_flow = flow
                st.session_state.oauth_url = url
                st.session_state.auth_required = True
                st.info("Re-auth required for this run. Open the URL shown in the Authentication panel.")
                _rerun()
            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")
        else:
            # --- Live run with timer + last log (no full log after completion) ---
            # Write all logs to last_run.log; overwrite on each run
            logger = StatusLogger(print_to_console=dev_mode, file_path="last_run.log", overwrite=True)
            result_holder = {"value": None, "error": None}

            def _runner():
                try:
                    result_holder["value"] = run_pipeline(cfg, logger=logger)
                except Exception as e:
                    result_holder["error"] = e

            t0 = time.perf_counter()
            th = threading.Thread(target=_runner, daemon=True)
            th.start()

            with st.status("Running pipeline…", expanded=True) as status:
                timer_ph = st.empty()
                lastlog_ph = st.empty()

                while th.is_alive():
                    elapsed = int(time.perf_counter() - t0)
                    timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                    lastlog_ph.markdown(f"**Last:** {logger.last_line()}")
                    time.sleep(0.5)

                th.join()
                elapsed = int(time.perf_counter() - t0)
                timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                lastlog_ph.markdown(f"**Last:** {logger.last_line()}")

                if result_holder["error"]:
                    st.error(f"Run failed: {result_holder['error']}")
                    status.update(label="❌ Failed", state="error")
                else:
                    result = result_holder["value"]
                    st.write("### Outputs")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Location", result.location)
                    col2.metric("Timestamp", result.timestamp)
                    mm = result.elapsed_seconds
                    col3.metric("Elapsed", f"{mm//3600:02d}:{(mm%3600)//60:02d}:{mm%60:02d}")
                    if result.manager_pdf_link:
                        st.success(f"Manager PDF: {result.manager_pdf_link}")
                    if result.full_order_link:
                        st.success(f"Full Order Sheet: {result.full_order_link}")
                    # Note: intentionally NOT showing the full live log post-run
                    status.update(label="✅ Completed", state="complete")
```

```

---
### file: documentation/README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: documentation/generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd C:\Users\rjrul\Downloads\favtrip_pipeline\documentation\
#python generate_code_bundle.py
```

---
### file: documentation/requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit

```

---
### file: favtrip/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: favtrip/config.py

```python
from __future__ import annotations

import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv
from typing import Any, Dict

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: str = ""):
    """
    Prefer Streamlit secrets (typed TOML values), else fall back to environment.
    Returns the raw value from st.secrets (could be bool/list/dict/str) or a str from env.
    """
    try:
        import streamlit as st  # imported lazily so local CLI still works
        val = st.secrets.get(key, None)
        if val is None:
            return os.getenv(key, default)
        return val
    except Exception:
        return os.getenv(key, default)


_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _coerce_bool(v, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default


def _coerce_csv(v) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]


def _coerce_json(v):
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}


# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300


import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv
from typing import Any, Dict

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: str = ""):
    """
    Prefer Streamlit secrets (typed TOML values), else fall back to environment.
    Returns the raw value from st.secrets (could be bool/list/dict/str) or a str from env.
    """
    try:
        import streamlit as st  # imported lazily so local CLI still works
        val = st.secrets.get(key, None)
        if val is None:
            return os.getenv(key, default)
        return val
    except Exception:
        return os.getenv(key, default)


_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _coerce_bool(v, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default


def _coerce_csv(v) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]


def _coerce_json(v):
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}


# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        """
        Load config from Streamlit secrets (preferred on cloud) or from .env (local dev),
        then overlay any values found in a Drive-backed JSON config (optional).
        Secrets may be typed (bool/list/dict), so we coerce safely.
        """
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)

        # ---------- 1) Base config from Secrets / .env ----------
        cfg = Config(
            CALC_SPREADSHEET_ID=str(_get_secret("CALC_SPREADSHEET_ID", "")),
            INCOMING_FOLDER_ID=str(_get_secret("INCOMING_FOLDER_ID", "")),
            MANAGER_REPORT_FOLDER_ID=str(_get_secret("MANAGER_REPORT_FOLDER_ID", "")),
            ORDER_REPORT_FOLDER_ID=str(_get_secret("ORDER_REPORT_FOLDER_ID", "")),

            GID_MANAGER_PDF=str(_get_secret("GID_MANAGER_PDF", "1921812573")),
            GID_ORDER_CSV=str(_get_secret("GID_ORDER_CSV", "1875928148")),
            LOCATION_SHEET_TITLE=str(_get_secret("LOCATION_SHEET_TITLE", "REFR: Values")),
            LOCATION_NAMED_RANGE=str(_get_secret("LOCATION_NAMED_RANGE", "_locations")),
            TIMESTAMP_TZ=str(_get_secret("TIMESTAMP_TZ", "America/Chicago")),
            TIMESTAMP_FMT=str(_get_secret("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p")),

            TO_RECIPIENTS=_coerce_csv(_get_secret("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_coerce_csv(_get_secret("CC_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_coerce_bool(_get_secret("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=[s.upper() for s in _coerce_csv(_get_secret("REPORT_KEY_RUN_LIST", ""))],
            REPORT_KEY_RECIPIENTS=_coerce_json(_get_secret("REPORT_KEY_RECIPIENTS", "{}")),
            DEFAULT_ORDER_RECIPIENTS=_coerce_csv(_get_secret("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_coerce_bool(
                _get_secret("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_coerce_bool(
                _get_secret("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),

            SCOPES=_coerce_csv(
                _get_secret(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,"
                    "https://www.googleapis.com/auth/spreadsheets,"
                    "https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_coerce_bool(_get_secret("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(str(_get_secret("REDIRECT_PORT", "58285")) or "58285"),
            HTTP_TIMEOUT_SECONDS=int(str(_get_secret("HTTP_TIMEOUT_SECONDS", "300")) or "300"),
        )

        # ---------- 2) Overlay from Drive JSON config (optional) ----------
        # If present, read a JSON file on Drive and apply known keys over cfg.
        try:
            import streamlit as st
            from favtrip.google_client import load_valid_token, services
            from favtrip.config_store import load_config_from_drive

            CONFIG_FILE_ID = ""
            if hasattr(st, "secrets"):
                CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

            creds = load_valid_token(cfg.SCOPES)
            if creds:
                _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
                overrides = load_config_from_drive(drive, CONFIG_FILE_ID or None)  # {} if not found
                if isinstance(overrides, dict) and overrides:
                    for k, v in overrides.items():
                        if hasattr(cfg, k):
                            setattr(cfg, k, v)
        except Exception:
            # Fail-open: if Drive/token not ready yet, just return base cfg
            pass

        return cfg

    # -----------------------------------------------------------------------------
    # .env serialization (still useful if you let users "Update defaults in .env")
    # -----------------------------------------------------------------------------

    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")

    # -----------------------------------------------------------------------------
    # .env serialization (still useful if you let users "Update defaults in .env")
    # -----------------------------------------------------------------------------

    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")
```

---
### file: favtrip/config_store.py

```python
# favtrip/config_store.py
from __future__ import annotations
import io
import json
from typing import Any, Dict, Optional
from googleapiclient.discovery import Resource
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

DEFAULT_CONFIG_FILENAME = "favtrip_config.json"
DEFAULT_MIMETYPE = "application/json"

def load_config_from_drive(drive: Resource, file_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Read the JSON config stored in Google Drive.
    If file_id is None, try to discover the newest file named DEFAULT_CONFIG_FILENAME.
    Returns {} if the file doesn't exist or is empty/invalid JSON.
    """
    # Discover by name if a specific id wasn't provided
    if not file_id:
        resp = drive.files().list(
            q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
            orderBy="modifiedTime desc",
            pageSize=1,
            fields="files(id,name,modifiedTime)"
        ).execute() or {}
        files = resp.get("files", [])
        if not files:
            return {}
        file_id = files[0]["id"]

    # Stream download the file
    request = drive.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    raw = buf.getvalue().decode("utf-8", errors="replace").strip()
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except Exception:
        return {}

def save_config_to_drive(
    drive: Resource,
    data: Dict[str, Any],
    file_id: Optional[str] = None,
    parent_folder_id: Optional[str] = None
) -> str:
    """
    Write JSON config to Google Drive.
    - If file_id provided, update that file.
    - Else upsert (update if found by name, otherwise create) DEFAULT_CONFIG_FILENAME.
    Returns the Drive file ID.
    """
    payload = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(payload), mimetype=DEFAULT_MIMETYPE, resumable=True)

    if file_id:
        updated = drive.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        return updated["id"]

    # Try to find an existing file by name to update
    resp = drive.files().list(
        q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
        orderBy="modifiedTime desc",
        pageSize=1,
        fields="files(id,name)"
    ).execute() or {}
    files = resp.get("files", [])
    if files:
        fid = files[0]["id"]
        updated = drive.files().update(fileId=fid, media_body=media).execute()
        return updated["id"]

    # Create a new file
    meta = {"name": DEFAULT_CONFIG_FILENAME}
    if parent_folder_id:
        meta["parents"] = [parent_folder_id]

    created = drive.files().create(
        body=meta,
        media_body=media,
        fields="id,name"
    ).execute()
    return created["id"]
```

---
### file: favtrip/drive_utils.py

```python
from __future__ import annotations
import io
from googleapiclient.http import MediaIoBaseUpload


def find_latest_sheet(drive_svc, folder_id: str):
    q = (
        f"'{folder_id}' in parents and "
        "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    )
    resp = drive_svc.files().list(
        q=q, orderBy="createdTime desc", pageSize=1,
        fields="files(id,name,createdTime)"
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def upload_to_drive(drive_svc, data: bytes, name: str, mime: str, folder_id: str, to_sheet: bool=False):
    meta = {"name": name, "parents": [folder_id]}
    if to_sheet:
        meta["mimeType"] = "application/vnd.google-apps.spreadsheet"
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
    return drive_svc.files().create(
        body=meta, media_body=media, fields="id,name,mimeType,webViewLink"
    ).execute()

```

---
### file: favtrip/gmail_utils.py

```python
from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {ts} – {location}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Automated")
    msg.add_alternative(
        f"<p>Hi team,</p><p>Manager Report ({location})</p>"
        f"<a href='{pdf_link}'>Backup Link</a>", subtype="html"
    )
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)

```

---
### file: favtrip/google_client.py

```python
from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail
```

---
### file: favtrip/logger.py

```python
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass
```

---
### file: favtrip/pipeline.py

```python
from __future__ import annotations
import csv
import io
import re
import requests
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

from .config import Config
from .google_client import get_credentials, services
from .sheets_utils import (
    delete_sheet, copy_sheet_as, copy_first_sheet_as, refresh_sheets_with_prefix,
    get_value, first_gid
)
from .drive_utils import find_latest_sheet, upload_to_drive
from .gmail_utils import send_email, email_manager_report


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}&gid={gid}"
    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content


def timestamp_now(tz: str, fmt: str) -> str:
    return datetime.now(ZoneInfo(tz)).strftime(fmt)

import re

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _clean_emails(items):
    """
    Accepts a list or a comma-separated string and returns a list of valid emails.
    Trailing commas and blanks are removed. Invalid tokens are dropped silently.
    """
    if items is None:
        return []
    if isinstance(items, str):
        items = [p.strip() for p in items.split(",")]
    return [e for e in (p.strip() for p in items) if e and _EMAIL_RE.match(e)]

def _fallback_recipients(hint, *candidates):
    """
    Return the first non-empty, valid recipient list from the provided candidates.
    If all candidates are empty/invalid, raise a friendly error.
    """
    for c in candidates:
        cleaned = _clean_emails(c)
        if cleaned:
            return cleaned
    # Nothing usable found:
    raise ValueError(
        f"No valid recipients available for: {hint}. "
        f"Please provide at least one email in the UI or .env "
        f"(TO_RECIPIENTS, DEFAULT_ORDER_RECIPIENTS, or per-report-key)."
    )


@dataclass
class RunResult:
    ok: bool
    elapsed_seconds: int
    location: str
    timestamp: str
    manager_pdf_link: str | None
    full_order_link: str | None


def run_pipeline(cfg: Config, logger=None) -> RunResult:
    import time
    start = time.perf_counter()

    if logger:
        logger.info("Authorizing with Google APIs…")
    creds = get_credentials(cfg.SCOPES, cfg.REDIRECT_PORT, cfg.FORCE_REAUTH)
    sheets_svc, drive_svc, gmail_svc = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    if logger:
        logger.info("Google services ready")

    # Step 1: latest incoming
    if logger:
        logger.info("Finding latest incoming spreadsheet…")
    latest = find_latest_sheet(drive_svc, cfg.INCOMING_FOLDER_ID)
    if not latest:
        raise RuntimeError("No incoming report found.")
    new_report_id = latest["id"]
    if logger:
        logger.info(f"Latest incoming: {latest['name']} ({new_report_id})")

    # Step 2: prep calculations workbook
    if logger:
        logger.info("Preparing calculations workbook…")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
    try:
        copy_sheet_as(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", "Last Week")
        if logger:
            logger.info("Copied old 'Current Week' to 'Last Week'")
    except Exception:
        if logger:
            logger.warn("No 'Current Week' sheet exists to copy")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
    copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")
    if logger:
        logger.info("Inserted new 'Current Week' from latest incoming report")

    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ')…")
    refresh_sheets_with_prefix(sheets_svc, cfg.CALC_SPREADSHEET_ID, prefix="REFR: ", logger=logger)

    # Step 3: read location code
    location = get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.LOCATION_NAMED_RANGE)
    ts = timestamp_now(cfg.TIMESTAMP_TZ, cfg.TIMESTAMP_FMT)
    if logger:
        logger.info(f"Location: {location}; Timestamp: {ts}")

    # Step 4A: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_MANAGER_PDF, "pdf")
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 4B: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_ORDER_CSV, "csv")

    # Step 4C: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_{ts}_{location}_FULL.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, "text/csv", cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_{ts}_{location}_FULL.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4D: Create per-report-key outputs and email
    text = master_csv_bytes.decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames:
        raise RuntimeError("CSV has no header.")
    report_col = next((h for h in reader.fieldnames if h.lower() == "report_key"), None)
    if not report_col:
        raise RuntimeError("Report_Key column missing.")
    rows = list(reader)
    groups = {}
    for r in rows:
        key = (r.get(report_col) or "").strip() or "UNASSIGNED"
        groups.setdefault(key, []).append(r)

    for key, key_rows in groups.items():
        if not cfg.USE_ALL_REPORT_KEYS and key.upper() not in (cfg.REPORT_KEY_RUN_LIST or []):
            continue
        csv_buf = io.StringIO()
        w = csv.DictWriter(csv_buf, fieldnames=reader.fieldnames, lineterminator="")
        w.writeheader()
        w.writerows(key_rows)
        key_bytes = csv_buf.getvalue().encode("utf-8")
        tag = clean_tag(key.upper())
        filename = f"Order_Report_{ts}_{location}_{tag}.csv"
        created = upload_to_drive(drive_svc, key_bytes, filename, "text/csv", cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{ts}_{location}_{tag}.pdf"

        # Prefer per‑key recipients; else default; else TO
        candidates = None
        if cfg.REPORT_KEY_RECIPIENTS:
            # The keys in UI are upper-cased; normalize here too
            candidates = cfg.REPORT_KEY_RECIPIENTS.get(tag)

        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.DEFAULT_ORDER_RECIPIENTS,
            cfg.TO_RECIPIENTS,
        )

        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – {tag}"
        msg["From"] = "me"
        msg["To"] = ", ".join(recipients)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi {key} team,\nYour order report is ready. \nGoogle Sheet: {created.get('webViewLink')}\nAttached: {pdfname}\n—Automated")
        msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=pdfname)
        if cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL:
            msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info(f"Emailed {tag}")

    # Step 4E: Send Manager Report

    to_list = _fallback_recipients("Manager Report (TO_RECIPIENTS)", cfg.TO_RECIPIENTS)
    cc_list = _clean_emails(cfg.CC_RECIPIENTS)
    email_manager_report(gmail_svc, "me", to_list, cc_list, pdf_name, pdf_bytes, manager_link, ts, location)
    
    if logger:
        logger.info("Manager email sent")

    # Step 4F: Send Full Order if needed
    full_link = full_created.get('webViewLink')
    if cfg.SEND_SEPARATE_FULL_ORDER_EMAIL:
        to_full = _fallback_recipients("FULL order", cfg.DEFAULT_ORDER_RECIPIENTS, cfg.TO_RECIPIENTS)
        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – FULL"
        msg["From"] = "me"
        msg["To"] = ", ".join(to_full)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi team,\nFULL order report is ready.\nSheet: {full_link}\nAttached: {full_pdf_name}\n—Automated")
        msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info("FULL order email sent")
    else:
        if logger:
            logger.info("Separate full order email disabled")

    elapsed = int(time.perf_counter() - start)
    if logger:
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        logger.info(f"Run completed in {h:02d}:{m:02d}:{s:02d}")

    return RunResult(True, elapsed, location, ts, manager_link, full_link)

```

---
### file: favtrip/sheets_utils.py

```python
from __future__ import annotations
import random
import time
from typing import Any, Dict, List


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def delete_sheet(svc, spreadsheet_id: str, title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), title)
    if s:
        body = {"requests": [{"deleteSheet": {"sheetId": s["sheetId"]}}]}
        svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def copy_sheet_as(svc, spreadsheet_id: str, src_title: str, new_title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), src_title)
    if not s:
        return None
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=spreadsheet_id,
        sheetId=s["sheetId"],
        body={"destinationSpreadsheetId": spreadsheet_id}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return new_id


def copy_first_sheet_as(svc, src_spreadsheet: str, dest_spreadsheet: str, new_title: str):
    meta = svc.spreadsheets().get(spreadsheetId=src_spreadsheet).execute()
    first_id = meta["sheets"][0]["properties"]["sheetId"]
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=src_spreadsheet,
        sheetId=first_id,
        body={"destinationSpreadsheetId": dest_spreadsheet}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=dest_spreadsheet, body=body).execute()
    return new_id


def refresh_sheets_with_prefix(svc, spreadsheet_id: str, prefix: str = "REFR: ", retries: int = 5, logger=None):
    sheets = list_sheets(svc, spreadsheet_id)
    targets = [s["properties"] for s in sheets if s["properties"]["title"].startswith(prefix)]
    for idx, t in enumerate(targets, start=1):
        body = {"requests": [{
            "findReplace": {
                "find": "=",
                "replacement": "=",
                "includeFormulas": True,
                "sheetId": t["sheetId"]
            }
        }]}
        attempt = 0
        while True:
            try:
                svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                if logger:
                    logger.info(f"[{idx}/{len(targets)}] Recalc OK: {t['title']}")
                break
            except Exception:
                attempt += 1
                if attempt > retries:
                    if logger:
                        logger.warn(f"FAILED recalc for {t['title']}")
                    break
                time.sleep(1 + random.random())


def get_value(svc, spreadsheet_id: str, sheet_title: str, named_range: str) -> str:
    try:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=named_range
        ).execute().get("values", [])
    except Exception:
        vals = []
    if not vals:
        try:
            vals = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_title}'!A1:A"
            ).execute().get("values", [])
        except Exception:
            vals = []
    return vals[0][0] if vals and vals[0] else "UNKNOWN"


def first_gid(svc, spreadsheet_id: str) -> int:
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return meta["sheets"][0]["properties"]["sheetId"]

```

---
### file: last_run.log

```
[2026-03-03 23:16:33] INFO: Authorizing with Google APIs…
[2026-03-03 23:16:33] INFO: Google services ready
[2026-03-03 23:16:33] INFO: Finding latest incoming spreadsheet…
[2026-03-03 23:16:34] INFO: Latest incoming: Testing Sales Report - Week 5 (1kNjmEbljdUIqJUjwh8e2-plfLWRZSGHxMHxL50-2ce4)
[2026-03-03 23:16:34] INFO: Preparing calculations workbook…
[2026-03-03 23:16:42] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 23:16:50] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 23:16:50] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 23:17:17] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 23:17:22] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 23:18:15] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 23:18:17] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-11-18-PM
[2026-03-03 23:18:17] INFO: Exporting Manager Report (PDF)…
[2026-03-03 23:18:21] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/1UILYk_poamIwe6JrFQUGZIz9az9iBsdU/view?usp=drivesdk
[2026-03-03 23:18:21] INFO: Exporting Master Order (CSV)…
[2026-03-03 23:18:34] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1bOcle5eMseFc43TZMnXLqgHR7_DcPs13ntb6PSBians/edit?usp=drivesdk
[2026-03-03 23:18:42] INFO: Emailed COFFEE
[2026-03-03 23:18:43] INFO: Manager email sent
[2026-03-03 23:18:43] INFO: Separate full order email disabled
[2026-03-03 23:18:43] INFO: Run completed in 00:02:09

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: requirements.txt

```text
-r documentation/requirements.txt
```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```

---
### file: ui_streamlit.py

```python
import os
import time
import threading
import streamlit as st


import base64, hashlib, secrets, json
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials


# --- add near the top of ui_streamlit.py (after imports) ---
import base64, hashlib, json, secrets
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials


from favtrip.google_client import load_valid_token, services
from favtrip.config_store import load_config_from_drive

from streamlit.components.v1 import html
from streamlit.components.v1 import html as _html_msg


def _b64url(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")

def _pkce_pair() -> tuple[str, str]:
    """
    Generate a high-entropy PKCE code_verifier and its S256 code_challenge.
    RFC 7636 requires 43–128 chars; this approach yields a URL-safe value.  [2](https://auth0.com/docs/get-started/authentication-and-authorization-flow/authorization-code-flow-with-pkce)
    """
    verifier = _b64url(secrets.token_bytes(64))        # ~86 chars, URL-safe, no padding
    challenge = _b64url(hashlib.sha256(verifier.encode("ascii")).digest())
    return verifier, challenge

def _redirect_base() -> str:
    """
    Always return a non-empty redirect base that exactly matches your OAuth client's
    Authorized redirect URI. Prefer Secrets; normalize to one trailing slash.
    """
    base = (st.secrets.get("APP_BASE_URL", "") or "").strip()
    if not base:
        # Fallback to request (often available), still normalized
        try:
            base = (st.request.url_root or "").strip()
        except Exception:
            base = ""
    if not base:
        st.error("OAuth redirect base is not set. Define APP_BASE_URL in Secrets.")
        st.stop()
    return base.rstrip("/") + "/"

def start_web_oauth(scopes):
    """
    Build an authorization URL that:
      - uses a stable redirect_uri (from Secrets)
      - uses explicit PKCE (S256)
      - embeds the code_verifier inside the state (base64url(JSON))
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    redirect = _redirect_base()

    # Explicit PKCE (stateless across redirect)
    code_verifier, code_challenge = _pkce_pair()

    # CSRF token + verifier encoded into state that Google will return unchanged.
    state_obj = {
        "csrf": _b64url(secrets.token_bytes(16)),
        "v": code_verifier,
        "r": redirect,
    }
    state_b64 = _b64url(json.dumps(state_obj).encode("utf-8"))

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
        state=state_b64,
        # PKCE parameters:
        code_challenge=code_challenge,
        code_challenge_method="S256",
    )

    # Keep only minimal context; state carries the verifier.
    st.session_state["_oauth_redirect"] = redirect
    return auth_url

def _parse_state(state_b64: str) -> dict:
    # Add padding back for base64 decoding if needed
    padding = "=" * ((4 - len(state_b64) % 4) % 4)
    raw = base64.urlsafe_b64decode(state_b64 + padding)
    return json.loads(raw.decode("utf-8"))

def finish_web_oauth(code: str, state_b64: str, scopes):
    """
    Recreate a Flow with the same redirect_uri and exchange code + code_verifier for tokens.
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    state_obj = _parse_state(state_b64)
    code_verifier = state_obj.get("v")
    redirect = state_obj.get("r") or st.session_state.get("_oauth_redirect") or _redirect_base()

    if not code_verifier:
        st.error("OAuth state did not include a PKCE code_verifier.")
        st.stop()

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    # IMPORTANT: include code_verifier here to satisfy PKCE.  [1](https://github.com/googleapis/google-auth-library-python-oauthlib/issues/46)
    flow.fetch_token(code=code, code_verifier=code_verifier)

    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


from favtrip.config_store import save_config_to_drive
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.google_client import (
    start_oauth,
    finish_oauth,
    load_valid_token,
    clear_token,
    services,
)

def _rerun():
    # Works on Streamlit >= 1.27 (st.rerun) and older (experimental_rerun)
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

st.set_page_config(page_title="FavTrip Reporting Pipeline", page_icon="🧾", layout="wide")


# --- Larger Run button ---
st.markdown(
    """
    <style>
      div.stButton > button:first-child {
        font-size: 1.1rem;
        padding: 0.6rem 1.2rem;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🧾 FavTrip Reporting Pipeline")

cfg = Config.load()

#Old Config Diagonistics
def hide():
    with st.expander("🔎 Config diagnostics", expanded=False):
        try:
            CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
            st.write(f"CONFIG_FILE_ID (Secrets): `{CONFIG_FILE_ID or '(unset)'}`")

            token = load_valid_token(cfg.SCOPES)
            if not token:
                st.warning("No valid Google token yet. Drive overlay hasn't run.")
            else:
                _sheets, drive, _gmail = services(token, cfg.HTTP_TIMEOUT_SECONDS)
                overrides = load_config_from_drive(drive, CONFIG_FILE_ID or None)
                if overrides:
                    st.success(f"Loaded Drive overrides ({len(overrides)} keys).")
                    # show just the most relevant keys
                    show = {k: overrides.get(k) for k in [
                        "CALC_SPREADSHEET_ID", "INCOMING_FOLDER_ID",
                        "MANAGER_REPORT_FOLDER_ID", "ORDER_REPORT_FOLDER_ID",
                        "TO_RECIPIENTS", "REPORT_KEY_RUN_LIST"
                    ]}
                    st.json(show)
                else:
                    st.error("No Drive config found (or file empty).")
                    st.caption("If you just saved defaults to Drive, copy its file id into Secrets as CONFIG_FILE_ID and rerun.")
        except Exception as e:
            st.error(f"Diagnostics failed: {e}")

    st.caption("**Effective config (current page state):**")
    colA, colB = st.columns(2)
    with colA:
        st.write("**IDs**")
        st.code({
            "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
            "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
            "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
            "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,
        }, language="json")
        st.write("**Recipients**")
        st.code({
            "TO_RECIPIENTS": cfg.TO_RECIPIENTS,
            "CC_RECIPIENTS": cfg.CC_RECIPIENTS,
            "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,
        }, language="json")
    with colB:
        st.write("**Flags & lists**")
        st.code({
            "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
            "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,
            "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,
        }, language="json")
        st.write("**GIDs / TZ**")
        st.code({
            "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
            "GID_ORDER_CSV": cfg.GID_ORDER_CSV,
            "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
            "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,
            "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
            "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,
        }, language="json")
    
params = st.query_params
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        st.success("✅ Google authentication complete.")
        st.query_params.clear()
        # NEW: tell the opener to refresh and then close the popup
        html(
            """
            <script>
              try {
                if (window.opener && !window.opener.closed) {
                  window.opener.postMessage({type: "favtrip_oauth_done"}, "*");
                }
              } catch (e) {}
              // Close this popup
              window.close();
            </script>
            """,
            height=0,
        )
        # Fallback: in case the browser blocks close(), allow manual continue
        st.caption("You can close this window if it didn't close automatically.")
    except Exception as e:
        st.error(f"OAuth error: {e}")
if load_valid_token(cfg.SCOPES):
    cfg = Config.load()

# ----------------------------
# Session state: auth gating
# ----------------------------
if "auth_checked" not in st.session_state:
    # On first load: if token is missing/invalid/unrefreshable -> require auth
    st.session_state.auth_required = (load_valid_token(cfg.SCOPES) is None)
    st.session_state.oauth_flow = None
    st.session_state.oauth_url = None
    st.session_state.auth_checked = True

# Sidebar controls (always visible)
with st.sidebar:
    st.header("Utilities")

    if st.button("Google Sign Out", type="secondary", use_container_width=True):
        clear_token()
        for key in ["auth_required", "oauth_flow", "oauth_url", "auth_checked"]:
            if key in st.session_state:
                del st.session_state[key]
        _rerun()


# ----------------------------
# Authentication panel (shown only if auth required)
# ----------------------------

# --- Authentication expander (shown only if auth required) ---
if st.session_state.auth_required:
    with st.expander("Google Authentication", expanded=True):
        st.caption(
            "Authentication is required before running. "
            "Click **Sign in with Google** to open the consent screen (it will open in a new tab)."
        )

        # Use a placeholder so we can remove the button immediately
        sign_in_ph = st.empty()

        # Render the button inside the placeholder
        clicked = sign_in_ph.button("Sign in with Google", type="primary", use_container_width=True)

        if clicked:
            try:
                auth_url = start_web_oauth(cfg.SCOPES)

                # 1) Clear the placeholder to remove the button immediately
                sign_in_ph.empty()

                # 2) Show the message in the current tab
                st.markdown(
                    """
                    <div style="
                        display:flex;align-items:center;justify-content:center;
                        height:55vh;text-align:center;
                        font-family: system-ui, Segoe UI, Roboto, Helvetica, Arial, sans-serif;">
                      <div>
                        <h2 style="margin-bottom:0.5rem;">You're being signed in…</h2>
                        <p style="font-size:1.05rem;opacity:.9;">
                          A new browser tab was opened for Google sign‑in.<br/>
                          <strong>After it completes, you may close this tab.</strong>
                        </p>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # (Optional) If user returns to this tab later, refresh to show signed-in state
                html(
                    """
                    <script>
                      document.addEventListener("visibilitychange", function() {
                        if (!document.hidden) { location.reload(); }
                      });
                    </script>
                    """,
                    height=0,
                )

                # 3) Open Google auth in a NEW tab (leave this tab on the message)
                html(
                    f"""
                    <script>
                      window.open({json.dumps(auth_url)}, "_blank", "noopener");
                    </script>
                    """,
                    height=0,
                )

                # 4) Stop further rendering so the message persists
                st.stop()

            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")

        with st.expander("Having trouble?", expanded=False):
            st.write(
                "- The Google authorization page opens in a **new browser tab**.\n"
                "- After completing consent, **close this tab** and use the new tab.\n"
                "- If you renamed your Streamlit app or URL, ensure the Google OAuth "
                "Authorized redirect URI matches exactly (including trailing slash)."
            )

# ----------------------------
# Only show Run Options if NOT requiring auth
# ----------------------------
if not st.session_state.auth_required:

    # ---- Run Form (Run button top-right) ----
    with st.form("run_form"):
        tl, tr = st.columns([4, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")
        with tr:
            submitted = st.form_submit_button("▶️ Run Pipeline", use_container_width=True)

        # --- Main options ---
        colA, colB, colC = st.columns(3)
        with colA:
            to = st.text_input("To Recipients (comma)", value=",".join(cfg.TO_RECIPIENTS or []))
            cc = st.text_input("CC Recipients (comma)", value=",".join(cfg.CC_RECIPIENTS or []))
        with colB:
            use_all = st.checkbox("Use all Report Keys in CSV", value=cfg.USE_ALL_REPORT_KEYS)
            report_keys = st.text_input("Report Keys to run (comma)", value=",".join(cfg.REPORT_KEY_RUN_LIST or []))
        with colC:
            include_full = st.checkbox("Attach FULL order in each email", value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL)
            send_full = st.checkbox("Send separate FULL order email", value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL)
            #force_reauth = st.checkbox("Force Google re-auth for this run", value=cfg.FORCE_REAUTH)

        # --- Per-report-key recipients editor (above Advanced) ---
        with st.expander("Per-Report-Key Recipients (optional)"):
            st.caption('Map **REPORT KEY (ALL CAPS)** → **Emails (comma)** (friendly editor for `REPORT_KEY_RECIPIENTS`).')
            rows = []
            if cfg.REPORT_KEY_RECIPIENTS:
                for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({"REPORT KEY (ALL CAPS)": k, "Emails (comma)": ",".join(v or [])})
            else:
                rows = [{"REPORT KEY (ALL CAPS)": "", "Emails (comma)": ""}]
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                use_container_width=True,
                key="rk_editor",
            )

        # --- Advanced (IDs, GIDs, Timezone, Redirect Port) ---
        with st.expander("Advanced (IDs, GIDs, Timezone, Redirect Port)"):
            col1, col2 = st.columns(2)
            with col1:
                calc_id = st.text_input("Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID)
                incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID)
                mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID)
                order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID)
                
                raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                redirect_port = st.number_input(
                    "Redirect Port (0 = auto)",
                    min_value=0,             # allow 0 explicitly
                    max_value=65535,
                    value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                    help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535.",
                )

            with col2:
                gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF))
                gid_order = st.text_input("Order CSV gid", value=str(cfg.GID_ORDER_CSV))
                loc_sheet = st.text_input("Location Sheet Title", value=cfg.LOCATION_SHEET_TITLE)
                loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE)
                tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ)
                tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT)

        save_drive_defaults = st.checkbox("Update defaults", value=False)

    # ----------------------------
    # Submission handling
    # ----------------------------
    def _split_emails(csv_str: str):
        return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

    if submitted:
        # Apply per-run config
        cfg.TO_RECIPIENTS = _split_emails(to)
        cfg.CC_RECIPIENTS = _split_emails(cc)
        cfg.USE_ALL_REPORT_KEYS = use_all
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

        cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
        cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full

        cfg.CALC_SPREADSHEET_ID = calc_id
        cfg.INCOMING_FOLDER_ID = incoming_id
        cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
        cfg.ORDER_REPORT_FOLDER_ID = order_folder
        cfg.REDIRECT_PORT = int(redirect_port)

        cfg.GID_MANAGER_PDF = gid_mgr
        cfg.GID_ORDER_CSV = gid_order
        cfg.LOCATION_SHEET_TITLE = loc_sheet
        cfg.LOCATION_NAMED_RANGE = loc_range
        cfg.TIMESTAMP_TZ = tz
        cfg.TIMESTAMP_FMT = tfmt

        # Per-key recipients from editor
        rk_map = {}
        for r in edited_rows:
            key = (r.get("REPORT KEY (ALL CAPS)") or "").strip().upper()
            emails_csv = r.get("Emails (comma)") or ""
            emails = _split_emails(emails_csv)
            if key and emails:
                rk_map[key] = emails
        cfg.REPORT_KEY_RECIPIENTS = rk_map

        # --- Save edited defaults to Drive JSON (optional) ---
        if save_drive_defaults:
            try:
                # Ensure we have a user token first
                creds = load_valid_token(cfg.SCOPES)
                if not creds:
                    st.error("Not authenticated. Please complete Google sign‑in first (top of page).")
                else:
                    # Drive service
                    _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                    # What we persist (the fields you asked to move out of Secrets)
                    drive_defaults = {
                        "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
                        "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
                        "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
                        "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,

                        "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
                        "GID_ORDER_CSV": cfg.GID_ORDER_CSV,

                        "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
                        "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,

                        "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
                        "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,

                        "TO_RECIPIENTS": cfg.TO_RECIPIENTS,   # lists are fine; JSON keeps types
                        "CC_RECIPIENTS": cfg.CC_RECIPIENTS,

                        "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
                        "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,

                        "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,

                        "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,

                        "INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL": cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL,
                        "SEND_SEPARATE_FULL_ORDER_EMAIL": cfg.SEND_SEPARATE_FULL_ORDER_EMAIL,
                    }

                    # If you have CONFIG_FILE_ID in Secrets, we update that exact file.
                    # Otherwise we'll upsert a file named 'favtrip_config.json' and return its id.
                    CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
                    new_id = save_config_to_drive(
                        drive,
                        drive_defaults,
                        file_id=CONFIG_FILE_ID or None,   # update/pin if set, else upsert by name
                        # parent_folder_id=None,          # optional: set a folder id to create under
                    )

                    st.success(f"Saved defaults to Drive config (file id: {new_id}).")
                    if not CONFIG_FILE_ID:
                        st.info(
                            "Tip: add this ID to Streamlit Secrets as `CONFIG_FILE_ID` to pin the same file for all runs:\n"
                            f"`{new_id}`"
                        )
            except Exception as e:
                st.error(f"Failed to save defaults to Drive: {e}")

        # If user checked "Force Google re-auth for this run", kick them into auth gating first.
        if cfg.FORCE_REAUTH:
            clear_token()
            try:
                flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.session_state.oauth_flow = flow
                st.session_state.oauth_url = url
                st.session_state.auth_required = True
                st.info("Re-auth required for this run. Open the URL shown in the Authentication panel.")
                _rerun()
            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")
        else:
            # --- Live run with timer + last log (no full log after completion) ---
            # Write all logs to last_run.log; overwrite on each run
            logger = StatusLogger(print_to_console=True, file_path="last_run.log", overwrite=True)
            result_holder = {"value": None, "error": None}

            def _runner():
                try:
                    result_holder["value"] = run_pipeline(cfg, logger=logger)
                # Catch BaseException so SystemExit / KeyboardInterrupt are captured too
                except BaseException as e:
                    result_holder["error"] = e

            t0 = time.perf_counter()
            th = threading.Thread(target=_runner, daemon=True)
            th.start()

            with st.status("Running pipeline…", expanded=True) as status:
                timer_ph = st.empty()
                lastlog_ph = st.empty()
                while th.is_alive():
                    elapsed = int(time.perf_counter() - t0)
                    timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                    lastlog_ph.markdown(f"**Last:** {logger.last_line()}")
                    time.sleep(0.5)

                th.join()
                elapsed = int(time.perf_counter() - t0)
                timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                lastlog_ph.markdown(f"**Last:** {logger.last_line()}")

                if result_holder["error"]:
                    st.error(f"Run failed: {result_holder['error']}")
                    # Optional during debugging: show stack trace (remove later for a cleaner UI)
                    try:
                        st.exception(result_holder["error"])
                    except Exception:
                        pass
                    status.update(label="❌ Failed", state="error")
                else:
                    result = result_holder["value"]
                    if result is None:
                        st.error("Run finished without returning a result. Check logs and inputs (IDs, Drive access).")
                        status.update(label="⚠️ No result", state="error")
                    else:
                        st.write("### Outputs")
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Location", result.location)
                        col2.metric("Timestamp", result.timestamp)
                        mm = result.elapsed_seconds
                        col3.metric("Elapsed", f"{mm//3600:02d}:{(mm%3600)//60:02d}:{mm%60:02d}")

                        if getattr(result, "manager_pdf_link", None):
                            st.success(f"Manager PDF: {result.manager_pdf_link}")
                        if getattr(result, "full_order_link", None):
                            st.success(f"Full Order Sheet: {result.full_order_link}")

                        status.update(label="✅ Completed", state="complete")
```

---
### file: web_url_credentials.json

```json
{"web":{"client_id":"674901450584-i6ncesf794ecg6u7olebqkag8gavbuc3.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["https://favtripreporting-dev1.streamlit.app/"]}}
```

```

---
### file: documentation/README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: documentation/generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd "C:\Users\rjrul\OneDrive - University of Iowa\000 Current Semester\004 BAIS 4150 BAIS Capstone\favtrip_reporting"
#python documentation/generate_code_bundle.py
```

---
### file: documentation/git_workflow.txt

```text

---
1. Create a dev branch

# Switch to dev branch
    git checkout dev

# Push dev branch to GitHub (the -u flag sets upstream tracking)
    git push -u origin dev

---
2. Do all development on dev
Make changes normally, then stage, commit, and push:

    git add .
    git commit -m "your message here"
    git push

All pushes go to dev (not main).

---
3. Push dev → main when ready (production release)
Use a Pull Request on GitHub:
1. Go to the repo on GitHub.
2. You will see a banner offering “Compare & Pull Request” (dev → main).
3. Open the Pull Request.
4. Review and click “Merge Pull Request”.

After merging, update your local main branch:

    git checkout main
    git pull

---
4. Keep dev updated with the latest main
After merging into main, update dev:

    git checkout dev
    git pull origin main

Or:

    git merge main

This prevents dev from drifting behind main.

---
5. Summary
- main = production, always stable
- dev = development, experimental work
- Merge dev → main only after testing

```

---
### file: documentation/requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit

```

---
### file: favtrip/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: favtrip/config.py

```python
from __future__ import annotations

import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv
from typing import Any, Dict

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: str = ""):
    """
    Prefer Streamlit secrets (typed TOML values), else fall back to environment.
    Returns the raw value from st.secrets (could be bool/list/dict/str) or a str from env.
    """
    try:
        import streamlit as st  # imported lazily so local CLI still works
        val = st.secrets.get(key, None)
        if val is None:
            return os.getenv(key, default)
        return val
    except Exception:
        return os.getenv(key, default)


_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _coerce_bool(v, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default


def _coerce_csv(v) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]


def _coerce_json(v):
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}


# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True
    EMAIL_MANAGER_REPORT: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300


import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv
from typing import Any, Dict

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------


def _get_secret(key: str, default: Any = None) -> Any:
    """
    Read from Streamlit secrets if present, else env var, else default.
    Does not raise if key missing; returns `default`.
    """
    # Try Streamlit secrets (if running inside Streamlit with secrets configured)
    try:
        import streamlit as st  # imported lazily to avoid hard dependency
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets.get(key, default)
    except Exception:
        pass

    # Fallback to environment variables (including those loaded by dotenv)
    val = os.getenv(key, default)
    return val



_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _coerce_bool(v, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default


def _coerce_csv(v) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]


def _coerce_json(v):
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}


# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True
    EMAIL_MANAGER_REPORT: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        """
        Load config from Streamlit secrets (preferred on cloud) or from .env (local dev),
        then overlay any values found in a Drive-backed JSON config (optional).
        Secrets may be typed (bool/list/dict), so we coerce safely.
        """
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)

        # ---------- 1) Base config from Secrets / .env ----------
        cfg = Config(
            CALC_SPREADSHEET_ID=str(_get_secret("CALC_SPREADSHEET_ID", "")),
            INCOMING_FOLDER_ID=str(_get_secret("INCOMING_FOLDER_ID", "")),
            MANAGER_REPORT_FOLDER_ID=str(_get_secret("MANAGER_REPORT_FOLDER_ID", "")),
            ORDER_REPORT_FOLDER_ID=str(_get_secret("ORDER_REPORT_FOLDER_ID", "")),

            GID_MANAGER_PDF=str(_get_secret("GID_MANAGER_PDF", "1921812573")),
            GID_ORDER_CSV=str(_get_secret("GID_ORDER_CSV", "1875928148")),
            LOCATION_SHEET_TITLE=str(_get_secret("LOCATION_SHEET_TITLE", "REFR: Values")),
            LOCATION_NAMED_RANGE=str(_get_secret("LOCATION_NAMED_RANGE", "_locations")),
            TIMESTAMP_TZ=str(_get_secret("TIMESTAMP_TZ", "America/Chicago")),
            TIMESTAMP_FMT=str(_get_secret("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p")),

            TO_RECIPIENTS=_coerce_csv(_get_secret("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_coerce_csv(_get_secret("CC_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_coerce_bool(_get_secret("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=[s.upper() for s in _coerce_csv(_get_secret("REPORT_KEY_RUN_LIST", ""))],
            REPORT_KEY_RECIPIENTS=_coerce_json(_get_secret("REPORT_KEY_RECIPIENTS", "{}")),
            DEFAULT_ORDER_RECIPIENTS=_coerce_csv(_get_secret("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_coerce_bool(
                _get_secret("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_coerce_bool(
                _get_secret("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),

            EMAIL_MANAGER_REPORT=_coerce_bool(
                _get_secret("EMAIL_MANAGER_REPORT", "true")
            ),

            SCOPES=_coerce_csv(
                _get_secret(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,"
                    "https://www.googleapis.com/auth/spreadsheets,"
                    "https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_coerce_bool(_get_secret("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(str(_get_secret("REDIRECT_PORT", "58285")) or "58285"),
            HTTP_TIMEOUT_SECONDS=int(str(_get_secret("HTTP_TIMEOUT_SECONDS", "300")) or "300"),
        )

        # ---------- 2) Overlay from Drive JSON config (optional) ----------
        # If present, read a JSON file on Drive and apply known keys over cfg.
        try:
            import streamlit as st
            from favtrip.google_client import load_valid_token, services
            from favtrip.config_store import load_config_from_drive

            CONFIG_FILE_ID = ""
            if hasattr(st, "secrets"):
                CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

            creds = load_valid_token(cfg.SCOPES)
            if creds:
                _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
                overrides = load_config_from_drive(drive, CONFIG_FILE_ID or None)  # {} if not found
                if isinstance(overrides, dict) and overrides:
                    for k, v in overrides.items():
                        if hasattr(cfg, k):
                            setattr(cfg, k, v)
        except Exception:
            # Fail-open: if Drive/token not ready yet, just return base cfg
            pass

        return cfg

    # -----------------------------------------------------------------------------
    # .env serialization (still useful if you let users "Update defaults in .env")
    # -----------------------------------------------------------------------------

    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")

    # -----------------------------------------------------------------------------
    # .env serialization (still useful if you let users "Update defaults in .env")
    # -----------------------------------------------------------------------------

    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")
```

---
### file: favtrip/config_store.py

```python
# favtrip/config_store.py
from __future__ import annotations
import io
import json
from typing import Any, Dict, Optional
from googleapiclient.discovery import Resource
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

DEFAULT_CONFIG_FILENAME = "favtrip_config.json"
DEFAULT_MIMETYPE = "application/json"

def load_config_from_drive(drive: Resource, file_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Read the JSON config stored in Google Drive.
    If file_id is None, try to discover the newest file named DEFAULT_CONFIG_FILENAME.
    Returns {} if the file doesn't exist or is empty/invalid JSON.
    """
    # Discover by name if a specific id wasn't provided
    if not file_id:
        resp = drive.files().list(
            q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
            orderBy="modifiedTime desc",
            pageSize=1,
            fields="files(id,name,modifiedTime)"
        ).execute() or {}
        files = resp.get("files", [])
        if not files:
            return {}
        file_id = files[0]["id"]

    # Stream download the file
    request = drive.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    raw = buf.getvalue().decode("utf-8", errors="replace").strip()
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except Exception:
        return {}

def save_config_to_drive(
    drive: Resource,
    data: Dict[str, Any],
    file_id: Optional[str] = None,
    parent_folder_id: Optional[str] = None
) -> str:
    """
    Write JSON config to Google Drive.
    - If file_id provided, update that file.
    - Else upsert (update if found by name, otherwise create) DEFAULT_CONFIG_FILENAME.
    Returns the Drive file ID.
    """
    payload = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(payload), mimetype=DEFAULT_MIMETYPE, resumable=True)

    if file_id:
        updated = drive.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        return updated["id"]

    # Try to find an existing file by name to update
    resp = drive.files().list(
        q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
        orderBy="modifiedTime desc",
        pageSize=1,
        fields="files(id,name)"
    ).execute() or {}
    files = resp.get("files", [])
    if files:
        fid = files[0]["id"]
        updated = drive.files().update(fileId=fid, media_body=media).execute()
        return updated["id"]

    # Create a new file
    meta = {"name": DEFAULT_CONFIG_FILENAME}
    if parent_folder_id:
        meta["parents"] = [parent_folder_id]

    created = drive.files().create(
        body=meta,
        media_body=media,
        fields="id,name"
    ).execute()
    return created["id"]
```

---
### file: favtrip/drive_utils.py

```python
from __future__ import annotations
import io
from googleapiclient.http import MediaIoBaseUpload


def find_latest_sheet(drive_svc, folder_id: str):
    q = (
        f"'{folder_id}' in parents and "
        "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    )
    resp = drive_svc.files().list(
        q=q, orderBy="createdTime desc", pageSize=1,
        fields="files(id,name,createdTime)"
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def upload_to_drive(drive_svc, data: bytes, name: str, mime: str, folder_id: str, to_sheet: bool=False):
    meta = {"name": name, "parents": [folder_id]}
    if to_sheet:
        meta["mimeType"] = "application/vnd.google-apps.spreadsheet"
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
    return drive_svc.files().create(
        body=meta, media_body=media, fields="id,name,mimeType,webViewLink"
    ).execute()

```

---
### file: favtrip/gmail_utils.py

```python
from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {ts} – {location}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Automated")
    msg.add_alternative(
        f"<p>Hi team,</p><p>Manager Report ({location})</p>"
        f"<a href='{pdf_link}'>Backup Link</a>", subtype="html"
    )
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)

```

---
### file: favtrip/google_client.py

```python
from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail
```

---
### file: favtrip/logger.py

```python
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass
```

---
### file: favtrip/pipeline.py

```python
from __future__ import annotations
import csv
import io
import re
import requests
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

from .config import Config
from .google_client import get_credentials, services
from .sheets_utils import (
    delete_sheet, copy_sheet_as, copy_first_sheet_as, refresh_sheets_with_prefix,
    get_value, first_gid
)
from .drive_utils import find_latest_sheet, upload_to_drive
from .gmail_utils import send_email, email_manager_report


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}&gid={gid}"
    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content


def timestamp_now(tz: str, fmt: str) -> str:
    return datetime.now(ZoneInfo(tz)).strftime(fmt)

import re

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _clean_emails(items):
    """
    Accepts a list or a comma-separated string and returns a list of valid emails.
    Trailing commas and blanks are removed. Invalid tokens are dropped silently.
    """
    if items is None:
        return []
    if isinstance(items, str):
        items = [p.strip() for p in items.split(",")]
    return [e for e in (p.strip() for p in items) if e and _EMAIL_RE.match(e)]

def _fallback_recipients(hint, *candidates):
    """
    Return the first non-empty, valid recipient list from the provided candidates.
    If all candidates are empty/invalid, raise a friendly error.
    """
    for c in candidates:
        cleaned = _clean_emails(c)
        if cleaned:
            return cleaned
    # Nothing usable found:
    raise ValueError(
        f"No valid recipients available for: {hint}. "
        f"Please provide at least one email in the UI or .env "
        f"(TO_RECIPIENTS, DEFAULT_ORDER_RECIPIENTS, or per-report-key)."
    )


@dataclass
class RunResult:
    ok: bool
    elapsed_seconds: int
    location: str
    timestamp: str
    manager_pdf_link: str | None
    full_order_link: str | None


def run_pipeline(cfg: Config, logger=None) -> RunResult:
    import time
    start = time.perf_counter()

    if logger:
        logger.info("Authorizing with Google APIs…")
    creds = get_credentials(cfg.SCOPES, cfg.REDIRECT_PORT, cfg.FORCE_REAUTH)
    sheets_svc, drive_svc, gmail_svc = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    if logger:
        logger.info("Google services ready")

    # Step 1: latest incoming
    if logger:
        logger.info("Finding latest incoming spreadsheet…")
    latest = find_latest_sheet(drive_svc, cfg.INCOMING_FOLDER_ID)
    if not latest:
        raise RuntimeError("No incoming report found.")
    new_report_id = latest["id"]
    if logger:
        logger.info(f"Latest incoming: {latest['name']} ({new_report_id})")

    # Step 2: prep calculations workbook
    if logger:
        logger.info("Preparing calculations workbook…")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
    try:
        copy_sheet_as(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", "Last Week")
        if logger:
            logger.info("Copied old 'Current Week' to 'Last Week'")
    except Exception:
        if logger:
            logger.warn("No 'Current Week' sheet exists to copy")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
    copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")
    if logger:
        logger.info("Inserted new 'Current Week' from latest incoming report")

    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ')…")
    refresh_sheets_with_prefix(sheets_svc, cfg.CALC_SPREADSHEET_ID, prefix="REFR: ", logger=logger)

    # Step 3: read location code
    location = get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.LOCATION_NAMED_RANGE)
    ts = timestamp_now(cfg.TIMESTAMP_TZ, cfg.TIMESTAMP_FMT)
    if logger:
        logger.info(f"Location: {location}; Timestamp: {ts}")

    # Step 4A: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_MANAGER_PDF, "pdf")
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 4B: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_ORDER_CSV, "csv")

    # Step 4C: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_{ts}_{location}_FULL.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, "text/csv", cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_{ts}_{location}_FULL.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4D: Create per-report-key outputs and email
    text = master_csv_bytes.decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames:
        raise RuntimeError("CSV has no header.")
    report_col = next((h for h in reader.fieldnames if h.lower() == "report_key"), None)
    if not report_col:
        raise RuntimeError("Report_Key column missing.")
    rows = list(reader)
    groups = {}
    for r in rows:
        key = (r.get(report_col) or "").strip() or "UNASSIGNED"
        groups.setdefault(key, []).append(r)

    for key, key_rows in groups.items():
        if not cfg.USE_ALL_REPORT_KEYS and key.upper() not in (cfg.REPORT_KEY_RUN_LIST or []):
            continue
        csv_buf = io.StringIO()
        w = csv.DictWriter(csv_buf, fieldnames=reader.fieldnames, lineterminator="")
        w.writeheader()
        w.writerows(key_rows)
        key_bytes = csv_buf.getvalue().encode("utf-8")
        tag = clean_tag(key.upper())
        filename = f"Order_Report_{ts}_{location}_{tag}.csv"
        created = upload_to_drive(drive_svc, key_bytes, filename, "text/csv", cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{ts}_{location}_{tag}.pdf"

        # Prefer per‑key recipients; else default; else TO
        candidates = None
        if cfg.REPORT_KEY_RECIPIENTS:
            # The keys in UI are upper-cased; normalize here too
            candidates = cfg.REPORT_KEY_RECIPIENTS.get(tag)

        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.DEFAULT_ORDER_RECIPIENTS,
            cfg.TO_RECIPIENTS,
        )

        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – {tag}"
        msg["From"] = "me"
        msg["To"] = ", ".join(recipients)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi {key} team,\nYour order report is ready. \nGoogle Sheet: {created.get('webViewLink')}\nAttached: {pdfname}\n—Automated")
        msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=pdfname)
        if cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL:
            msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info(f"Emailed {tag}")

    # Step 4E: Send Manager Report (guarded by cfg.EMAIL_MANAGER_REPORT)
    if getattr(cfg, "EMAIL_MANAGER_REPORT", True):
        to_list = _fallback_recipients("Manager Report (TO_RECIPIENTS)", cfg.TO_RECIPIENTS)
        cc_list = _clean_emails(cfg.CC_RECIPIENTS)
        email_manager_report(
            gmail_svc, "me", to_list, cc_list,
            pdf_name, pdf_bytes, manager_link, ts, location
        )
        if logger:
            logger.info("Manager email sent")
    else:
        if logger:
            logger.info("Manager email skipped by configuration (EMAIL_MANAGER_REPORT = False)")

    

    # Step 4F: Send Full Order if needed
    full_link = full_created.get('webViewLink')
    if cfg.SEND_SEPARATE_FULL_ORDER_EMAIL:
        to_full = _fallback_recipients("FULL order", cfg.DEFAULT_ORDER_RECIPIENTS, cfg.TO_RECIPIENTS)
        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – FULL"
        msg["From"] = "me"
        msg["To"] = ", ".join(to_full)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi team,\nFULL order report is ready.\nSheet: {full_link}\nAttached: {full_pdf_name}\n—Automated")
        msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info("FULL order email sent")
    else:
        if logger:
            logger.info("Separate full order email disabled")

    elapsed = int(time.perf_counter() - start)
    if logger:
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        logger.info(f"Run completed in {h:02d}:{m:02d}:{s:02d}")

    return RunResult(True, elapsed, location, ts, manager_link, full_link)

```

---
### file: favtrip/sheets_utils.py

```python
from __future__ import annotations
import random
import time
from typing import Any, Dict, List


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def delete_sheet(svc, spreadsheet_id: str, title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), title)
    if s:
        body = {"requests": [{"deleteSheet": {"sheetId": s["sheetId"]}}]}
        svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def copy_sheet_as(svc, spreadsheet_id: str, src_title: str, new_title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), src_title)
    if not s:
        return None
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=spreadsheet_id,
        sheetId=s["sheetId"],
        body={"destinationSpreadsheetId": spreadsheet_id}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return new_id


def copy_first_sheet_as(svc, src_spreadsheet: str, dest_spreadsheet: str, new_title: str):
    meta = svc.spreadsheets().get(spreadsheetId=src_spreadsheet).execute()
    first_id = meta["sheets"][0]["properties"]["sheetId"]
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=src_spreadsheet,
        sheetId=first_id,
        body={"destinationSpreadsheetId": dest_spreadsheet}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=dest_spreadsheet, body=body).execute()
    return new_id


def refresh_sheets_with_prefix(svc, spreadsheet_id: str, prefix: str = "REFR: ", retries: int = 5, logger=None):
    sheets = list_sheets(svc, spreadsheet_id)
    targets = [s["properties"] for s in sheets if s["properties"]["title"].startswith(prefix)]
    for idx, t in enumerate(targets, start=1):
        body = {"requests": [{
            "findReplace": {
                "find": "=",
                "replacement": "=",
                "includeFormulas": True,
                "sheetId": t["sheetId"]
            }
        }]}
        attempt = 0
        while True:
            try:
                svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                if logger:
                    logger.info(f"[{idx}/{len(targets)}] Recalc OK: {t['title']}")
                break
            except Exception:
                attempt += 1
                if attempt > retries:
                    if logger:
                        logger.warn(f"FAILED recalc for {t['title']}")
                    break
                time.sleep(1 + random.random())


def get_value(svc, spreadsheet_id: str, sheet_title: str, named_range: str) -> str:
    try:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=named_range
        ).execute().get("values", [])
    except Exception:
        vals = []
    if not vals:
        try:
            vals = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_title}'!A1:A"
            ).execute().get("values", [])
        except Exception:
            vals = []
    return vals[0][0] if vals and vals[0] else "UNKNOWN"


def first_gid(svc, spreadsheet_id: str) -> int:
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return meta["sheets"][0]["properties"]["sheetId"]

```

---
### file: last_run.log

```
[2026-03-03 23:16:33] INFO: Authorizing with Google APIs…
[2026-03-03 23:16:33] INFO: Google services ready
[2026-03-03 23:16:33] INFO: Finding latest incoming spreadsheet…
[2026-03-03 23:16:34] INFO: Latest incoming: Testing Sales Report - Week 5 (1kNjmEbljdUIqJUjwh8e2-plfLWRZSGHxMHxL50-2ce4)
[2026-03-03 23:16:34] INFO: Preparing calculations workbook…
[2026-03-03 23:16:42] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 23:16:50] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 23:16:50] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 23:17:17] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 23:17:22] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 23:18:15] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 23:18:17] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-11-18-PM
[2026-03-03 23:18:17] INFO: Exporting Manager Report (PDF)…
[2026-03-03 23:18:21] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/1UILYk_poamIwe6JrFQUGZIz9az9iBsdU/view?usp=drivesdk
[2026-03-03 23:18:21] INFO: Exporting Master Order (CSV)…
[2026-03-03 23:18:34] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1bOcle5eMseFc43TZMnXLqgHR7_DcPs13ntb6PSBians/edit?usp=drivesdk
[2026-03-03 23:18:42] INFO: Emailed COFFEE
[2026-03-03 23:18:43] INFO: Manager email sent
[2026-03-03 23:18:43] INFO: Separate full order email disabled
[2026-03-03 23:18:43] INFO: Run completed in 00:02:09

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: requirements.txt

```text
-r documentation/requirements.txt
```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```

---
### file: ui_streamlit.py

```python
import os
import time
import threading
import streamlit as st


import base64, hashlib, secrets, json
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials


# --- add near the top of ui_streamlit.py (after imports) ---
import base64, hashlib, json, secrets
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials


from favtrip.google_client import load_valid_token, services
from favtrip.config_store import load_config_from_drive

from streamlit.components.v1 import html
from streamlit.components.v1 import html as _html_msg



import re

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _parse_emails(csv_str: str):
    return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

def _invalid_emails(csv_str: str):
    return [e for e in _parse_emails(csv_str) if not EMAIL_RE.match(e)]

def _analyze_rk_rows(rows):
    """
    Validate the 'Per-Report-Key Recipients' editor rows.
    Returns (issues: List[str], preview_lines: List[str], rk_map: Dict[str, List[str]])
    """
    issues, preview, rk_map = [], [], {}
    seen, dupes = set(), set()

    for idx, r in enumerate(rows or [], start=1):
        raw_key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
        emails_csv = r.get("Emails (comma)") or ""
        if not raw_key and not emails_csv:
            # allow a blank template row
            continue

        # uppercase flag
        if raw_key != raw_key.upper():
            issues.append(f"Row {idx}: key '{raw_key}' is not ALL CAPS.")

        # duplicate detection
        if raw_key:
            if raw_key in seen:
                dupes.add(raw_key)
            else:
                seen.add(raw_key)

        # email validation
        bads = _invalid_emails(emails_csv)
        if bads:
            issues.append(f"Row {idx}: invalid emails → {', '.join(bads)}")

        # mapping + preview
        if raw_key:
            emails = _parse_emails(emails_csv)
            if emails:
                rk_map[raw_key] = emails
            preview.append(f"{raw_key} → {', '.join(emails) if emails else emails_csv}")

    if dupes:
        issues.append(f"Duplicate keys detected: {', '.join(sorted(dupes))}")
    return issues, preview, rk_map



def _b64url(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")

def _pkce_pair() -> tuple[str, str]:
    """
    Generate a high-entropy PKCE code_verifier and its S256 code_challenge.
    RFC 7636 requires 43–128 chars; this approach yields a URL-safe value.  [2](https://auth0.com/docs/get-started/authentication-and-authorization-flow/authorization-code-flow-with-pkce)
    """
    verifier = _b64url(secrets.token_bytes(64))        # ~86 chars, URL-safe, no padding
    challenge = _b64url(hashlib.sha256(verifier.encode("ascii")).digest())
    return verifier, challenge

def _redirect_base() -> str:
    """
    Always return a non-empty redirect base that exactly matches your OAuth client's
    Authorized redirect URI. Prefer Secrets; normalize to one trailing slash.
    """
    base = (st.secrets.get("APP_BASE_URL", "") or "").strip()
    if not base:
        # Fallback to request (often available), still normalized
        try:
            base = (st.request.url_root or "").strip()
        except Exception:
            base = ""
    if not base:
        st.error("OAuth redirect base is not set. Define APP_BASE_URL in Secrets.")
        st.stop()
    return base.rstrip("/") + "/"

def start_web_oauth(scopes):
    """
    Build an authorization URL that:
      - uses a stable redirect_uri (from Secrets)
      - uses explicit PKCE (S256)
      - embeds the code_verifier inside the state (base64url(JSON))
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    redirect = _redirect_base()

    # Explicit PKCE (stateless across redirect)
    code_verifier, code_challenge = _pkce_pair()

    # CSRF token + verifier encoded into state that Google will return unchanged.
    state_obj = {
        "csrf": _b64url(secrets.token_bytes(16)),
        "v": code_verifier,
        "r": redirect,
    }
    state_b64 = _b64url(json.dumps(state_obj).encode("utf-8"))

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
        state=state_b64,
        # PKCE parameters:
        code_challenge=code_challenge,
        code_challenge_method="S256",
    )

    # Keep only minimal context; state carries the verifier.
    st.session_state["_oauth_redirect"] = redirect
    return auth_url

def _parse_state(state_b64: str) -> dict:
    # Add padding back for base64 decoding if needed
    padding = "=" * ((4 - len(state_b64) % 4) % 4)
    raw = base64.urlsafe_b64decode(state_b64 + padding)
    return json.loads(raw.decode("utf-8"))

def finish_web_oauth(code: str, state_b64: str, scopes):
    """
    Recreate a Flow with the same redirect_uri and exchange code + code_verifier for tokens.
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    state_obj = _parse_state(state_b64)
    code_verifier = state_obj.get("v")
    redirect = state_obj.get("r") or st.session_state.get("_oauth_redirect") or _redirect_base()

    if not code_verifier:
        st.error("OAuth state did not include a PKCE code_verifier.")
        st.stop()

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    # IMPORTANT: include code_verifier here to satisfy PKCE.  [1](https://github.com/googleapis/google-auth-library-python-oauthlib/issues/46)
    flow.fetch_token(code=code, code_verifier=code_verifier)

    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


from favtrip.config_store import save_config_to_drive
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.google_client import (
    start_oauth,
    finish_oauth,
    load_valid_token,
    clear_token,
    services,
)

def _rerun():
    # Works on Streamlit >= 1.27 (st.rerun) and older (experimental_rerun)
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

st.set_page_config(page_title="FavTrip Reporting Pipeline", page_icon="🧾", layout="wide")


# --- Larger Run button ---
# --- REPLACE existing CSS block with this compact style ---
# --- ADD / MERGE THIS CSS FOR THE RUN BUTTON ---

st.markdown(
    """
    <style>
    
    /* Right-align the run button in its column */
    .ft-runwrap {
        display: flex;
        justify-content: flex-end;
        width: 100%;
    }

    /* Target ONLY the submit button used by st.form */
    div[data-testid="stFormSubmitButton"] button {
        background-color: #1a73e8 !important;   /* BLUE */
        color: white !important;                /* WHITE TEXT */
        font-size: 1.20rem !important;          /* BIGGER TEXT */
        font-weight: 600 !important;            /* BOLD */
        padding: 0.90rem 2.0rem !important;     /* LONGER BUTTON */
        border-radius: 10px !important;         /* SOFT CORNERS */
        border: none !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15);
        width: 100%;                            /* STRETCH INSIDE BUTTON WRAP */
        text-align: center;
    }

    /* Hover and active states */
    div[data-testid="stFormSubmitButton"] button:hover {
        filter: brightness(0.93);
    }
    div[data-testid="stFormSubmitButton"] button:active {
        transform: translateY(1px);
    }

    </style>
    """,
    unsafe_allow_html=True
)

# --- END CSS ---
# --- END ADD / MERGE ---

st.title("🧾 FavTrip Reporting Pipeline")

cfg = Config.load()

#Old Config Diagonistics
def hide():
    with st.expander("🔎 Config diagnostics", expanded=False):
        try:
            CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
            st.write(f"CONFIG_FILE_ID (Secrets): `{CONFIG_FILE_ID or '(unset)'}`")

            token = load_valid_token(cfg.SCOPES)
            if not token:
                st.warning("No valid Google token yet. Drive overlay hasn't run.")
            else:
                _sheets, drive, _gmail = services(token, cfg.HTTP_TIMEOUT_SECONDS)
                overrides = load_config_from_drive(drive, CONFIG_FILE_ID or None)
                if overrides:
                    st.success(f"Loaded Drive overrides ({len(overrides)} keys).")
                    # show just the most relevant keys
                    show = {k: overrides.get(k) for k in [
                        "CALC_SPREADSHEET_ID", "INCOMING_FOLDER_ID",
                        "MANAGER_REPORT_FOLDER_ID", "ORDER_REPORT_FOLDER_ID",
                        "TO_RECIPIENTS", "REPORT_KEY_RUN_LIST"
                    ]}
                    st.json(show)
                else:
                    st.error("No Drive config found (or file empty).")
                    st.caption("If you just saved defaults to Drive, copy its file id into Secrets as CONFIG_FILE_ID and rerun.")
        except Exception as e:
            st.error(f"Diagnostics failed: {e}")

    st.caption("**Effective config (current page state):**")
    colA, colB = st.columns(2)
    with colA:
        st.write("**IDs**")
        st.code({
            "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
            "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
            "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
            "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,
        }, language="json")
        st.write("**Recipients**")
        st.code({
            "TO_RECIPIENTS": cfg.TO_RECIPIENTS,
            "CC_RECIPIENTS": cfg.CC_RECIPIENTS,
            "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,
        }, language="json")
    with colB:
        st.write("**Flags & lists**")
        st.code({
            "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
            "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,
            "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,
        }, language="json")
        st.write("**GIDs / TZ**")
        st.code({
            "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
            "GID_ORDER_CSV": cfg.GID_ORDER_CSV,
            "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
            "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,
            "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
            "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,
        }, language="json")
    
params = st.query_params
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        st.success("✅ Google authentication complete.")
        st.query_params.clear()
        # NEW: tell the opener to refresh and then close the popup
        html(
            """
            <script>
              try {
                if (window.opener && !window.opener.closed) {
                  window.opener.postMessage({type: "favtrip_oauth_done"}, "*");
                }
              } catch (e) {}
              // Close this popup
              window.close();
            </script>
            """,
            height=0,
        )
        # Fallback: in case the browser blocks close(), allow manual continue
        st.caption("You can close this window if it didn't close automatically.")
    except Exception as e:
        st.error(f"OAuth error: {e}")
if load_valid_token(cfg.SCOPES):
    cfg = Config.load()

# ----------------------------
# Session state: auth gating
# ----------------------------
if "auth_checked" not in st.session_state:
    # On first load: if token is missing/invalid/unrefreshable -> require auth
    st.session_state.auth_required = (load_valid_token(cfg.SCOPES) is None)
    st.session_state.oauth_flow = None
    st.session_state.oauth_url = None
    st.session_state.auth_checked = True


if "offer_log_download" not in st.session_state:
    st.session_state.offer_log_download = False


# Sidebar controls (always visible)
with st.sidebar:
    st.header("Utilities")

    if st.button("Google Sign Out", type="secondary", use_container_width=True):
        clear_token()
        for key in ["auth_required", "oauth_flow", "oauth_url", "auth_checked"]:
            if key in st.session_state:
                del st.session_state[key]
        _rerun()
        
    st.checkbox(
        "Offer full log download after completion",
        key="offer_log_download",
        help="If enabled, a 'Download last_run.log' button appears when a run finishes.")


# ----------------------------
# Authentication panel (shown only if auth required)
# ----------------------------

# --- Authentication expander (shown only if auth required) ---
if st.session_state.auth_required:
    with st.expander("Google Authentication", expanded=True):
        st.caption(
            "Authentication is required before running. "
            "Click **Sign in with Google** to open the consent screen (it will open in a new tab)."
        )

        # Use a placeholder so we can remove the button immediately
        sign_in_ph = st.empty()

        # Render the button inside the placeholder
        clicked = sign_in_ph.button("Sign in with Google", type="primary", use_container_width=True)

        if clicked:
            try:
                auth_url = start_web_oauth(cfg.SCOPES)

                # 1) Clear the placeholder to remove the button immediately
                sign_in_ph.empty()

                # 2) Show the message in the current tab
                st.markdown(
                    """
                    <div style="
                        display:flex;align-items:center;justify-content:center;
                        height:55vh;text-align:center;
                        font-family: system-ui, Segoe UI, Roboto, Helvetica, Arial, sans-serif;">
                      <div>
                        <h2 style="margin-bottom:0.5rem;">You're being signed in…</h2>
                        <p style="font-size:1.05rem;opacity:.9;">
                          A new browser tab was opened for Google sign‑in.<br/>
                          <strong>After it completes, you may close this tab.</strong>
                        </p>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # (Optional) If user returns to this tab later, refresh to show signed-in state
                html(
                    """
                    <script>
                      document.addEventListener("visibilitychange", function() {
                        if (!document.hidden) { location.reload(); }
                      });
                    </script>
                    """,
                    height=0,
                )

                # 3) Open Google auth in a NEW tab (leave this tab on the message)
                html(
                    f"""
                    <script>
                      window.open({json.dumps(auth_url)}, "_blank", "noopener");
                    </script>
                    """,
                    height=0,
                )

                # 4) Stop further rendering so the message persists
                st.stop()

            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")

        with st.expander("Having trouble?", expanded=False):
            st.write(
                "- The Google authorization page opens in a **new browser tab**.\n"
                "- After completing consent, **close this tab** and use the new tab.\n"
                "- If you renamed your Streamlit app or URL, ensure the Google OAuth "
                "Authorized redirect URI matches exactly (including trailing slash)."
            )

# ----------------------------
# Only show Run Options if NOT requiring auth
# ----------------------------
if not st.session_state.auth_required:

    # ---- Run Form (Run button top-right) ----
    with st.form("run_form"):
        
        header_left, header_right = st.columns([1, 0.32])

        with header_left:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")

        with header_right:
            st.markdown('<div class="ft-runwrap">', unsafe_allow_html=True)
            submitted = st.form_submit_button(
                "Run Pipeline",  # NO ICON
                help="Start the pipeline"
            )
            st.markdown('</div>', unsafe_allow_html=True)



        # --- Main options ---
        # --- REPLACE your current recipients/keys/email-behavior sections with this ---

        # ===== Recipients (compact two columns) =====
        with st.container():
            st.markdown("##### Recipients")
            with st.container():
                col1, col2 = st.columns([1, 1])
                with col1:
                    to = st.text_input(
                        "To (comma)", value=",".join(cfg.TO_RECIPIENTS or []),
                        help="Fallback recipients for Manager & Order emails."
                    )
                with col2:
                    cc = st.text_input(
                        "CC (comma)", value=",".join(cfg.CC_RECIPIENTS or []),
                        help="Optional CC added to all emails."
                    )

        # ===== Report Keys (toggle + input aligned) =====
        with st.container():
            st.markdown("##### Report Keys")
            with st.container():
                colk1, colk2 = st.columns([0.45, 1.55])
                with colk1:
                    # Toggle reads better than checkbox for 'all vs selected'
                    use_all = st.toggle(
                        "Use all keys from CSV",
                        value=cfg.USE_ALL_REPORT_KEYS,
                        help="ON: process every key found. OFF: only the keys you list."
                    )
                with colk2:
                    report_keys = st.text_input(
                        "Keys to run (comma)",
                        value=",".join(cfg.REPORT_KEY_RUN_LIST or []),
                        help="Used when 'Use all keys' is OFF. Example: COFFEE,GROCERY"
                    )

        # ===== Email Behavior (two toggles in one row) =====
        with st.container():
            st.markdown("##### Email Behavior")
            cole1, cole2, cole3 = st.columns([1, 1, 1])
            with cole1:
                include_full = st.toggle(
                    "Attach FULL order in each email",
                    value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL
                )
            with cole2:
                send_full = st.toggle(
                    "Send separate FULL order email",
                    value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL
                )
            with cole3:                
                email_mgr = st.toggle(
                    "Email Manager Report",
                    value=getattr(cfg, "EMAIL_MANAGER_REPORT", True),
                    help="When ON, the Manager Report email is sent. When OFF, it is skipped."
                )


        # ===== Put optional editors/switches in cards for light framing =====
        st.markdown('<div class="ft-card">', unsafe_allow_html=True)
        with st.expander("Per‑Report‑Key Recipients (optional)", expanded=False):
            st.caption("Map **REPORT KEY (ALL CAPS)** → **Emails (comma)**.")
            # existing rows construction remains the same...
            rows = []
            if cfg.REPORT_KEY_RECIPIENTS:
                for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({"REPORT KEY (ALL CAPS)": k, "Emails (comma)": ",".join(v or [])})
            else:
                rows = [{"REPORT KEY (ALL CAPS)": "", "Emails (comma)": ""}]
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                use_container_width=True,
                key="rk_editor",
            )

            # Preview + validation (from earlier step)
            rk_issues, rk_preview, rk_map_preview = _analyze_rk_rows(edited_rows)
            if rk_preview:
                with st.expander("Row template preview"):
                    st.code("\n".join(rk_preview), language="text")
            if rk_issues:
                st.warning("Per‑report‑key recipient issues:\n\n- " + "\n- ".join(rk_issues))

        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="ft-card">', unsafe_allow_html=True)
        with st.expander("Advanced (IDs, GIDs, Timezone, Redirect Port)", expanded=False):
            st.markdown("##### Google Drive / Sheets IDs")
            ga1, ga2 = st.columns([1, 1])
            with ga1:
                calc_id = st.text_input("Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID)
                mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID)
            with ga2:
                incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID)
                order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID)

            st.markdown("##### GIDs & Named Ranges")
            gb1, gb2 = st.columns([1, 1])
            with gb1:
                gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF))
                loc_sheet = st.text_input("Location Sheet Title", value=cfg.LOCATION_SHEET_TITLE)
            with gb2:
                gid_order = st.text_input("Order CSV gid", value=str(cfg.GID_ORDER_CSV))
                loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE)

            st.markdown("##### Time & OAuth")
            gc1, gc2 = st.columns([1, 1])
            with gc1:
                tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ)
                tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT)
            with gc2:
                raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                redirect_port = st.number_input(
                    "Redirect Port (0 = auto)",
                    min_value=0, max_value=65535,
                    value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                    help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535."
                )
        st.markdown('</div>', unsafe_allow_html=True)

        save_drive_defaults = st.checkbox("Update defaults", value=False)

    # ----------------------------
    # Submission handling
    # ----------------------------
    def _split_emails(csv_str: str):
        return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

    if submitted:
        # Apply per-run config
        cfg.TO_RECIPIENTS = _split_emails(to)
        cfg.CC_RECIPIENTS = _split_emails(cc)
        cfg.USE_ALL_REPORT_KEYS = use_all
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

        cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
        cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full
        cfg.EMAIL_MANAGER_REPORT = bool(email_mgr)

        cfg.CALC_SPREADSHEET_ID = calc_id
        cfg.INCOMING_FOLDER_ID = incoming_id
        cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
        cfg.ORDER_REPORT_FOLDER_ID = order_folder
        cfg.REDIRECT_PORT = int(redirect_port)

        cfg.GID_MANAGER_PDF = gid_mgr
        cfg.GID_ORDER_CSV = gid_order
        cfg.LOCATION_SHEET_TITLE = loc_sheet
        cfg.LOCATION_NAMED_RANGE = loc_range
        cfg.TIMESTAMP_TZ = tz
        cfg.TIMESTAMP_FMT = tfmt

        # Per-key recipients from editor
        
        rk_map = {}
        for r in edited_rows:
            key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
            emails_csv = r.get("Emails (comma)") or ""
            emails = _parse_emails(emails_csv)
            if key and emails:
                rk_map[key] = emails
        cfg.REPORT_KEY_RECIPIENTS = rk_map

        # --- ADD: warnings before kicking off the run ---
        # 1) Warn if use_all is OFF and no explicit keys provided
        requested_keys = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]
        if not use_all and not requested_keys:
            st.warning(
                "You left **Use all Report Keys** OFF but provided **no keys** to run. "
                "No per‑key outputs will be generated unless you add keys.",
                icon="⚠️"
            )

        # 2) Warn if no recipients anywhere (TO, DEFAULT_ORDER, or per‑key map)
        any_to = bool(_parse_emails(to) or (cfg.TO_RECIPIENTS or []))
        any_default = bool(cfg.DEFAULT_ORDER_RECIPIENTS or [])
        any_per_key = bool(cfg.REPORT_KEY_RECIPIENTS)
        if not (any_to or any_default or any_per_key):
            st.warning(
                "No recipients are defined: **TO**, **DEFAULT_ORDER_RECIPIENTS**, and **Per‑Report‑Key** are all empty. "
                "Emails will not be sent.",
                icon="⚠️"
            )

        # 3) Surface per-key table issues detected earlier
        if rk_issues:
            st.warning(
                "Per‑report‑key recipient issues detected above. These may prevent emails from sending correctly:\n\n- "
                + "\n- ".join(rk_issues),
                icon="⚠️"
            )
        # --- END ADD ---


        # --- Save edited defaults to Drive JSON (optional) ---
        if save_drive_defaults:
            try:
                # Ensure we have a user token first
                creds = load_valid_token(cfg.SCOPES)
                if not creds:
                    st.error("Not authenticated. Please complete Google sign‑in first (top of page).")
                else:
                    # Drive service
                    _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                    # What we persist (the fields you asked to move out of Secrets)
                    drive_defaults = {
                        "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
                        "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
                        "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
                        "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,

                        "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
                        "GID_ORDER_CSV": cfg.GID_ORDER_CSV,

                        "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
                        "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,

                        "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
                        "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,

                        "TO_RECIPIENTS": cfg.TO_RECIPIENTS,   # lists are fine; JSON keeps types
                        "CC_RECIPIENTS": cfg.CC_RECIPIENTS,

                        "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
                        "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,

                        "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,

                        "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,

                        "INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL": cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL,
                        "SEND_SEPARATE_FULL_ORDER_EMAIL": cfg.SEND_SEPARATE_FULL_ORDER_EMAIL,
                        "EMAIL_MANAGER_REPORT": cfg.EMAIL_MANAGER_REPORT
                    }

                    # If you have CONFIG_FILE_ID in Secrets, we update that exact file.
                    # Otherwise we'll upsert a file named 'favtrip_config.json' and return its id.
                    CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
                    new_id = save_config_to_drive(
                        drive,
                        drive_defaults,
                        file_id=CONFIG_FILE_ID or None,   # update/pin if set, else upsert by name
                        # parent_folder_id=None,          # optional: set a folder id to create under
                    )

                    st.success(f"Saved defaults to Drive config (file id: {new_id}).")
                    if not CONFIG_FILE_ID:
                        st.info(
                            "Tip: add this ID to Streamlit Secrets as `CONFIG_FILE_ID` to pin the same file for all runs:\n"
                            f"`{new_id}`"
                        )
            except Exception as e:
                st.error(f"Failed to save defaults to Drive: {e}")

        # If user checked "Force Google re-auth for this run", kick them into auth gating first.
        if cfg.FORCE_REAUTH:
            clear_token()
            try:
                flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.session_state.oauth_flow = flow
                st.session_state.oauth_url = url
                st.session_state.auth_required = True
                st.info("Re-auth required for this run. Open the URL shown in the Authentication panel.")
                _rerun()
            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")
        else:
            # --- Live run with timer + last log (no full log after completion) ---
            # Write all logs to last_run.log; overwrite on each run
            logger = StatusLogger(print_to_console=True, file_path="last_run.log", overwrite=True)
            result_holder = {"value": None, "error": None}

            def _runner():
                try:
                    result_holder["value"] = run_pipeline(cfg, logger=logger)
                # Catch BaseException so SystemExit / KeyboardInterrupt are captured too
                except BaseException as e:
                    result_holder["error"] = e

            t0 = time.perf_counter()
            th = threading.Thread(target=_runner, daemon=True)
            th.start()

            with st.status("Running pipeline…", expanded=True) as status:
                timer_ph = st.empty()
                lastlog_ph = st.empty()
                while th.is_alive():
                    elapsed = int(time.perf_counter() - t0)
                    timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                    lastlog_ph.markdown(f"**Last:** {logger.last_line()}")
                    time.sleep(0.5)

                th.join()
                elapsed = int(time.perf_counter() - t0)
                timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                lastlog_ph.markdown(f"**Last:** {logger.last_line()}")

                if result_holder["error"]:
                    st.error(f"Run failed: {result_holder['error']}")
                    # Optional during debugging: show stack trace (remove later for a cleaner UI)
                    try:
                        st.exception(result_holder["error"])
                    except Exception:
                        pass
                    status.update(label="❌ Failed", state="error")
                else:
                    result = result_holder["value"]
                    if result is None:
                        st.error("Run finished without returning a result. Check logs and inputs (IDs, Drive access).")
                        status.update(label="⚠️ No result", state="error")
                    else:
                        st.write("### Outputs")
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Location", result.location)
                        col2.metric("Timestamp", result.timestamp)
                        mm = result.elapsed_seconds
                        col3.metric("Elapsed", f"{mm//3600:02d}:{(mm%3600)//60:02d}:{mm%60:02d}")

                        if getattr(result, "manager_pdf_link", None):
                            st.success(f"Manager PDF: {result.manager_pdf_link}")
                        if getattr(result, "full_order_link", None):
                            st.success(f"Full Order Sheet: {result.full_order_link}")

                            
                        if st.session_state.offer_log_download and os.path.exists("last_run.log"):
                            with open("last_run.log", "rb") as f:
                                st.download_button(
                                    "⬇️ Download full log (last_run.log)",
                                    f,
                                    file_name=f"last_run_{result.timestamp}.log",
                                    mime="text/plain",
                                    use_container_width=True
                                )
                            

                        status.update(label="✅ Completed", state="complete")
```

---
### file: web_url_credentials.json

```json
{"web":{"client_id":"674901450584-i6ncesf794ecg6u7olebqkag8gavbuc3.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["https://favtripreporting-dev1.streamlit.app/"]}}
```

```

---
### file: documentation/README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: documentation/generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd "C:\Users\rjrul\OneDrive - University of Iowa\000 Current Semester\004 BAIS 4150 BAIS Capstone\favtrip_reporting"
#python documentation/generate_code_bundle.py
```

---
### file: documentation/git_workflow.txt

```text

---
1. Create a dev branch

# Switch to dev branch
    git checkout dev

# Push dev branch to GitHub (the -u flag sets upstream tracking)
    git push -u origin dev

---
2. Do all development on dev
Make changes normally, then stage, commit, and push:

    git add .
    git commit -m "your message here"
    git push

All pushes go to dev (not main).

---
3. Push dev → main when ready (production release)
Use a Pull Request on GitHub:
1. Go to the repo on GitHub.
2. You will see a banner offering “Compare & Pull Request” (dev → main).
3. Open the Pull Request.
4. Review and click “Merge Pull Request”.

After merging, update your local main branch:

    git checkout main
    git pull

---
4. Keep dev updated with the latest main
After merging into main, update dev:

    git checkout dev
    git pull origin main

Or:

    git merge main

This prevents dev from drifting behind main.

---
5. Summary
- main = production, always stable
- dev = development, experimental work
- Merge dev → main only after testing

```

---
### file: documentation/requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit
openpyxl
```

---
### file: favtrip/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: favtrip/config.py

```python
from __future__ import annotations

import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv
from typing import Any, Dict

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: str = ""):
    """
    Prefer Streamlit secrets (typed TOML values), else fall back to environment.
    Returns the raw value from st.secrets (could be bool/list/dict/str) or a str from env.
    """
    try:
        import streamlit as st  # imported lazily so local CLI still works
        val = st.secrets.get(key, None)
        if val is None:
            return os.getenv(key, default)
        return val
    except Exception:
        return os.getenv(key, default)


_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _coerce_bool(v, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default


def _coerce_csv(v) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]


def _coerce_json(v):
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}


# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True
    EMAIL_MANAGER_REPORT: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300


import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv
from typing import Any, Dict

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------


def _get_secret(key: str, default: Any = None) -> Any:
    """
    Read from Streamlit secrets if present, else env var, else default.
    Does not raise if key missing; returns `default`.
    """
    # Try Streamlit secrets (if running inside Streamlit with secrets configured)
    try:
        import streamlit as st  # imported lazily to avoid hard dependency
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets.get(key, default)
    except Exception:
        pass

    # Fallback to environment variables (including those loaded by dotenv)
    val = os.getenv(key, default)
    return val



_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _coerce_bool(v, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default


def _coerce_csv(v) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]


def _coerce_json(v):
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}


# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True
    EMAIL_MANAGER_REPORT: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        """
        Load config from Streamlit secrets (preferred on cloud) or from .env (local dev),
        then overlay any values found in a Drive-backed JSON config (optional).
        Secrets may be typed (bool/list/dict), so we coerce safely.
        """
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)

        # ---------- 1) Base config from Secrets / .env ----------
        cfg = Config(
            CALC_SPREADSHEET_ID=str(_get_secret("CALC_SPREADSHEET_ID", "")),
            INCOMING_FOLDER_ID=str(_get_secret("INCOMING_FOLDER_ID", "")),
            MANAGER_REPORT_FOLDER_ID=str(_get_secret("MANAGER_REPORT_FOLDER_ID", "")),
            ORDER_REPORT_FOLDER_ID=str(_get_secret("ORDER_REPORT_FOLDER_ID", "")),

            GID_MANAGER_PDF=str(_get_secret("GID_MANAGER_PDF", "1921812573")),
            GID_ORDER_CSV=str(_get_secret("GID_ORDER_CSV", "1875928148")),
            LOCATION_SHEET_TITLE=str(_get_secret("LOCATION_SHEET_TITLE", "REFR: Values")),
            LOCATION_NAMED_RANGE=str(_get_secret("LOCATION_NAMED_RANGE", "_locations")),
            TIMESTAMP_TZ=str(_get_secret("TIMESTAMP_TZ", "America/Chicago")),
            TIMESTAMP_FMT=str(_get_secret("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p")),

            TO_RECIPIENTS=_coerce_csv(_get_secret("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_coerce_csv(_get_secret("CC_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_coerce_bool(_get_secret("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=[s.upper() for s in _coerce_csv(_get_secret("REPORT_KEY_RUN_LIST", ""))],
            REPORT_KEY_RECIPIENTS=_coerce_json(_get_secret("REPORT_KEY_RECIPIENTS", "{}")),
            DEFAULT_ORDER_RECIPIENTS=_coerce_csv(_get_secret("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_coerce_bool(
                _get_secret("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_coerce_bool(
                _get_secret("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),

            EMAIL_MANAGER_REPORT=_coerce_bool(
                _get_secret("EMAIL_MANAGER_REPORT", "true")
            ),

            SCOPES=_coerce_csv(
                _get_secret(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,"
                    "https://www.googleapis.com/auth/spreadsheets,"
                    "https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_coerce_bool(_get_secret("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(str(_get_secret("REDIRECT_PORT", "58285")) or "58285"),
            HTTP_TIMEOUT_SECONDS=int(str(_get_secret("HTTP_TIMEOUT_SECONDS", "300")) or "300"),
        )

        # ---------- 2) Overlay from Drive JSON config (optional) ----------
        # If present, read a JSON file on Drive and apply known keys over cfg.
        try:
            import streamlit as st
            from favtrip.google_client import load_valid_token, services
            from favtrip.config_store import load_config_from_drive

            CONFIG_FILE_ID = ""
            if hasattr(st, "secrets"):
                CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

            creds = load_valid_token(cfg.SCOPES)
            if creds:
                _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
                overrides = load_config_from_drive(drive, CONFIG_FILE_ID or None)  # {} if not found
                if isinstance(overrides, dict) and overrides:
                    for k, v in overrides.items():
                        if hasattr(cfg, k):
                            setattr(cfg, k, v)
        except Exception:
            # Fail-open: if Drive/token not ready yet, just return base cfg
            pass

        return cfg

    # -----------------------------------------------------------------------------
    # .env serialization (still useful if you let users "Update defaults in .env")
    # -----------------------------------------------------------------------------

    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")

    # -----------------------------------------------------------------------------
    # .env serialization (still useful if you let users "Update defaults in .env")
    # -----------------------------------------------------------------------------

    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")
```

---
### file: favtrip/config_store.py

```python
# favtrip/config_store.py
from __future__ import annotations
import io
import json
from typing import Any, Dict, Optional
from googleapiclient.discovery import Resource
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

DEFAULT_CONFIG_FILENAME = "favtrip_config.json"
DEFAULT_MIMETYPE = "application/json"

def load_config_from_drive(drive: Resource, file_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Read the JSON config stored in Google Drive.
    If file_id is None, try to discover the newest file named DEFAULT_CONFIG_FILENAME.
    Returns {} if the file doesn't exist or is empty/invalid JSON.
    """
    # Discover by name if a specific id wasn't provided
    if not file_id:
        resp = drive.files().list(
            q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
            orderBy="modifiedTime desc",
            pageSize=1,
            fields="files(id,name,modifiedTime)"
        ).execute() or {}
        files = resp.get("files", [])
        if not files:
            return {}
        file_id = files[0]["id"]

    # Stream download the file
    request = drive.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    raw = buf.getvalue().decode("utf-8", errors="replace").strip()
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except Exception:
        return {}

def save_config_to_drive(
    drive: Resource,
    data: Dict[str, Any],
    file_id: Optional[str] = None,
    parent_folder_id: Optional[str] = None
) -> str:
    """
    Write JSON config to Google Drive.
    - If file_id provided, update that file.
    - Else upsert (update if found by name, otherwise create) DEFAULT_CONFIG_FILENAME.
    Returns the Drive file ID.
    """
    payload = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(payload), mimetype=DEFAULT_MIMETYPE, resumable=True)

    if file_id:
        updated = drive.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        return updated["id"]

    # Try to find an existing file by name to update
    resp = drive.files().list(
        q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
        orderBy="modifiedTime desc",
        pageSize=1,
        fields="files(id,name)"
    ).execute() or {}
    files = resp.get("files", [])
    if files:
        fid = files[0]["id"]
        updated = drive.files().update(fileId=fid, media_body=media).execute()
        return updated["id"]

    # Create a new file
    meta = {"name": DEFAULT_CONFIG_FILENAME}
    if parent_folder_id:
        meta["parents"] = [parent_folder_id]

    created = drive.files().create(
        body=meta,
        media_body=media,
        fields="id,name"
    ).execute()
    return created["id"]
```

---
### file: favtrip/drive_utils.py

```python
from __future__ import annotations
import io
from googleapiclient.http import MediaIoBaseUpload


def find_latest_sheet(drive_svc, folder_id: str):
    q = (
        f"'{folder_id}' in parents and "
        "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    )
    resp = drive_svc.files().list(
        q=q, orderBy="createdTime desc", pageSize=1,
        fields="files(id,name,createdTime)"
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def upload_to_drive(drive_svc, data: bytes, name: str, mime: str, folder_id: str, to_sheet: bool=False):
    meta = {"name": name, "parents": [folder_id]}
    if to_sheet:
        meta["mimeType"] = "application/vnd.google-apps.spreadsheet"
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
    return drive_svc.files().create(
        body=meta, media_body=media, fields="id,name,mimeType,webViewLink"
    ).execute()

```

---
### file: favtrip/gmail_utils.py

```python
from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {ts} – {location}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Automated")
    msg.add_alternative(
        f"<p>Hi team,</p><p>Manager Report ({location})</p>"
        f"<a href='{pdf_link}'>Backup Link</a>", subtype="html"
    )
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)

```

---
### file: favtrip/google_client.py

```python
from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail
```

---
### file: favtrip/logger.py

```python
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass
```

---
### file: favtrip/pipeline.py

```python
from __future__ import annotations
import csv
import io
import re
import requests
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

from io import BytesIO
from openpyxl import load_workbook, Workbook


from .config import Config
from .google_client import get_credentials, services
from .sheets_utils import (
    delete_sheet, copy_sheet_as, copy_first_sheet_as, refresh_sheets_with_prefix,
    get_value, first_gid
)
from .drive_utils import find_latest_sheet, upload_to_drive
from .gmail_utils import send_email, email_manager_report

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}&gid={gid}"
    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content


def timestamp_now(tz: str, fmt: str) -> str:
    return datetime.now(ZoneInfo(tz)).strftime(fmt)

import re

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _clean_emails(items):
    """
    Accepts a list or a comma-separated string and returns a list of valid emails.
    Trailing commas and blanks are removed. Invalid tokens are dropped silently.
    """
    if items is None:
        return []
    if isinstance(items, str):
        items = [p.strip() for p in items.split(",")]
    return [e for e in (p.strip() for p in items) if e and _EMAIL_RE.match(e)]

def _fallback_recipients(hint, *candidates):
    """
    Return the first non-empty, valid recipient list from the provided candidates.
    If all candidates are empty/invalid, raise a friendly error.
    """
    for c in candidates:
        cleaned = _clean_emails(c)
        if cleaned:
            return cleaned
    # Nothing usable found:
    raise ValueError(
        f"No valid recipients available for: {hint}. "
        f"Please provide at least one email in the UI or .env "
        f"(TO_RECIPIENTS, DEFAULT_ORDER_RECIPIENTS, or per-report-key)."
    )


@dataclass
class RunResult:
    ok: bool
    elapsed_seconds: int
    location: str
    timestamp: str
    manager_pdf_link: str | None
    full_order_link: str | None


def run_pipeline(cfg: Config, logger=None) -> RunResult:
    import time
    start = time.perf_counter()

    if logger:
        logger.info("Authorizing with Google APIs…")
    creds = get_credentials(cfg.SCOPES, cfg.REDIRECT_PORT, cfg.FORCE_REAUTH)
    sheets_svc, drive_svc, gmail_svc = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    if logger:
        logger.info("Google services ready")

    # Step 1: latest incoming
    if logger:
        logger.info("Finding latest incoming spreadsheet…")
    latest = find_latest_sheet(drive_svc, cfg.INCOMING_FOLDER_ID)
    if not latest:
        raise RuntimeError("No incoming report found.")
    new_report_id = latest["id"]
    if logger:
        logger.info(f"Latest incoming: {latest['name']} ({new_report_id})")

    # Step 2: prep calculations workbook
    if logger:
        logger.info("Preparing calculations workbook…")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
    try:
        copy_sheet_as(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", "Last Week")
        if logger:
            logger.info("Copied old 'Current Week' to 'Last Week'")
    except Exception:
        if logger:
            logger.warn("No 'Current Week' sheet exists to copy")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
    copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")
    if logger:
        logger.info("Inserted new 'Current Week' from latest incoming report")

    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ')…")
    refresh_sheets_with_prefix(sheets_svc, cfg.CALC_SPREADSHEET_ID, prefix="REFR: ", logger=logger)

    # Step 3: read location code
    location = get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.LOCATION_NAMED_RANGE)
    ts = timestamp_now(cfg.TIMESTAMP_TZ, cfg.TIMESTAMP_FMT)
    if logger:
        logger.info(f"Location: {location}; Timestamp: {ts}")

    # Step 4A: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_MANAGER_PDF, "pdf")
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 4B: Master Order XLSX
    if logger:
        logger.info("Exporting Master Order (XLSX)…")
    master_xlsx_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_ORDER_CSV, "xlsx")

    # Step 4C: Full order upload (XLSX) and export (PDF)
    full_xlsx_name = f"Order_Report_{ts}_{location}_FULL.xlsx"
    full_created = upload_to_drive(drive_svc, master_xlsx_bytes, full_xlsx_name, XLSX_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_{ts}_{location}_FULL.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4D: Create per-report-key outputs (XLSX) and email

    # --- Parse the master XLSX into rows of dicts ---
    wb = load_workbook(filename=BytesIO(master_xlsx_bytes), read_only=True, data_only=True)
    ws = wb.active  # or specify a sheet name if needed
    rows_iter = ws.iter_rows(values_only=True)

    # Header row
    headers = next(rows_iter, None)
    if not headers:
        raise RuntimeError("XLSX has no header.")

    headers = [str(h).strip() if h is not None else "" for h in headers]
    report_col = next((h for h in headers if h and h.lower() == "report_key"), None)
    if not report_col:
        raise RuntimeError("Report_Key column missing.")

    # Materialize rows as list[dict]
    rows = []
    for r in rows_iter:
        rows.append({headers[i]: (r[i] if i < len(r) else None) for i in range(len(headers))})

    # Group by report key
    groups: dict[str, list[dict]] = {}
    for r in rows:
        key = (str(r.get(report_col) or "").strip()) or "UNASSIGNED"
        groups.setdefault(key, []).append(r)

    for key, key_rows in groups.items():
        if not cfg.USE_ALL_REPORT_KEYS and key.upper() not in (cfg.REPORT_KEY_RUN_LIST or []):
            continue

        # Build a per-key XLSX in memory
        out_wb = Workbook(write_only=True)
        out_ws = out_wb.create_sheet("Sheet1")
        out_ws.append(headers)
        for rr in key_rows:
            out_ws.append([rr.get(h) for h in headers])

        bio = BytesIO()
        out_wb.save(bio)
        key_xlsx_bytes = bio.getvalue()

        tag = clean_tag(key.upper())
        xlsx_name = f"Order_Report_{ts}_{location}_{tag}.xlsx"

        # Upload XLSX to Drive; set to_sheet=True so Drive converts to Google Sheet (needed for your PDF export + link)
        created = upload_to_drive(drive_svc, key_xlsx_bytes, xlsx_name, XLSX_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True) 
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)

        # Export the Google Sheet as PDF (unchanged behavior)
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{ts}_{location}_{tag}.pdf"

        # Prefer per‑key recipients; else default; else TO
        candidates = None
        if cfg.REPORT_KEY_RECIPIENTS:
            # The keys in UI are upper-cased; normalize here too
            candidates = cfg.REPORT_KEY_RECIPIENTS.get(tag)

        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.DEFAULT_ORDER_RECIPIENTS,
            cfg.TO_RECIPIENTS,
        )

        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – {tag}"
        msg["From"] = "me"
        msg["To"] = ", ".join(recipients)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)

        # Keep the Google Sheet link; PDF remains the attached artifact
        msg.set_content(
            f"Hi {key} team,\nYour order report is ready. \n"
            f"Google Sheet: {created.get('webViewLink')}\n"
            f"Attached: {pdfname}\n—Automated"
        )
        msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=pdfname)
        if cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL:
            msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)

        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info(f"Emailed {tag}")

    # Step 4E: Send Manager Report (guarded by cfg.EMAIL_MANAGER_REPORT)
    if getattr(cfg, "EMAIL_MANAGER_REPORT", True):
        to_list = _fallback_recipients("Manager Report (TO_RECIPIENTS)", cfg.TO_RECIPIENTS)
        cc_list = _clean_emails(cfg.CC_RECIPIENTS)
        email_manager_report(
            gmail_svc, "me", to_list, cc_list,
            pdf_name, pdf_bytes, manager_link, ts, location
        )
        if logger:
            logger.info("Manager email sent")
    else:
        if logger:
            logger.info("Manager email skipped by configuration (EMAIL_MANAGER_REPORT = False)")

    

    # Step 4F: Send Full Order if needed
    full_link = full_created.get('webViewLink')
    if cfg.SEND_SEPARATE_FULL_ORDER_EMAIL:
        to_full = _fallback_recipients("FULL order", cfg.DEFAULT_ORDER_RECIPIENTS, cfg.TO_RECIPIENTS)
        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – FULL"
        msg["From"] = "me"
        msg["To"] = ", ".join(to_full)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi team,\nFULL order report is ready.\nSheet: {full_link}\nAttached: {full_pdf_name}\n—Automated")
        msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info("FULL order email sent")
    else:
        if logger:
            logger.info("Separate full order email disabled")

    elapsed = int(time.perf_counter() - start)
    if logger:
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        logger.info(f"Run completed in {h:02d}:{m:02d}:{s:02d}")

    return RunResult(True, elapsed, location, ts, manager_link, full_link)

```

---
### file: favtrip/sheets_utils.py

```python
from __future__ import annotations
import random
import time
from typing import Any, Dict, List


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def delete_sheet(svc, spreadsheet_id: str, title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), title)
    if s:
        body = {"requests": [{"deleteSheet": {"sheetId": s["sheetId"]}}]}
        svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def copy_sheet_as(svc, spreadsheet_id: str, src_title: str, new_title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), src_title)
    if not s:
        return None
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=spreadsheet_id,
        sheetId=s["sheetId"],
        body={"destinationSpreadsheetId": spreadsheet_id}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return new_id


def copy_first_sheet_as(svc, src_spreadsheet: str, dest_spreadsheet: str, new_title: str):
    meta = svc.spreadsheets().get(spreadsheetId=src_spreadsheet).execute()
    first_id = meta["sheets"][0]["properties"]["sheetId"]
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=src_spreadsheet,
        sheetId=first_id,
        body={"destinationSpreadsheetId": dest_spreadsheet}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=dest_spreadsheet, body=body).execute()
    return new_id


def refresh_sheets_with_prefix(svc, spreadsheet_id: str, prefix: str = "REFR: ", retries: int = 5, logger=None):
    sheets = list_sheets(svc, spreadsheet_id)
    targets = [s["properties"] for s in sheets if s["properties"]["title"].startswith(prefix)]
    for idx, t in enumerate(targets, start=1):
        body = {"requests": [{
            "findReplace": {
                "find": "=",
                "replacement": "=",
                "includeFormulas": True,
                "sheetId": t["sheetId"]
            }
        }]}
        attempt = 0
        while True:
            try:
                svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                if logger:
                    logger.info(f"[{idx}/{len(targets)}] Recalc OK: {t['title']}")
                break
            except Exception:
                attempt += 1
                if attempt > retries:
                    if logger:
                        logger.warn(f"FAILED recalc for {t['title']}")
                    break
                time.sleep(1 + random.random())


def get_value(svc, spreadsheet_id: str, sheet_title: str, named_range: str) -> str:
    try:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=named_range
        ).execute().get("values", [])
    except Exception:
        vals = []
    if not vals:
        try:
            vals = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_title}'!A1:A"
            ).execute().get("values", [])
        except Exception:
            vals = []
    return vals[0][0] if vals and vals[0] else "UNKNOWN"


def first_gid(svc, spreadsheet_id: str) -> int:
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return meta["sheets"][0]["properties"]["sheetId"]

```

---
### file: last_run.log

```
[2026-03-03 23:16:33] INFO: Authorizing with Google APIs…
[2026-03-03 23:16:33] INFO: Google services ready
[2026-03-03 23:16:33] INFO: Finding latest incoming spreadsheet…
[2026-03-03 23:16:34] INFO: Latest incoming: Testing Sales Report - Week 5 (1kNjmEbljdUIqJUjwh8e2-plfLWRZSGHxMHxL50-2ce4)
[2026-03-03 23:16:34] INFO: Preparing calculations workbook…
[2026-03-03 23:16:42] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 23:16:50] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 23:16:50] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 23:17:17] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 23:17:22] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 23:18:15] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 23:18:17] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-11-18-PM
[2026-03-03 23:18:17] INFO: Exporting Manager Report (PDF)…
[2026-03-03 23:18:21] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/1UILYk_poamIwe6JrFQUGZIz9az9iBsdU/view?usp=drivesdk
[2026-03-03 23:18:21] INFO: Exporting Master Order (CSV)…
[2026-03-03 23:18:34] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1bOcle5eMseFc43TZMnXLqgHR7_DcPs13ntb6PSBians/edit?usp=drivesdk
[2026-03-03 23:18:42] INFO: Emailed COFFEE
[2026-03-03 23:18:43] INFO: Manager email sent
[2026-03-03 23:18:43] INFO: Separate full order email disabled
[2026-03-03 23:18:43] INFO: Run completed in 00:02:09

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: requirements.txt

```text
-r documentation/requirements.txt
```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```

---
### file: ui_streamlit.py

```python
import os
import time
import threading
import streamlit as st


import base64, hashlib, secrets, json
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials


# --- add near the top of ui_streamlit.py (after imports) ---
import base64, hashlib, json, secrets
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials


from favtrip.google_client import load_valid_token, services
from favtrip.config_store import load_config_from_drive

from streamlit.components.v1 import html
from streamlit.components.v1 import html as _html_msg



import re

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _parse_emails(csv_str: str):
    return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

def _invalid_emails(csv_str: str):
    return [e for e in _parse_emails(csv_str) if not EMAIL_RE.match(e)]

def _analyze_rk_rows(rows):
    """
    Validate the 'Per-Report-Key Recipients' editor rows.
    Returns (issues: List[str], preview_lines: List[str], rk_map: Dict[str, List[str]])
    """
    issues, preview, rk_map = [], [], {}
    seen, dupes = set(), set()

    for idx, r in enumerate(rows or [], start=1):
        raw_key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
        emails_csv = r.get("Emails (comma)") or ""
        if not raw_key and not emails_csv:
            # allow a blank template row
            continue

        # uppercase flag
        if raw_key != raw_key.upper():
            issues.append(f"Row {idx}: key '{raw_key}' is not ALL CAPS.")

        # duplicate detection
        if raw_key:
            if raw_key in seen:
                dupes.add(raw_key)
            else:
                seen.add(raw_key)

        # email validation
        bads = _invalid_emails(emails_csv)
        if bads:
            issues.append(f"Row {idx}: invalid emails → {', '.join(bads)}")

        # mapping + preview
        if raw_key:
            emails = _parse_emails(emails_csv)
            if emails:
                rk_map[raw_key] = emails
            preview.append(f"{raw_key} → {', '.join(emails) if emails else emails_csv}")

    if dupes:
        issues.append(f"Duplicate keys detected: {', '.join(sorted(dupes))}")
    return issues, preview, rk_map



def _b64url(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")

def _pkce_pair() -> tuple[str, str]:
    """
    Generate a high-entropy PKCE code_verifier and its S256 code_challenge.
    RFC 7636 requires 43–128 chars; this approach yields a URL-safe value.  [2](https://auth0.com/docs/get-started/authentication-and-authorization-flow/authorization-code-flow-with-pkce)
    """
    verifier = _b64url(secrets.token_bytes(64))        # ~86 chars, URL-safe, no padding
    challenge = _b64url(hashlib.sha256(verifier.encode("ascii")).digest())
    return verifier, challenge

def _redirect_base() -> str:
    """
    Always return a non-empty redirect base that exactly matches your OAuth client's
    Authorized redirect URI. Prefer Secrets; normalize to one trailing slash.
    """
    base = (st.secrets.get("APP_BASE_URL", "") or "").strip()
    if not base:
        # Fallback to request (often available), still normalized
        try:
            base = (st.request.url_root or "").strip()
        except Exception:
            base = ""
    if not base:
        st.error("OAuth redirect base is not set. Define APP_BASE_URL in Secrets.")
        st.stop()
    return base.rstrip("/") + "/"

def start_web_oauth(scopes):
    """
    Build an authorization URL that:
      - uses a stable redirect_uri (from Secrets)
      - uses explicit PKCE (S256)
      - embeds the code_verifier inside the state (base64url(JSON))
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    redirect = _redirect_base()

    # Explicit PKCE (stateless across redirect)
    code_verifier, code_challenge = _pkce_pair()

    # CSRF token + verifier encoded into state that Google will return unchanged.
    state_obj = {
        "csrf": _b64url(secrets.token_bytes(16)),
        "v": code_verifier,
        "r": redirect,
    }
    state_b64 = _b64url(json.dumps(state_obj).encode("utf-8"))

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
        state=state_b64,
        # PKCE parameters:
        code_challenge=code_challenge,
        code_challenge_method="S256",
    )

    # Keep only minimal context; state carries the verifier.
    st.session_state["_oauth_redirect"] = redirect
    return auth_url

def _parse_state(state_b64: str) -> dict:
    # Add padding back for base64 decoding if needed
    padding = "=" * ((4 - len(state_b64) % 4) % 4)
    raw = base64.urlsafe_b64decode(state_b64 + padding)
    return json.loads(raw.decode("utf-8"))

def finish_web_oauth(code: str, state_b64: str, scopes):
    """
    Recreate a Flow with the same redirect_uri and exchange code + code_verifier for tokens.
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    state_obj = _parse_state(state_b64)
    code_verifier = state_obj.get("v")
    redirect = state_obj.get("r") or st.session_state.get("_oauth_redirect") or _redirect_base()

    if not code_verifier:
        st.error("OAuth state did not include a PKCE code_verifier.")
        st.stop()

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    # IMPORTANT: include code_verifier here to satisfy PKCE.  [1](https://github.com/googleapis/google-auth-library-python-oauthlib/issues/46)
    flow.fetch_token(code=code, code_verifier=code_verifier)

    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


from favtrip.config_store import save_config_to_drive
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.google_client import (
    start_oauth,
    finish_oauth,
    load_valid_token,
    clear_token,
    services,
)

def _rerun():
    # Works on Streamlit >= 1.27 (st.rerun) and older (experimental_rerun)
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

st.set_page_config(page_title="FavTrip Reporting Pipeline", page_icon="🧾", layout="wide")


# --- Larger Run button ---
# --- REPLACE existing CSS block with this compact style ---
# --- ADD / MERGE THIS CSS FOR THE RUN BUTTON ---

st.markdown(
    """
    <style>
    
    /* Right-align the run button in its column */
    .ft-runwrap {
        display: flex;
        justify-content: flex-end;
        width: 100%;
    }

    /* Target ONLY the submit button used by st.form */
    div[data-testid="stFormSubmitButton"] button {
        background-color: #1a73e8 !important;   /* BLUE */
        color: white !important;                /* WHITE TEXT */
        font-size: 1.20rem !important;          /* BIGGER TEXT */
        font-weight: 600 !important;            /* BOLD */
        padding: 0.90rem 2.0rem !important;     /* LONGER BUTTON */
        border-radius: 10px !important;         /* SOFT CORNERS */
        border: none !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15);
        width: 100%;                            /* STRETCH INSIDE BUTTON WRAP */
        text-align: center;
    }

    /* Hover and active states */
    div[data-testid="stFormSubmitButton"] button:hover {
        filter: brightness(0.93);
    }
    div[data-testid="stFormSubmitButton"] button:active {
        transform: translateY(1px);
    }

    </style>
    """,
    unsafe_allow_html=True
)

# --- END CSS ---
# --- END ADD / MERGE ---

st.title("🧾 FavTrip Reporting Pipeline")

cfg = Config.load()

#Old Config Diagonistics
def hide():
    with st.expander("🔎 Config diagnostics", expanded=False):
        try:
            CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
            st.write(f"CONFIG_FILE_ID (Secrets): `{CONFIG_FILE_ID or '(unset)'}`")

            token = load_valid_token(cfg.SCOPES)
            if not token:
                st.warning("No valid Google token yet. Drive overlay hasn't run.")
            else:
                _sheets, drive, _gmail = services(token, cfg.HTTP_TIMEOUT_SECONDS)
                overrides = load_config_from_drive(drive, CONFIG_FILE_ID or None)
                if overrides:
                    st.success(f"Loaded Drive overrides ({len(overrides)} keys).")
                    # show just the most relevant keys
                    show = {k: overrides.get(k) for k in [
                        "CALC_SPREADSHEET_ID", "INCOMING_FOLDER_ID",
                        "MANAGER_REPORT_FOLDER_ID", "ORDER_REPORT_FOLDER_ID",
                        "TO_RECIPIENTS", "REPORT_KEY_RUN_LIST"
                    ]}
                    st.json(show)
                else:
                    st.error("No Drive config found (or file empty).")
                    st.caption("If you just saved defaults to Drive, copy its file id into Secrets as CONFIG_FILE_ID and rerun.")
        except Exception as e:
            st.error(f"Diagnostics failed: {e}")

    st.caption("**Effective config (current page state):**")
    colA, colB = st.columns(2)
    with colA:
        st.write("**IDs**")
        st.code({
            "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
            "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
            "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
            "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,
        }, language="json")
        st.write("**Recipients**")
        st.code({
            "TO_RECIPIENTS": cfg.TO_RECIPIENTS,
            "CC_RECIPIENTS": cfg.CC_RECIPIENTS,
            "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,
        }, language="json")
    with colB:
        st.write("**Flags & lists**")
        st.code({
            "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
            "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,
            "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,
        }, language="json")
        st.write("**GIDs / TZ**")
        st.code({
            "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
            "GID_ORDER_CSV": cfg.GID_ORDER_CSV,
            "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
            "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,
            "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
            "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,
        }, language="json")
    
params = st.query_params
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        st.success("✅ Google authentication complete.")
        st.query_params.clear()
        # NEW: tell the opener to refresh and then close the popup
        html(
            """
            <script>
              try {
                if (window.opener && !window.opener.closed) {
                  window.opener.postMessage({type: "favtrip_oauth_done"}, "*");
                }
              } catch (e) {}
              // Close this popup
              window.close();
            </script>
            """,
            height=0,
        )
        # Fallback: in case the browser blocks close(), allow manual continue
        st.caption("You can close this window if it didn't close automatically.")
    except Exception as e:
        st.error(f"OAuth error: {e}")
if load_valid_token(cfg.SCOPES):
    cfg = Config.load()

# ----------------------------
# Session state: auth gating
# ----------------------------
if "auth_checked" not in st.session_state:
    # On first load: if token is missing/invalid/unrefreshable -> require auth
    st.session_state.auth_required = (load_valid_token(cfg.SCOPES) is None)
    st.session_state.oauth_flow = None
    st.session_state.oauth_url = None
    st.session_state.auth_checked = True


if "offer_log_download" not in st.session_state:
    st.session_state.offer_log_download = False


# Sidebar controls (always visible)
with st.sidebar:
    st.header("Utilities")

    if st.button("Google Sign Out", type="secondary", use_container_width=True):
        clear_token()
        for key in ["auth_required", "oauth_flow", "oauth_url", "auth_checked"]:
            if key in st.session_state:
                del st.session_state[key]
        _rerun()
        
    st.checkbox(
        "Offer full log download after completion",
        key="offer_log_download",
        help="If enabled, a 'Download last_run.log' button appears when a run finishes.")


# ----------------------------
# Authentication panel (shown only if auth required)
# ----------------------------

# --- Authentication expander (shown only if auth required) ---
if st.session_state.auth_required:
    with st.expander("Google Authentication", expanded=True):
        st.caption(
            "Authentication is required before running. "
            "Click **Sign in with Google** to open the consent screen (it will open in a new tab)."
        )

        # Use a placeholder so we can remove the button immediately
        sign_in_ph = st.empty()

        # Render the button inside the placeholder
        clicked = sign_in_ph.button("Sign in with Google", type="primary", use_container_width=True)

        if clicked:
            try:
                auth_url = start_web_oauth(cfg.SCOPES)

                # 1) Clear the placeholder to remove the button immediately
                sign_in_ph.empty()

                # 2) Show the message in the current tab
                st.markdown(
                    """
                    <div style="
                        display:flex;align-items:center;justify-content:center;
                        height:55vh;text-align:center;
                        font-family: system-ui, Segoe UI, Roboto, Helvetica, Arial, sans-serif;">
                      <div>
                        <h2 style="margin-bottom:0.5rem;">You're being signed in…</h2>
                        <p style="font-size:1.05rem;opacity:.9;">
                          A new browser tab was opened for Google sign‑in.<br/>
                          <strong>After it completes, you may close this tab.</strong>
                        </p>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # (Optional) If user returns to this tab later, refresh to show signed-in state
                html(
                    """
                    <script>
                      document.addEventListener("visibilitychange", function() {
                        if (!document.hidden) { location.reload(); }
                      });
                    </script>
                    """,
                    height=0,
                )

                # 3) Open Google auth in a NEW tab (leave this tab on the message)
                html(
                    f"""
                    <script>
                      window.open({json.dumps(auth_url)}, "_blank", "noopener");
                    </script>
                    """,
                    height=0,
                )

                # 4) Stop further rendering so the message persists
                st.stop()

            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")

        with st.expander("Having trouble?", expanded=False):
            st.write(
                "- The Google authorization page opens in a **new browser tab**.\n"
                "- After completing consent, **close this tab** and use the new tab.\n"
                "- If you renamed your Streamlit app or URL, ensure the Google OAuth "
                "Authorized redirect URI matches exactly (including trailing slash)."
            )

# ----------------------------
# Only show Run Options if NOT requiring auth
# ----------------------------
if not st.session_state.auth_required:

    # ---- Run Form (Run button top-right) ----
    with st.form("run_form"):
        
        header_left, header_right = st.columns([1, 0.32])

        with header_left:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")

        with header_right:
            st.markdown('<div class="ft-runwrap">', unsafe_allow_html=True)
            submitted = st.form_submit_button(
                "Run Pipeline",  # NO ICON
                help="Start the pipeline"
            )
            st.markdown('</div>', unsafe_allow_html=True)



        # --- Main options ---
        # --- REPLACE your current recipients/keys/email-behavior sections with this ---

        # ===== Recipients (compact two columns) =====
        with st.container():
            st.markdown("##### Recipients")
            with st.container():
                col1, col2 = st.columns([1, 1])
                with col1:
                    to = st.text_input(
                        "To (comma)", value=",".join(cfg.TO_RECIPIENTS or []),
                        help="Fallback recipients for Manager & Order emails."
                    )
                with col2:
                    cc = st.text_input(
                        "CC (comma)", value=",".join(cfg.CC_RECIPIENTS or []),
                        help="Optional CC added to all emails."
                    )

        # ===== Report Keys (toggle + input aligned) =====
        with st.container():
            st.markdown("##### Report Keys")
            with st.container():
                colk1, colk2 = st.columns([0.45, 1.55])
                with colk1:
                    # Toggle reads better than checkbox for 'all vs selected'
                    use_all = st.toggle(
                        "Use all keys from CSV",
                        value=cfg.USE_ALL_REPORT_KEYS,
                        help="ON: process every key found. OFF: only the keys you list."
                    )
                with colk2:
                    report_keys = st.text_input(
                        "Keys to run (comma)",
                        value=",".join(cfg.REPORT_KEY_RUN_LIST or []),
                        help="Used when 'Use all keys' is OFF. Example: COFFEE,GROCERY"
                    )

        # ===== Email Behavior (two toggles in one row) =====
        with st.container():
            st.markdown("##### Email Behavior")
            cole1, cole2, cole3 = st.columns([1, 1, 1])
            with cole1:
                include_full = st.toggle(
                    "Attach FULL order in each email",
                    value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL
                )
            with cole2:
                send_full = st.toggle(
                    "Send separate FULL order email",
                    value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL
                )
            with cole3:                
                email_mgr = st.toggle(
                    "Email Manager Report",
                    value=getattr(cfg, "EMAIL_MANAGER_REPORT", True),
                    help="When ON, the Manager Report email is sent. When OFF, it is skipped."
                )


        # ===== Put optional editors/switches in cards for light framing =====
        st.markdown('<div class="ft-card">', unsafe_allow_html=True)
        with st.expander("Per‑Report‑Key Recipients (optional)", expanded=False):
            st.caption("Map **REPORT KEY (ALL CAPS)** → **Emails (comma)**.")
            # existing rows construction remains the same...
            rows = []
            if cfg.REPORT_KEY_RECIPIENTS:
                for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({"REPORT KEY (ALL CAPS)": k, "Emails (comma)": ",".join(v or [])})
            else:
                rows = [{"REPORT KEY (ALL CAPS)": "", "Emails (comma)": ""}]
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                use_container_width=True,
                key="rk_editor",
            )

            # Preview + validation (from earlier step)
            rk_issues, rk_preview, rk_map_preview = _analyze_rk_rows(edited_rows)
            if rk_preview:
                with st.expander("Row template preview"):
                    st.code("\n".join(rk_preview), language="text")
            if rk_issues:
                st.warning("Per‑report‑key recipient issues:\n\n- " + "\n- ".join(rk_issues))

        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="ft-card">', unsafe_allow_html=True)
        with st.expander("Advanced (IDs, GIDs, Timezone, Redirect Port)", expanded=False):
            st.markdown("##### Google Drive / Sheets IDs")
            ga1, ga2 = st.columns([1, 1])
            with ga1:
                calc_id = st.text_input("Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID)
                mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID)
            with ga2:
                incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID)
                order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID)

            st.markdown("##### GIDs & Named Ranges")
            gb1, gb2 = st.columns([1, 1])
            with gb1:
                gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF))
                loc_sheet = st.text_input("Location Sheet Title", value=cfg.LOCATION_SHEET_TITLE)
            with gb2:
                gid_order = st.text_input("Order CSV gid", value=str(cfg.GID_ORDER_CSV))
                loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE)

            st.markdown("##### Time & OAuth")
            gc1, gc2 = st.columns([1, 1])
            with gc1:
                tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ)
                tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT)
            with gc2:
                raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                redirect_port = st.number_input(
                    "Redirect Port (0 = auto)",
                    min_value=0, max_value=65535,
                    value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                    help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535."
                )
        st.markdown('</div>', unsafe_allow_html=True)

        save_drive_defaults = st.checkbox("Update defaults", value=False)

    # ----------------------------
    # Submission handling
    # ----------------------------
    def _split_emails(csv_str: str):
        return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

    if submitted:
        # Apply per-run config
        cfg.TO_RECIPIENTS = _split_emails(to)
        cfg.CC_RECIPIENTS = _split_emails(cc)
        cfg.USE_ALL_REPORT_KEYS = use_all
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

        cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
        cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full
        cfg.EMAIL_MANAGER_REPORT = bool(email_mgr)

        cfg.CALC_SPREADSHEET_ID = calc_id
        cfg.INCOMING_FOLDER_ID = incoming_id
        cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
        cfg.ORDER_REPORT_FOLDER_ID = order_folder
        cfg.REDIRECT_PORT = int(redirect_port)

        cfg.GID_MANAGER_PDF = gid_mgr
        cfg.GID_ORDER_CSV = gid_order
        cfg.LOCATION_SHEET_TITLE = loc_sheet
        cfg.LOCATION_NAMED_RANGE = loc_range
        cfg.TIMESTAMP_TZ = tz
        cfg.TIMESTAMP_FMT = tfmt

        # Per-key recipients from editor
        
        rk_map = {}
        for r in edited_rows:
            key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
            emails_csv = r.get("Emails (comma)") or ""
            emails = _parse_emails(emails_csv)
            if key and emails:
                rk_map[key] = emails
        cfg.REPORT_KEY_RECIPIENTS = rk_map

        # --- ADD: warnings before kicking off the run ---
        # 1) Warn if use_all is OFF and no explicit keys provided
        requested_keys = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]
        if not use_all and not requested_keys:
            st.warning(
                "You left **Use all Report Keys** OFF but provided **no keys** to run. "
                "No per‑key outputs will be generated unless you add keys.",
                icon="⚠️"
            )

        # 2) Warn if no recipients anywhere (TO, DEFAULT_ORDER, or per‑key map)
        any_to = bool(_parse_emails(to) or (cfg.TO_RECIPIENTS or []))
        any_default = bool(cfg.DEFAULT_ORDER_RECIPIENTS or [])
        any_per_key = bool(cfg.REPORT_KEY_RECIPIENTS)
        if not (any_to or any_default or any_per_key):
            st.warning(
                "No recipients are defined: **TO**, **DEFAULT_ORDER_RECIPIENTS**, and **Per‑Report‑Key** are all empty. "
                "Emails will not be sent.",
                icon="⚠️"
            )

        # 3) Surface per-key table issues detected earlier
        if rk_issues:
            st.warning(
                "Per‑report‑key recipient issues detected above. These may prevent emails from sending correctly:\n\n- "
                + "\n- ".join(rk_issues),
                icon="⚠️"
            )
        # --- END ADD ---


        # --- Save edited defaults to Drive JSON (optional) ---
        if save_drive_defaults:
            try:
                # Ensure we have a user token first
                creds = load_valid_token(cfg.SCOPES)
                if not creds:
                    st.error("Not authenticated. Please complete Google sign‑in first (top of page).")
                else:
                    # Drive service
                    _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                    # What we persist (the fields you asked to move out of Secrets)
                    drive_defaults = {
                        "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
                        "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
                        "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
                        "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,

                        "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
                        "GID_ORDER_CSV": cfg.GID_ORDER_CSV,

                        "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
                        "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,

                        "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
                        "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,

                        "TO_RECIPIENTS": cfg.TO_RECIPIENTS,   # lists are fine; JSON keeps types
                        "CC_RECIPIENTS": cfg.CC_RECIPIENTS,

                        "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
                        "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,

                        "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,

                        "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,

                        "INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL": cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL,
                        "SEND_SEPARATE_FULL_ORDER_EMAIL": cfg.SEND_SEPARATE_FULL_ORDER_EMAIL,
                        "EMAIL_MANAGER_REPORT": cfg.EMAIL_MANAGER_REPORT
                    }

                    # If you have CONFIG_FILE_ID in Secrets, we update that exact file.
                    # Otherwise we'll upsert a file named 'favtrip_config.json' and return its id.
                    CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
                    new_id = save_config_to_drive(
                        drive,
                        drive_defaults,
                        file_id=CONFIG_FILE_ID or None,   # update/pin if set, else upsert by name
                        # parent_folder_id=None,          # optional: set a folder id to create under
                    )

                    st.success(f"Saved defaults to Drive config (file id: {new_id}).")
                    if not CONFIG_FILE_ID:
                        st.info(
                            "Tip: add this ID to Streamlit Secrets as `CONFIG_FILE_ID` to pin the same file for all runs:\n"
                            f"`{new_id}`"
                        )
            except Exception as e:
                st.error(f"Failed to save defaults to Drive: {e}")

        # If user checked "Force Google re-auth for this run", kick them into auth gating first.
        if cfg.FORCE_REAUTH:
            clear_token()
            try:
                flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.session_state.oauth_flow = flow
                st.session_state.oauth_url = url
                st.session_state.auth_required = True
                st.info("Re-auth required for this run. Open the URL shown in the Authentication panel.")
                _rerun()
            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")
        else:
            # --- Live run with timer + last log (no full log after completion) ---
            # Write all logs to last_run.log; overwrite on each run
            logger = StatusLogger(print_to_console=True, file_path="last_run.log", overwrite=True)
            result_holder = {"value": None, "error": None}

            def _runner():
                try:
                    result_holder["value"] = run_pipeline(cfg, logger=logger)
                # Catch BaseException so SystemExit / KeyboardInterrupt are captured too
                except BaseException as e:
                    result_holder["error"] = e

            t0 = time.perf_counter()
            th = threading.Thread(target=_runner, daemon=True)
            th.start()

            with st.status("Running pipeline…", expanded=True) as status:
                timer_ph = st.empty()
                lastlog_ph = st.empty()
                while th.is_alive():
                    elapsed = int(time.perf_counter() - t0)
                    timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                    lastlog_ph.markdown(f"**Last:** {logger.last_line()}")
                    time.sleep(0.5)

                th.join()
                elapsed = int(time.perf_counter() - t0)
                timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                lastlog_ph.markdown(f"**Last:** {logger.last_line()}")

                if result_holder["error"]:
                    st.error(f"Run failed: {result_holder['error']}")
                    # Optional during debugging: show stack trace (remove later for a cleaner UI)
                    try:
                        st.exception(result_holder["error"])
                    except Exception:
                        pass
                    status.update(label="❌ Failed", state="error")
                else:
                    result = result_holder["value"]
                    if result is None:
                        st.error("Run finished without returning a result. Check logs and inputs (IDs, Drive access).")
                        status.update(label="⚠️ No result", state="error")
                    else:
                        st.write("### Outputs")
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Location", result.location)
                        col2.metric("Timestamp", result.timestamp)
                        mm = result.elapsed_seconds
                        col3.metric("Elapsed", f"{mm//3600:02d}:{(mm%3600)//60:02d}:{mm%60:02d}")

                        if getattr(result, "manager_pdf_link", None):
                            st.success(f"Manager PDF: {result.manager_pdf_link}")
                        if getattr(result, "full_order_link", None):
                            st.success(f"Full Order Sheet: {result.full_order_link}")

                            
                        if st.session_state.offer_log_download and os.path.exists("last_run.log"):
                            with open("last_run.log", "rb") as f:
                                st.download_button(
                                    "⬇️ Download full log (last_run.log)",
                                    f,
                                    file_name=f"last_run_{result.timestamp}.log",
                                    mime="text/plain",
                                    use_container_width=True
                                )
                            

                        status.update(label="✅ Completed", state="complete")
```

---
### file: web_url_credentials.json

```json
{"web":{"client_id":"674901450584-i6ncesf794ecg6u7olebqkag8gavbuc3.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["https://favtripreporting-dev1.streamlit.app/"]}}
```

```

---
### file: documentation/README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: documentation/generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd "C:\Users\rjrul\OneDrive - University of Iowa\000 Current Semester\004 BAIS 4150 BAIS Capstone\favtrip_reporting"
#python documentation/generate_code_bundle.py
```

---
### file: documentation/git_workflow.txt

```text

---
1. Create a dev branch

# Switch to dev branch
    git checkout dev

# Push dev branch to GitHub (the -u flag sets upstream tracking)
    git push -u origin dev

---
2. Do all development on dev
Make changes normally, then stage, commit, and push:

    git add .
    git commit -m "your message here"
    git push

All pushes go to dev (not main).

---
3. Push dev → main when ready (production release)
Use a Pull Request on GitHub:
1. Go to the repo on GitHub.
2. You will see a banner offering “Compare & Pull Request” (dev → main).
3. Open the Pull Request.
4. Review and click “Merge Pull Request”.

After merging, update your local main branch:

    git checkout main
    git pull

---
4. Keep dev updated with the latest main
After merging into main, update dev:

    git checkout dev
    git pull origin main

Or:

    git merge main

This prevents dev from drifting behind main.

---
5. Summary
- main = production, always stable
- dev = development, experimental work
- Merge dev → main only after testing

```

---
### file: documentation/requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit
openpyxl
pandas
```

---
### file: favtrip/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: favtrip/config.py

```python
from __future__ import annotations

import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv
from typing import Any, Dict

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: str = ""):
    """
    Prefer Streamlit secrets (typed TOML values), else fall back to environment.
    Returns the raw value from st.secrets (could be bool/list/dict/str) or a str from env.
    """
    try:
        import streamlit as st  # imported lazily so local CLI still works
        val = st.secrets.get(key, None)
        if val is None:
            return os.getenv(key, default)
        return val
    except Exception:
        return os.getenv(key, default)


_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _coerce_bool(v, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default


def _coerce_csv(v) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]


def _coerce_json(v):
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}


# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True
    EMAIL_MANAGER_REPORT: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300


import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv
from typing import Any, Dict

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------


def _get_secret(key: str, default: Any = None) -> Any:
    """
    Read from Streamlit secrets if present, else env var, else default.
    Does not raise if key missing; returns `default`.
    """
    # Try Streamlit secrets (if running inside Streamlit with secrets configured)
    try:
        import streamlit as st  # imported lazily to avoid hard dependency
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets.get(key, default)
    except Exception:
        pass

    # Fallback to environment variables (including those loaded by dotenv)
    val = os.getenv(key, default)
    return val



_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _coerce_bool(v, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default


def _coerce_csv(v) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]


def _coerce_json(v):
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}


# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True
    EMAIL_MANAGER_REPORT: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        """
        Load config from Streamlit secrets (preferred on cloud) or from .env (local dev),
        then overlay any values found in a Drive-backed JSON config (optional).
        Secrets may be typed (bool/list/dict), so we coerce safely.
        """
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)

        # ---------- 1) Base config from Secrets / .env ----------
        cfg = Config(
            CALC_SPREADSHEET_ID=str(_get_secret("CALC_SPREADSHEET_ID", "")),
            INCOMING_FOLDER_ID=str(_get_secret("INCOMING_FOLDER_ID", "")),
            MANAGER_REPORT_FOLDER_ID=str(_get_secret("MANAGER_REPORT_FOLDER_ID", "")),
            ORDER_REPORT_FOLDER_ID=str(_get_secret("ORDER_REPORT_FOLDER_ID", "")),

            GID_MANAGER_PDF=str(_get_secret("GID_MANAGER_PDF", "1921812573")),
            GID_ORDER_CSV=str(_get_secret("GID_ORDER_CSV", "1875928148")),
            LOCATION_SHEET_TITLE=str(_get_secret("LOCATION_SHEET_TITLE", "REFR: Values")),
            LOCATION_NAMED_RANGE=str(_get_secret("LOCATION_NAMED_RANGE", "_locations")),
            TIMESTAMP_TZ=str(_get_secret("TIMESTAMP_TZ", "America/Chicago")),
            TIMESTAMP_FMT=str(_get_secret("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p")),

            TO_RECIPIENTS=_coerce_csv(_get_secret("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_coerce_csv(_get_secret("CC_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_coerce_bool(_get_secret("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=[s.upper() for s in _coerce_csv(_get_secret("REPORT_KEY_RUN_LIST", ""))],
            REPORT_KEY_RECIPIENTS=_coerce_json(_get_secret("REPORT_KEY_RECIPIENTS", "{}")),
            DEFAULT_ORDER_RECIPIENTS=_coerce_csv(_get_secret("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_coerce_bool(
                _get_secret("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_coerce_bool(
                _get_secret("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),

            EMAIL_MANAGER_REPORT=_coerce_bool(
                _get_secret("EMAIL_MANAGER_REPORT", "true")
            ),

            SCOPES=_coerce_csv(
                _get_secret(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,"
                    "https://www.googleapis.com/auth/spreadsheets,"
                    "https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_coerce_bool(_get_secret("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(str(_get_secret("REDIRECT_PORT", "58285")) or "58285"),
            HTTP_TIMEOUT_SECONDS=int(str(_get_secret("HTTP_TIMEOUT_SECONDS", "300")) or "300"),
        )

        # ---------- 2) Overlay from Drive JSON config (optional) ----------
        # If present, read a JSON file on Drive and apply known keys over cfg.
        try:
            import streamlit as st
            from favtrip.google_client import load_valid_token, services
            from favtrip.config_store import load_config_from_drive

            CONFIG_FILE_ID = ""
            if hasattr(st, "secrets"):
                CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

            creds = load_valid_token(cfg.SCOPES)
            if creds:
                _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
                overrides = load_config_from_drive(drive, CONFIG_FILE_ID or None)  # {} if not found
                if isinstance(overrides, dict) and overrides:
                    for k, v in overrides.items():
                        if hasattr(cfg, k):
                            setattr(cfg, k, v)
        except Exception:
            # Fail-open: if Drive/token not ready yet, just return base cfg
            pass

        return cfg

    # -----------------------------------------------------------------------------
    # .env serialization (still useful if you let users "Update defaults in .env")
    # -----------------------------------------------------------------------------

    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")

    # -----------------------------------------------------------------------------
    # .env serialization (still useful if you let users "Update defaults in .env")
    # -----------------------------------------------------------------------------

    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")
```

---
### file: favtrip/config_store.py

```python
# favtrip/config_store.py
from __future__ import annotations
import io
import json
from typing import Any, Dict, Optional
from googleapiclient.discovery import Resource
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

DEFAULT_CONFIG_FILENAME = "favtrip_config.json"
DEFAULT_MIMETYPE = "application/json"

def load_config_from_drive(drive: Resource, file_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Read the JSON config stored in Google Drive.
    If file_id is None, try to discover the newest file named DEFAULT_CONFIG_FILENAME.
    Returns {} if the file doesn't exist or is empty/invalid JSON.
    """
    # Discover by name if a specific id wasn't provided
    if not file_id:
        resp = drive.files().list(
            q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
            orderBy="modifiedTime desc",
            pageSize=1,
            fields="files(id,name,modifiedTime)"
        ).execute() or {}
        files = resp.get("files", [])
        if not files:
            return {}
        file_id = files[0]["id"]

    # Stream download the file
    request = drive.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    raw = buf.getvalue().decode("utf-8", errors="replace").strip()
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except Exception:
        return {}

def save_config_to_drive(
    drive: Resource,
    data: Dict[str, Any],
    file_id: Optional[str] = None,
    parent_folder_id: Optional[str] = None
) -> str:
    """
    Write JSON config to Google Drive.
    - If file_id provided, update that file.
    - Else upsert (update if found by name, otherwise create) DEFAULT_CONFIG_FILENAME.
    Returns the Drive file ID.
    """
    payload = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(payload), mimetype=DEFAULT_MIMETYPE, resumable=True)

    if file_id:
        updated = drive.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        return updated["id"]

    # Try to find an existing file by name to update
    resp = drive.files().list(
        q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
        orderBy="modifiedTime desc",
        pageSize=1,
        fields="files(id,name)"
    ).execute() or {}
    files = resp.get("files", [])
    if files:
        fid = files[0]["id"]
        updated = drive.files().update(fileId=fid, media_body=media).execute()
        return updated["id"]

    # Create a new file
    meta = {"name": DEFAULT_CONFIG_FILENAME}
    if parent_folder_id:
        meta["parents"] = [parent_folder_id]

    created = drive.files().create(
        body=meta,
        media_body=media,
        fields="id,name"
    ).execute()
    return created["id"]
```

---
### file: favtrip/drive_utils.py

```python
from __future__ import annotations
import io
from googleapiclient.http import MediaIoBaseUpload


def find_latest_sheet(drive_svc, folder_id: str):
    q = (
        f"'{folder_id}' in parents and "
        "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    )
    resp = drive_svc.files().list(
        q=q, orderBy="createdTime desc", pageSize=1,
        fields="files(id,name,createdTime)"
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def upload_to_drive(drive_svc, data: bytes, name: str, mime: str, folder_id: str, to_sheet: bool=False):
    meta = {"name": name, "parents": [folder_id]}
    if to_sheet:
        meta["mimeType"] = "application/vnd.google-apps.spreadsheet"
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
    return drive_svc.files().create(
        body=meta, media_body=media, fields="id,name,mimeType,webViewLink"
    ).execute()

```

---
### file: favtrip/gmail_utils.py

```python
from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {ts} – {location}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Automated")
    msg.add_alternative(
        f"<p>Hi team,</p><p>Manager Report ({location})</p>"
        f"<a href='{pdf_link}'>Backup Link</a>", subtype="html"
    )
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)

```

---
### file: favtrip/google_client.py

```python
from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail
```

---
### file: favtrip/logger.py

```python
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass
```

---
### file: favtrip/pipeline.py

```python
from __future__ import annotations
import pandas
import csv
import io
import re
import requests
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

from io import BytesIO
from openpyxl import load_workbook, Workbook


from .config import Config
from .google_client import get_credentials, services
from .sheets_utils import (
    delete_sheet, copy_sheet_as, copy_first_sheet_as, refresh_sheets_with_prefix,
    get_value, first_gid
)
from .drive_utils import find_latest_sheet, upload_to_drive
from .gmail_utils import send_email, email_manager_report

CSV_MIME = "text/csv"


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


import requests
from io import BytesIO
from openpyxl import Workbook


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}&gid={gid}"
    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content


def timestamp_now(tz: str, fmt: str) -> str:
    return datetime.now(ZoneInfo(tz)).strftime(fmt)

import re

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _clean_emails(items):
    """
    Accepts a list or a comma-separated string and returns a list of valid emails.
    Trailing commas and blanks are removed. Invalid tokens are dropped silently.
    """
    if items is None:
        return []
    if isinstance(items, str):
        items = [p.strip() for p in items.split(",")]
    return [e for e in (p.strip() for p in items) if e and _EMAIL_RE.match(e)]

def _fallback_recipients(hint, *candidates):
    """
    Return the first non-empty, valid recipient list from the provided candidates.
    If all candidates are empty/invalid, raise a friendly error.
    """
    for c in candidates:
        cleaned = _clean_emails(c)
        if cleaned:
            return cleaned
    # Nothing usable found:
    raise ValueError(
        f"No valid recipients available for: {hint}. "
        f"Please provide at least one email in the UI or .env "
        f"(TO_RECIPIENTS, DEFAULT_ORDER_RECIPIENTS, or per-report-key)."
    )


@dataclass
class RunResult:
    ok: bool
    elapsed_seconds: int
    location: str
    timestamp: str
    manager_pdf_link: str | None
    full_order_link: str | None


def run_pipeline(cfg: Config, logger=None) -> RunResult:
    import time
    start = time.perf_counter()

    if logger:
        logger.info("Authorizing with Google APIs…")
    creds = get_credentials(cfg.SCOPES, cfg.REDIRECT_PORT, cfg.FORCE_REAUTH)
    sheets_svc, drive_svc, gmail_svc = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    if logger:
        logger.info("Google services ready")

    # Step 1: latest incoming
    if logger:
        logger.info("Finding latest incoming spreadsheet…")
    latest = find_latest_sheet(drive_svc, cfg.INCOMING_FOLDER_ID)
    if not latest:
        raise RuntimeError("No incoming report found.")
    new_report_id = latest["id"]
    if logger:
        logger.info(f"Latest incoming: {latest['name']} ({new_report_id})")

    # Step 2: prep calculations workbook
    if logger:
        logger.info("Preparing calculations workbook…")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
    try:
        copy_sheet_as(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", "Last Week")
        if logger:
            logger.info("Copied old 'Current Week' to 'Last Week'")
    except Exception:
        if logger:
            logger.warn("No 'Current Week' sheet exists to copy")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
    copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")
    if logger:
        logger.info("Inserted new 'Current Week' from latest incoming report")

    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ')…")
    refresh_sheets_with_prefix(sheets_svc, cfg.CALC_SPREADSHEET_ID, prefix="REFR: ", logger=logger)

    # Step 3: read location code
    location = get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.LOCATION_NAMED_RANGE)
    ts = timestamp_now(cfg.TIMESTAMP_TZ, cfg.TIMESTAMP_FMT)
    if logger:
        logger.info(f"Location: {location}; Timestamp: {ts}")

    # Step 4A: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_MANAGER_PDF, "pdf")
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 4B: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_ORDER_CSV, "csv")

    # Step 4C: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_{ts}_{location}_FULL.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_{ts}_{location}_FULL.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4D: Create per-report-key outputs (CSV) and email

    # --- Parse the master CSV into rows of dicts ---

    # master_csv_bytes = export_sheet(..., "csv")
    text = master_csv_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))

    rows_list = list(reader)
    if not rows_list:
        raise RuntimeError("CSV has no rows.")

    headers = [h.strip() for h in rows_list[0]]
    if not headers:
        raise RuntimeError("CSV has no header.")

    # Find 'report_key' column, case-insensitive
    lower_idx = {h.lower(): i for i, h in enumerate(headers)}
    if "report_key" not in lower_idx:
        raise RuntimeError("Report_Key column missing.")
    report_idx = lower_idx["report_key"]

    # Materialize rows as list[dict]
    rows = []
    for row in rows_list[1:]:
        rows.append({headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))})

    # Group by report key
    groups: dict[str, list[dict]] = {}
    for r in rows:
        key = (str(r.get(headers[report_idx]) or "").strip()) or "UNASSIGNED"
        groups.setdefault(key.upper(), []).append(r)


    
    for key, key_rows in groups.items():
        if not cfg.USE_ALL_REPORT_KEYS and key.upper() not in (cfg.REPORT_KEY_RUN_LIST or []):
            continue

        # Build CSV text in memory
        sio = io.StringIO()
        w = csv.writer(sio, lineterminator="\n")
        w.writerow(headers)
        for rr in key_rows:
            w.writerow([rr.get(h, "") for h in headers])

        key_csv_bytes = sio.getvalue().encode("utf-8")
        tag = clean_tag(key)  # already upper

        csv_name = f"Order_Report_{ts}_{location}_{tag}.csv"

        # Upload CSV to Drive; conversion to Google Sheet happens via to_sheet=True
        created = upload_to_drive(
            drive_svc, key_csv_bytes, csv_name,
            CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True
        )
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)


        # Export the Google Sheet as PDF (unchanged behavior)
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{ts}_{location}_{tag}.pdf"

        # Prefer per‑key recipients; else default; else TO
        candidates = None
        if cfg.REPORT_KEY_RECIPIENTS:
            # The keys in UI are upper-cased; normalize here too
            candidates = cfg.REPORT_KEY_RECIPIENTS.get(tag)

        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.DEFAULT_ORDER_RECIPIENTS,
            cfg.TO_RECIPIENTS,
        )

        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – {tag}"
        msg["From"] = "me"
        msg["To"] = ", ".join(recipients)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)

        # Keep the Google Sheet link; PDF remains the attached artifact
        msg.set_content(
            f"Hi {key} team,\nYour order report is ready. \n"
            f"Google Sheet: {created.get('webViewLink')}\n"
            f"Attached: {pdfname}\n—Automated"
        )
        msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=pdfname)
        if cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL:
            msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)

        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info(f"Emailed {tag}")

    # Step 4E: Send Manager Report (guarded by cfg.EMAIL_MANAGER_REPORT)
    if getattr(cfg, "EMAIL_MANAGER_REPORT", True):
        to_list = _fallback_recipients("Manager Report (TO_RECIPIENTS)", cfg.TO_RECIPIENTS)
        cc_list = _clean_emails(cfg.CC_RECIPIENTS)
        email_manager_report(
            gmail_svc, "me", to_list, cc_list,
            pdf_name, pdf_bytes, manager_link, ts, location
        )
        if logger:
            logger.info("Manager email sent")
    else:
        if logger:
            logger.info("Manager email skipped by configuration (EMAIL_MANAGER_REPORT = False)")

    

    # Step 4F: Send Full Order if needed
    full_link = full_created.get('webViewLink')
    if cfg.SEND_SEPARATE_FULL_ORDER_EMAIL:
        to_full = _fallback_recipients("FULL order", cfg.DEFAULT_ORDER_RECIPIENTS, cfg.TO_RECIPIENTS)
        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – FULL"
        msg["From"] = "me"
        msg["To"] = ", ".join(to_full)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi team,\nFULL order report is ready.\nSheet: {full_link}\nAttached: {full_pdf_name}\n—Automated")
        msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info("FULL order email sent")
    else:
        if logger:
            logger.info("Separate full order email disabled")

    elapsed = int(time.perf_counter() - start)
    if logger:
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        logger.info(f"Run completed in {h:02d}:{m:02d}:{s:02d}")

    return RunResult(True, elapsed, location, ts, manager_link, full_link)

```

---
### file: favtrip/sheets_utils.py

```python
from __future__ import annotations
import random
import time
from typing import Any, Dict, List


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def delete_sheet(svc, spreadsheet_id: str, title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), title)
    if s:
        body = {"requests": [{"deleteSheet": {"sheetId": s["sheetId"]}}]}
        svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def copy_sheet_as(svc, spreadsheet_id: str, src_title: str, new_title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), src_title)
    if not s:
        return None
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=spreadsheet_id,
        sheetId=s["sheetId"],
        body={"destinationSpreadsheetId": spreadsheet_id}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return new_id


def copy_first_sheet_as(svc, src_spreadsheet: str, dest_spreadsheet: str, new_title: str):
    meta = svc.spreadsheets().get(spreadsheetId=src_spreadsheet).execute()
    first_id = meta["sheets"][0]["properties"]["sheetId"]
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=src_spreadsheet,
        sheetId=first_id,
        body={"destinationSpreadsheetId": dest_spreadsheet}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=dest_spreadsheet, body=body).execute()
    return new_id


def refresh_sheets_with_prefix(svc, spreadsheet_id: str, prefix: str = "REFR: ", retries: int = 5, logger=None):
    sheets = list_sheets(svc, spreadsheet_id)
    targets = [s["properties"] for s in sheets if s["properties"]["title"].startswith(prefix)]
    for idx, t in enumerate(targets, start=1):
        body = {"requests": [{
            "findReplace": {
                "find": "=",
                "replacement": "=",
                "includeFormulas": True,
                "sheetId": t["sheetId"]
            }
        }]}
        attempt = 0
        while True:
            try:
                svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                if logger:
                    logger.info(f"[{idx}/{len(targets)}] Recalc OK: {t['title']}")
                break
            except Exception:
                attempt += 1
                if attempt > retries:
                    if logger:
                        logger.warn(f"FAILED recalc for {t['title']}")
                    break
                time.sleep(1 + random.random())


def get_value(svc, spreadsheet_id: str, sheet_title: str, named_range: str) -> str:
    try:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=named_range
        ).execute().get("values", [])
    except Exception:
        vals = []
    if not vals:
        try:
            vals = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_title}'!A1:A"
            ).execute().get("values", [])
        except Exception:
            vals = []
    return vals[0][0] if vals and vals[0] else "UNKNOWN"


def first_gid(svc, spreadsheet_id: str) -> int:
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return meta["sheets"][0]["properties"]["sheetId"]

```

---
### file: last_run.log

```
[2026-03-03 23:16:33] INFO: Authorizing with Google APIs…
[2026-03-03 23:16:33] INFO: Google services ready
[2026-03-03 23:16:33] INFO: Finding latest incoming spreadsheet…
[2026-03-03 23:16:34] INFO: Latest incoming: Testing Sales Report - Week 5 (1kNjmEbljdUIqJUjwh8e2-plfLWRZSGHxMHxL50-2ce4)
[2026-03-03 23:16:34] INFO: Preparing calculations workbook…
[2026-03-03 23:16:42] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 23:16:50] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 23:16:50] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 23:17:17] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 23:17:22] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 23:18:15] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 23:18:17] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-11-18-PM
[2026-03-03 23:18:17] INFO: Exporting Manager Report (PDF)…
[2026-03-03 23:18:21] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/1UILYk_poamIwe6JrFQUGZIz9az9iBsdU/view?usp=drivesdk
[2026-03-03 23:18:21] INFO: Exporting Master Order (CSV)…
[2026-03-03 23:18:34] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1bOcle5eMseFc43TZMnXLqgHR7_DcPs13ntb6PSBians/edit?usp=drivesdk
[2026-03-03 23:18:42] INFO: Emailed COFFEE
[2026-03-03 23:18:43] INFO: Manager email sent
[2026-03-03 23:18:43] INFO: Separate full order email disabled
[2026-03-03 23:18:43] INFO: Run completed in 00:02:09

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: requirements.txt

```text
-r documentation/requirements.txt
```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```

---
### file: ui_streamlit.py

```python
import os
import time
import threading
import json
import base64
import hashlib
import secrets
import re

import streamlit as st
from streamlit.components.v1 import html
from streamlit.components.v1 import html as _html_listener

from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials

from favtrip.google_client import load_valid_token, services, clear_token
from favtrip.config_store import save_config_to_drive
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.drive_utils import upload_to_drive


# =========================
# Constants & Simple Helpers
# =========================

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _split_emails(csv_str: str):
    return [e.strip() for e in (csv_str or "").split(",") if e.strip()]


def _parse_emails(csv_str: str):
    return _split_emails(csv_str)


def _invalid_emails(csv_str: str):
    return [e for e in _parse_emails(csv_str) if not EMAIL_RE.match(e)]


def _analyze_rk_rows(rows):
    """
    Validate the 'Per-Report-Key Recipients' editor rows.
    Returns (issues: List[str], preview_lines: List[str], rk_map: Dict[str, List[str]])
    """
    issues, preview, rk_map = [], [], {}
    seen, dupes = set(), set()

    for idx, r in enumerate(rows or [], start=1):
        raw_key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
        emails_csv = r.get("Emails (comma)") or ""
        if not raw_key and not emails_csv:
            # allow a blank template row
            continue

        # uppercase flag
        if raw_key != raw_key.upper():
            issues.append(f"Row {idx}: key '{raw_key}' is not ALL CAPS.")

        # duplicate detection
        if raw_key:
            if raw_key in seen:
                dupes.add(raw_key)
            else:
                seen.add(raw_key)

        # email validation
        bads = _invalid_emails(emails_csv)
        if bads:
            issues.append(f"Row {idx}: invalid emails → {', '.join(bads)}")

        # mapping + preview
        if raw_key:
            emails = _parse_emails(emails_csv)
            if emails:
                rk_map[raw_key] = emails
            preview.append(f"{raw_key} → {', '.join(emails) if emails else emails_csv}")

    if dupes:
        issues.append(f"Duplicate keys detected: {', '.join(sorted(dupes))}")
    return issues, preview, rk_map


def _b64url(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")


def _pkce_pair() -> tuple[str, str]:
    """
    Generate a high-entropy PKCE code_verifier and its S256 code_challenge.
    RFC 7636 requires 43–128 chars; this approach yields a URL-safe value.
    """
    verifier = _b64url(secrets.token_bytes(64))        # ~86 chars, URL-safe, no padding
    challenge = _b64url(hashlib.sha256(verifier.encode("ascii")).digest())
    return verifier, challenge


def _redirect_base() -> str:
    """
    Always return a non-empty redirect base that exactly matches your OAuth client's
    Authorized redirect URI. Prefer Secrets; normalize to one trailing slash.
    """
    base = (st.secrets.get("APP_BASE_URL", "") or "").strip()
    if not base:
        # Fallback to request (often available), still normalized
        try:
            base = (st.request.url_root or "").strip()
        except Exception:
            base = ""
    if not base:
        st.error("OAuth redirect base is not set. Define APP_BASE_URL in Secrets.")
        st.stop()
    return base.rstrip("/") + "/"


def _parse_state(state_b64: str) -> dict:
    # Add padding back for base64 decoding if needed
    padding = "=" * ((4 - len(state_b64) % 4) % 4)
    raw = base64.urlsafe_b64decode(state_b64 + padding)
    return json.loads(raw.decode("utf-8"))


def _infer_media_mime(name: str) -> str:
    n = (name or "").lower()
    if n.endswith(".csv"):
        return "text/csv"
    if n.endswith(".xlsx"):
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return "application/octet-stream"


def _get_drive_service_or_raise(cfg):
    creds = load_valid_token(cfg.SCOPES)
    if not creds:
        raise RuntimeError("Google authorization required. Please sign in first.")
    _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    return drive


def _rerun():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()


# =========================
# OAuth (Web / PKCE)
# =========================

def start_web_oauth(scopes):
    """
    Build an authorization URL that:
      - uses a stable redirect_uri (from Secrets)
      - uses explicit PKCE (S256)
      - embeds the code_verifier inside the state (base64url(JSON))
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    redirect = _redirect_base()

    # Explicit PKCE (stateless across redirect)
    code_verifier, code_challenge = _pkce_pair()

    # CSRF token + verifier encoded into state that Google will return unchanged.
    state_obj = {
        "csrf": _b64url(secrets.token_bytes(16)),
        "v": code_verifier,
        "r": redirect,
    }
    state_b64 = _b64url(json.dumps(state_obj).encode("utf-8"))

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
        state=state_b64,
        code_challenge=code_challenge,
        code_challenge_method="S256",
    )

    # Keep only minimal context; state carries the verifier.
    st.session_state["_oauth_redirect"] = redirect
    return auth_url


def finish_web_oauth(code: str, state_b64: str, scopes):
    """
    Recreate a Flow with the same redirect_uri and exchange code + code_verifier for tokens.
    (No UI side effects here.)
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    state_obj = _parse_state(state_b64)
    code_verifier = state_obj.get("v")
    redirect = state_obj.get("r") or st.session_state.get("_oauth_redirect") or _redirect_base()

    if not code_verifier:
        st.error("OAuth state did not include a PKCE code_verifier.")
        st.stop()

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    flow.fetch_token(code=code, code_verifier=code_verifier)

    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds

# --- OAuth Redirect Handler (Web/PKCE only) ---


# =========================
# UI Sections
# =========================

def render_run_form(cfg):
    # --- STATE INIT ---
    if "incoming_uploader_version" not in st.session_state:
        st.session_state.incoming_uploader_version = 0
    if "incoming_selected_name" not in st.session_state:
        st.session_state.incoming_selected_name = None
    if "incoming_uploaded_ok" not in st.session_state:
        st.session_state.incoming_uploaded_ok = False

    uploader_key = f"incoming_upload_v{st.session_state.incoming_uploader_version}"

    # =========================
    # UPLOAD CARD (outside form)
    # =========================
    
    with st.container(border=True):
        st.subheader("Upload Current Week Sales Report")
        # Use the SAME column grid as the run form header: [4, 1, 1]
        up_col, _, upbtn_col = st.columns([4, 1, 1])

        with up_col:
            incoming_file = st.file_uploader(
                "Upload Current Week Sales Report",
                type=["xlsx", "csv"],
                key=uploader_key,
                help="Please upload the current week's 'Live Items Report' from Modisoft as an XLSX or CSV file.",
                label_visibility="collapsed",
                accept_multiple_files=False,
            )

        # Track selection to manage gating (must click Upload Now successfully before Run)
        file_selected = incoming_file is not None
        if file_selected and st.session_state.get("incoming_selected_name") != incoming_file.name:
            st.session_state.incoming_selected_name = incoming_file.name
            st.session_state.incoming_uploaded_ok = False

        with upbtn_col:
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)
            upload_clicked = st.button(
                "⬆️ Upload Now",
                use_container_width=True,
                disabled=(not file_selected),
                type="primary",
                key="upload_submit",
            )
            st.markdown('</div>', unsafe_allow_html=True)

        # --- Handle the upload action immediately ---
        if upload_clicked:
            if not cfg.INCOMING_FOLDER_ID:
                st.error("Incoming Folder ID is empty. Set it under **Advanced → Incoming Folder ID**.")
            elif incoming_file is None:
                st.warning("Choose a .xlsx or .csv file first.")
            else:
                try:
                    drive = _get_drive_service_or_raise(cfg)
                    media_mime = _infer_media_mime(incoming_file.name)
                    base_name = os.path.splitext(incoming_file.name)[0]
                    nice_name = f"{base_name} (uploaded via UI)"
                    created = upload_to_drive(
                        drive,
                        data=incoming_file.getvalue(),
                        name=nice_name,
                        mime=media_mime,
                        folder_id=cfg.INCOMING_FOLDER_ID,
                        to_sheet=True,
                    )
                    link = created.get("webViewLink", "")

                    st.session_state.incoming_uploaded_ok = True
                    st.session_state.incoming_uploader_version += 1
                    st.session_state.incoming_selected_name = None
           

                    st.success("✅ Uploaded to Incoming as a Google Sheet.")
                    if link:
                        st.link_button("Open uploaded Sheet", link, use_container_width=True)
                    st.caption("This will be treated as the latest incoming report on the next run.")
                    _rerun()
                except Exception as e:
                    st.error(f"Upload failed: {e}")


    # =========================
    # RUN FORM CARD
    # =========================
    # When upload is OK, make the Run button green by adding .ft-run-green to the page segment
    
    run_form_wrapper_classes = "ft-card ft-row"
    had_ok = st.session_state.get("incoming_uploaded_ok", False)
    if had_ok:
        run_form_wrapper_classes += " ft-run-green"

    
    # OPEN the wrapper with real HTML (no entities)
    st.markdown(f'<div class="{run_form_wrapper_classes}">', unsafe_allow_html=True)



    with st.form("run_form"):
        # Header row uses the same columns to align the Run button with Upload button above
        tl, _, col_run = st.columns([4, 1, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")

        # --- Unified gating logic ---
        # A) If a file is currently selected but not uploaded -> disable Run
        # B) If no file selected and we have a prior successful upload -> enable Run
        # C) Otherwise (no prior upload or ambiguous state) -> disable Run
        file_selected = st.session_state.get("incoming_selected_name") is not None
        had_ok = st.session_state.get("incoming_uploaded_ok", False)

        run_enabled = (had_ok and not file_selected)
        run_disabled = not run_enabled

        with col_run:
            # Right-align and full-width, matching Upload Now
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)
            submitted = st.form_submit_button(
                "▶️ Run Pipeline",
                use_container_width=True,
                disabled=run_disabled,
                type="primary",
                key="run_submit"
            )
            st.markdown('</div>', unsafe_allow_html=True)

        # ----- Main options -----

        # Recipients
        st.markdown("##### Recipients")
        col1, col2 = st.columns([1, 1])
        with col1:
            to = st.text_input(
                "To (comma)", value=",".join(cfg.TO_RECIPIENTS or []),
                help="Fallback recipients for Manager & Order emails."
            )
        with col2:
            cc = st.text_input(
                "CC (comma)", value=",".join(cfg.CC_RECIPIENTS or []),
                help="Optional CC added to all emails."
            )

        # Report Keys
        st.markdown("##### Report Keys")
        colk1, colk2 = st.columns([0.45, 1.55])
        with colk1:
            use_all = st.toggle(
                "Use all keys from CSV",
                value=cfg.USE_ALL_REPORT_KEYS,
                help="ON: process every key found. OFF: only the keys you list."
            )
        with colk2:
            report_keys = st.text_input(
                "Keys to run (comma)",
                value=",".join(cfg.REPORT_KEY_RUN_LIST or []),
                help="Used when 'Use all keys' is OFF. Example: COFFEE,GROCERY"
            )

        # Email Behavior
        st.markdown("##### Email Behavior")
        cole1, cole2, cole3 = st.columns([1, 1, 1])
        with cole1:
            include_full = st.toggle(
                "Attach FULL order in each email",
                value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL
            )
        with cole2:
            send_full = st.toggle(
                "Send separate FULL order email",
                value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL
            )
        with cole3:
            email_mgr = st.toggle(
                "Email Manager Report",
                value=getattr(cfg, "EMAIL_MANAGER_REPORT", True),
                help="When ON, the Manager Report email is sent. When OFF, it is skipped."
            )

        # Per‑Report‑Key Recipients
        with st.expander("Per‑Report‑Key Recipients (optional)", expanded=False):
            st.caption("Map **REPORT KEY (ALL CAPS)** → **Emails (comma)**.")
            rows = []
            if cfg.REPORT_KEY_RECIPIENTS:
                for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({"REPORT KEY (ALL CAPS)": k, "Emails (comma)": ",".join(v or [])})
            else:
                rows = [{"REPORT KEY (ALL CAPS)": "", "Emails (comma)": ""}]
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                use_container_width=True,
                key="rk_editor",
            )

            rk_issues, rk_preview, rk_map_preview = _analyze_rk_rows(edited_rows)
            if rk_preview:
                with st.expander("Row template preview"):
                    st.code("\n".join(rk_preview), language="text")
            if rk_issues:
                st.warning("Per‑report‑key recipient issues:\n\n- " + "\n- ".join(rk_issues))

        # Advanced
        with st.expander("Advanced (IDs, GIDs, Timezone, Redirect Port)", expanded=False):
            st.markdown("##### Google Drive / Sheets IDs")
            ga1, ga2 = st.columns([1, 1])
            with ga1:
                calc_id = st.text_input("Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID)
                mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID)
            with ga2:
                incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID)
                order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID)

            st.markdown("##### GIDs & Named Ranges")
            gb1, gb2 = st.columns([1, 1])
            with gb1:
                gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF))
                loc_sheet = st.text_input("Location Sheet Title", value=cfg.LOCATION_SHEET_TITLE)
            with gb2:
                gid_order = st.text_input("Order CSV gid", value=str(cfg.GID_ORDER_CSV))
                loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE)

            st.markdown("##### Time & OAuth")
            gc1, gc2 = st.columns([1, 1])
            with gc1:
                tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ)
                tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT)
            with gc2:
                raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                redirect_port = st.number_input(
                    "Redirect Port (0 = auto)",
                    min_value=0, max_value=65535,
                    value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                    help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535."
                )

        save_drive_defaults = st.checkbox("Update defaults", value=False)

        # ----- Submission handling -----
        def _split_emails(csv_str: str):
            return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

        if submitted:
            # Apply per-run config
            cfg.TO_RECIPIENTS = _split_emails(to)
            cfg.CC_RECIPIENTS = _split_emails(cc)
            cfg.USE_ALL_REPORT_KEYS = use_all
            cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

            cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
            cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full
            cfg.EMAIL_MANAGER_REPORT = bool(email_mgr)

            cfg.CALC_SPREADSHEET_ID = calc_id
            cfg.INCOMING_FOLDER_ID = incoming_id
            cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
            cfg.ORDER_REPORT_FOLDER_ID = order_folder
            cfg.REDIRECT_PORT = int(redirect_port)

            cfg.GID_MANAGER_PDF = gid_mgr
            cfg.GID_ORDER_CSV = gid_order
            cfg.LOCATION_SHEET_TITLE = loc_sheet
            cfg.LOCATION_NAMED_RANGE = loc_range
            cfg.TIMESTAMP_TZ = tz
            cfg.TIMESTAMP_FMT = tfmt

            # Per-key recipients from editor
            
            rk_map = {}
            for r in edited_rows:
                key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
                emails_csv = r.get("Emails (comma)") or ""
                emails = _parse_emails(emails_csv)
                if key and emails:
                    rk_map[key] = emails
            cfg.REPORT_KEY_RECIPIENTS = rk_map

            # --- ADD: warnings before kicking off the run ---
            # 1) Warn if use_all is OFF and no explicit keys provided
            requested_keys = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]
            if not use_all and not requested_keys:
                st.warning(
                    "You left **Use all Report Keys** OFF but provided **no keys** to run. "
                    "No per‑key outputs will be generated unless you add keys.",
                    icon="⚠️"
                )

            # 2) Warn if no recipients anywhere (TO, DEFAULT_ORDER, or per‑key map)
            any_to = bool(_parse_emails(to) or (cfg.TO_RECIPIENTS or []))
            any_default = bool(cfg.DEFAULT_ORDER_RECIPIENTS or [])
            any_per_key = bool(cfg.REPORT_KEY_RECIPIENTS)
            if not (any_to or any_default or any_per_key):
                st.warning(
                    "No recipients are defined: **TO**, **DEFAULT_ORDER_RECIPIENTS**, and **Per‑Report‑Key** are all empty. "
                    "Emails will not be sent.",
                    icon="⚠️"
                )

            # 3) Surface per-key table issues detected earlier
            if rk_issues:
                st.warning(
                    "Per‑report‑key recipient issues detected above. These may prevent emails from sending correctly:\n\n- "
                    + "\n- ".join(rk_issues),
                    icon="⚠️"
                )
            # --- END ADD ---


            # --- Save edited defaults to Drive JSON (optional) ---
            if save_drive_defaults:
                try:
                    # Ensure we have a user token first
                    creds = load_valid_token(cfg.SCOPES)
                    if not creds:
                        st.error("Not authenticated. Please complete Google sign‑in first (top of page).")
                    else:
                        # Drive service
                        _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                        # What we persist (the fields you asked to move out of Secrets)
                        drive_defaults = {
                            "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
                            "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
                            "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
                            "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,

                            "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
                            "GID_ORDER_CSV": cfg.GID_ORDER_CSV,

                            "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
                            "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,

                            "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
                            "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,

                            "TO_RECIPIENTS": cfg.TO_RECIPIENTS,   # lists are fine; JSON keeps types
                            "CC_RECIPIENTS": cfg.CC_RECIPIENTS,

                            "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
                            "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,

                            "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,

                            "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,

                            "INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL": cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL,
                            "SEND_SEPARATE_FULL_ORDER_EMAIL": cfg.SEND_SEPARATE_FULL_ORDER_EMAIL,
                            "EMAIL_MANAGER_REPORT": cfg.EMAIL_MANAGER_REPORT
                        }

                        # If you have CONFIG_FILE_ID in Secrets, we update that exact file.
                        # Otherwise we'll upsert a file named 'favtrip_config.json' and return its id.
                        CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
                        new_id = save_config_to_drive(
                            drive,
                            drive_defaults,
                            file_id=CONFIG_FILE_ID or None,   # update/pin if set, else upsert by name
                            # parent_folder_id=None,          # optional: set a folder id to create under
                        )

                        st.success(f"Saved defaults to Drive config (file id: {new_id}).")
                        if not CONFIG_FILE_ID:
                            st.info(
                                "Tip: add this ID to Streamlit Secrets as `CONFIG_FILE_ID` to pin the same file for all runs:\n"
                                f"`{new_id}`"
                            )
                except Exception as e:
                    st.error(f"Failed to save defaults to Drive: {e}")

            # If user checked "Force Google re-auth for this run", kick them into auth gating first.
            if cfg.FORCE_REAUTH:
                clear_token()
                try:
                    flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                    st.session_state.oauth_flow = flow
                    st.session_state.oauth_url = url
                    st.session_state.auth_required = True
                    st.info("Re-auth required for this run. Open the URL shown in the Authentication panel.")
                    _rerun()
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")
            else:
                # --- Live run with timer + last log (no full log after completion) ---
                # Write all logs to last_run.log; overwrite on each run
                logger = StatusLogger(print_to_console=True, file_path="last_run.log", overwrite=True)
                result_holder = {"value": None, "error": None}

                def _runner():
                    try:
                        result_holder["value"] = run_pipeline(cfg, logger=logger)
                    # Catch BaseException so SystemExit / KeyboardInterrupt are captured too
                    except BaseException as e:
                        result_holder["error"] = e

                t0 = time.perf_counter()
                th = threading.Thread(target=_runner, daemon=True)
                th.start()

                with st.status("Running pipeline…", expanded=True) as status:
                    timer_ph = st.empty()
                    lastlog_ph = st.empty()
                    while th.is_alive():
                        elapsed = int(time.perf_counter() - t0)
                        timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                        lastlog_ph.markdown(f"**Last:** {logger.last_line()}")
                        time.sleep(0.5)

                    th.join()
                    elapsed = int(time.perf_counter() - t0)
                    timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                    lastlog_ph.markdown(f"**Last:** {logger.last_line()}")

                    if result_holder["error"]:
                        st.error(f"Run failed: {result_holder['error']}")
                        # Optional during debugging: show stack trace (remove later for a cleaner UI)
                        try:
                            st.exception(result_holder["error"])
                        except Exception:
                            pass
                        status.update(label="❌ Failed", state="error")
                    else:
                        result = result_holder["value"]
                        if result is None:
                            st.error("Run finished without returning a result. Check logs and inputs (IDs, Drive access).")
                            status.update(label="⚠️ No result", state="error")
                        else:
                            st.write("### Outputs")
                            col1, col2, col3 = st.columns(3)
                            col1.metric("Location", result.location)
                            col2.metric("Timestamp", result.timestamp)
                            mm = result.elapsed_seconds
                            col3.metric("Elapsed", f"{mm//3600:02d}:{(mm%3600)//60:02d}:{mm%60:02d}")

                            if getattr(result, "manager_pdf_link", None):
                                st.success(f"Manager PDF: {result.manager_pdf_link}")
                            if getattr(result, "full_order_link", None):
                                st.success(f"Full Order Sheet: {result.full_order_link}")

                                
                            if os.path.exists("last_run.log"):
                                with open("last_run.log", "rb") as f:
                                    st.session_state["last_run_log"] = f.read()
                                    st.session_state["last_run_timestamp"] = result.timestamp
                            
                            st.session_state.incoming_uploaded_ok = False
                            st.session_state.incoming_uploader_version += 1
                            st.session_state.incoming_selected_name = None

                            status.update(label="✅ Completed", state="complete")
                            time.sleep(10)
                            _rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    if "last_run_log" in st.session_state:
        st.download_button(
            "⬇️ Download full log (last_run.log)",
            st.session_state["last_run_log"],
            file_name=f"last_run_{st.session_state['last_run_timestamp']}.log",
            mime="text/plain",
            use_container_width=True
        )




# =========================
# App Entrypoint
# =========================

#st.title("🧾 FavTrip Reporting Pipeline")


st.set_page_config(
    page_title="FT Reporting",
    page_icon="🧾",          # emoji or path/URL to an image
    layout="wide",           # "centered" or "wide"
    initial_sidebar_state="expanded",  # "auto", "expanded", "collapsed"
    menu_items={
        "Get Help": "mailto:ryan-morrow@uiowa.edu",
        "Report a bug": "https://github.com/ryan-j-morrow/favtrip_reporting/issues",
        "About": "FavTrip Reporting Pipeline",
    },
)


st.markdown("""
<style>
/* Make only buttons inside a .ft-run-green scope look green */
.ft-run-green .stButton > button {
    background-color: #22c55e !important;  /* green */
    border-color: #16a34a !important;
    color: #ffffff !important;
}
/* Optional: dim disabled state a bit */
.ft-run-green .stButton > button:disabled {
    opacity: .6 !important;
    cursor: not-allowed !important;
}
/* Utility: right-align containers where used */
.ft-right-btn { display:flex; justify-content:flex-end; }
</style>
""", unsafe_allow_html=True)


cfg = Config.load()

# --- Finish OAuth inline when redirect comes back (this is in the NEW TAB) ---
params = st.query_params
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        # Token is saved locally in this new tab's app process
        st.success("✅ Google authentication complete.")

        # Remove code/state from URL
        st.query_params.clear()

        # No messaging back to opener and NO window.close().
        # This tab becomes the main app; just rerun to flip UI.
        st.toast("Signed in. Loading the app…")
        st.rerun()
    except Exception as e:
        st.error(f"OAuth error: {e}")

# If a valid token now exists (in this tab), reload cfg (Drive overlays, etc.)
st.session_state.auth_required = (load_valid_token(cfg.SCOPES) is None)
if not st.session_state.auth_required:
    cfg = Config.load()

# Session state defaults
if "incoming_selected_name" not in st.session_state:
    st.session_state.incoming_selected_name = None
if "incoming_uploaded_ok" not in st.session_state:
    st.session_state.incoming_uploaded_ok = False
if "offer_log_download" not in st.session_state:
    st.session_state.offer_log_download = False
if "offer_log_download" not in st.session_state:
    st.session_state.offer_log_download = False


# Sidebar (always visible)
with st.sidebar:
    st.header("Utilities")

    if st.button("Google Sign Out", type="secondary", use_container_width=True):
        clear_token()
        for key in ["auth_required", "oauth_flow", "oauth_url", "auth_checked"]:
            if key in st.session_state:
                del st.session_state[key]
        try:
            st.rerun()
        except AttributeError:
            st.experimental_rerun()
    
    st.checkbox(
        "Offer full log download after completion",
        key="offer_log_download",
        help="If enabled, a 'Download last_run.log' button appears when a run finishes.")

# Auth gate
if st.session_state.auth_required:
    # ----------------------------
    # Authentication panel (shown only if auth required)
    # ----------------------------
    if st.session_state.auth_required:
        with st.expander("Google Authentication", expanded=True):
            st.caption(
                "Authentication is required before running. "
                "Click **Sign in with Google** to open the consent screen (it will open in a new tab)."
            )

            sign_in_ph = st.empty()
            clicked = sign_in_ph.button("Sign in with Google", type="primary", use_container_width=True)

            if clicked:
                try:
                    auth_url = start_web_oauth(cfg.SCOPES)
                    sign_in_ph.empty()

                    # Friendly message in this (original) tab
                    st.markdown(
                        """
                        <div style="
                            display:flex;align-items:center;justify-content:center;
                            height:55vh;text-align:center;
                            font-family: system-ui, Segoe UI, Roboto, Helvetica, Arial, sans-serif;">
                        <div>
                            <h2 style="margin-bottom:0.5rem;">You're being signed in…</h2>
                            <p style="font-size:1.05rem;opacity:.9;">
                            A new browser tab was opened for Google sign‑in.<br/>
                            <strong>After it completes, continue in that tab.</strong>
                            </p>
                        </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    # Optional: refresh this tab when user returns (not required)
                    html(
                        """
                        <script>
                        document.addEventListener("visibilitychange", function() {
                            if (!document.hidden) { location.reload(); }
                        });
                        </script>
                        """,
                        height=0,
                    )

                    # Open Google auth in a NEW tab (this will ultimately become the main app)
                    html(
                        f"""
                        <script>
                        window.open({json.dumps(auth_url)}, "_blank", "noopener");
                        </script>
                        """,
                        height=0,
                    )

                    st.stop()
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")

            with st.expander("Having trouble?", expanded=False):
                st.write(
                    "- The Google authorization page opens in a **new browser tab**.\n"
                    "- After completing consent, the **new tab** will load the app.\n"
                    "- If you renamed your Streamlit app or URL, ensure the Google OAuth "
                    "Authorized redirect URI matches exactly (including trailing slash)."
                )
else:
    # Ensure config reflects Drive-based overrides after auth
    cfg = Config.load()
    render_run_form(cfg)
```

---
### file: web_url_credentials.json

```json
{"web":{"client_id":"674901450584-i6ncesf794ecg6u7olebqkag8gavbuc3.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["https://favtripreporting-dev1.streamlit.app/"]}}
```

```

---
### file: documentation/README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: documentation/generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd "C:\Users\rjrul\OneDrive - University of Iowa\000 Current Semester\004 BAIS 4150 BAIS Capstone\favtrip_reporting"
#python documentation/generate_code_bundle.py
```

---
### file: documentation/git_workflow.txt

```text

---
1. Create a dev branch

# Switch to dev branch
    git checkout dev

# Push dev branch to GitHub (the -u flag sets upstream tracking)
    git push -u origin dev

---
2. Do all development on dev
Make changes normally, then stage, commit, and push:

    git add .
    git commit -m "your message here"
    git push

All pushes go to dev (not main).

---
3. Push dev → main when ready (production release)
Use a Pull Request on GitHub:
1. Go to the repo on GitHub.
2. You will see a banner offering “Compare & Pull Request” (dev → main).
3. Open the Pull Request.
4. Review and click “Merge Pull Request”.

After merging, update your local main branch:

    git checkout main
    git pull

---
4. Keep dev updated with the latest main
After merging into main, update dev:

    git checkout dev
    git pull origin main

Or:

    git merge main

This prevents dev from drifting behind main.

---
5. Summary
- main = production, always stable
- dev = development, experimental work
- Merge dev → main only after testing

```

---
### file: documentation/requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit
openpyxl
pandas
```

---
### file: favtrip/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: favtrip/config.py

```python
from __future__ import annotations

import os
import json
from dataclasses import dataclass, asdict
from typing import Any, Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: Any = None) -> Any:
    """
    Read from Streamlit secrets if present, else env var, else default.
    Does not raise if key missing; returns `default`.
    """
    try:
        import streamlit as st  # imported lazily to avoid hard dependency
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets.get(key, default)
    except Exception:
        pass
    return os.getenv(key, default)

_TRUE = {"1", "true", "yes", "on", "y", "t"}

def _coerce_bool(v: Any, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default

def _coerce_csv(v: Any) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]

def _coerce_json(v: Any) -> Dict[str, Any]:
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}

# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True
    EMAIL_MANAGER_REPORT: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300

    # Advanced intake settings
    USE_AUTO_ROLLOVER_IF_ONE_WEEK: bool = True
    START_DAY_OF_WEEK: str = "Sunday"    # Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Any
    END_DAY_OF_WEEK: str = "Saturday"    # Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Any

    # Cleanup
    OUTPUT_TIME_TO_LIFE: int = 30
    FAILED_INPUT_TIME_TO_LIFE: int = 1

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        """
        Load config from Streamlit secrets (preferred on cloud) or from env/.env (local dev),
        then overlay any values found in a Drive-backed JSON config (optional).
        Secrets may be typed (bool/list/dict), so we coerce safely.
        """
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)

        cfg = Config(
            CALC_SPREADSHEET_ID=str(_get_secret("CALC_SPREADSHEET_ID", "")),
            INCOMING_FOLDER_ID=str(_get_secret("INCOMING_FOLDER_ID", "")),
            MANAGER_REPORT_FOLDER_ID=str(_get_secret("MANAGER_REPORT_FOLDER_ID", "")),
            ORDER_REPORT_FOLDER_ID=str(_get_secret("ORDER_REPORT_FOLDER_ID", "")),

            GID_MANAGER_PDF=str(_get_secret("GID_MANAGER_PDF", "1921812573")),
            GID_ORDER_CSV=str(_get_secret("GID_ORDER_CSV", "1875928148")),
            LOCATION_SHEET_TITLE=str(_get_secret("LOCATION_SHEET_TITLE", "REFR: Values")),
            LOCATION_NAMED_RANGE=str(_get_secret("LOCATION_NAMED_RANGE", "_locations")),
            TIMESTAMP_TZ=str(_get_secret("TIMESTAMP_TZ", "America/Chicago")),
            TIMESTAMP_FMT=str(_get_secret("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p")),

            OUTPUT_TIME_TO_LIFE=int(_get_secret("OUTPUT_TIME_TO_LIFE", 30)),
            FAILED_INPUT_TIME_TO_LIFE=int(_get_secret("FAILED_INPUT_TIME_TO_LIFE", 1)),

            TO_RECIPIENTS=_coerce_csv(_get_secret("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_coerce_csv(_get_secret("CC_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_coerce_bool(_get_secret("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=[s.upper() for s in _coerce_csv(_get_secret("REPORT_KEY_RUN_LIST", ""))],
            REPORT_KEY_RECIPIENTS=_coerce_json(_get_secret("REPORT_KEY_RECIPIENTS", "{}")),
            DEFAULT_ORDER_RECIPIENTS=_coerce_csv(_get_secret("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_coerce_bool(
                _get_secret("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_coerce_bool(
                _get_secret("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),
            EMAIL_MANAGER_REPORT=_coerce_bool(_get_secret("EMAIL_MANAGER_REPORT", "true")),

            SCOPES=_coerce_csv(
                _get_secret(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,"
                    "https://www.googleapis.com/auth/spreadsheets,"
                    "https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_coerce_bool(_get_secret("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(str(_get_secret("REDIRECT_PORT", "58285")) or "58285"),
            HTTP_TIMEOUT_SECONDS=int(str(_get_secret("HTTP_TIMEOUT_SECONDS", "300")) or "300"),

            USE_AUTO_ROLLOVER_IF_ONE_WEEK=_coerce_bool(_get_secret("USE_AUTO_ROLLOVER_IF_ONE_WEEK", "true")),
            START_DAY_OF_WEEK=str(_get_secret("START_DAY_OF_WEEK", "Sunday")),
            END_DAY_OF_WEEK=str(_get_secret("END_DAY_OF_WEEK", "Saturday")),
        )

        # Optional overlay from Drive JSON config (if creds + file present)
        try:
            import streamlit as st
            from favtrip.google_client import load_valid_token, services
            from favtrip.config_store import load_config_from_drive

            CONFIG_FILE_ID = ""
            if hasattr(st, "secrets"):
                CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

            creds = load_valid_token(cfg.SCOPES)
            if creds:
                _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
                overrides = load_config_from_drive(drive, CONFIG_FILE_ID or None)  # {} if not found
                if isinstance(overrides, dict) and overrides:
                    for k, v in overrides.items():
                        if hasattr(cfg, k):
                            setattr(cfg, k, v)
        except Exception:
            # Fail-open: if Drive/token not ready yet, just return base cfg
            pass

        return cfg

    # -------------------------------------------------------------------------
    # .env serialization (optional helper)
    # -------------------------------------------------------------------------
    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")

```

---
### file: favtrip/config_store.py

```python
# favtrip/config_store.py
from __future__ import annotations
import io
import json
from typing import Any, Dict, Optional
from googleapiclient.discovery import Resource
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

DEFAULT_CONFIG_FILENAME = "favtrip_config.json"
DEFAULT_MIMETYPE = "application/json"

def load_config_from_drive(drive: Resource, file_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Read the JSON config stored in Google Drive.
    If file_id is None, try to discover the newest file named DEFAULT_CONFIG_FILENAME.
    Returns {} if the file doesn't exist or is empty/invalid JSON.
    """
    # Discover by name if a specific id wasn't provided
    if not file_id:
        resp = drive.files().list(
            q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
            orderBy="modifiedTime desc",
            pageSize=1,
            fields="files(id,name,modifiedTime)"
        ).execute() or {}
        files = resp.get("files", [])
        if not files:
            return {}
        file_id = files[0]["id"]

    # Stream download the file
    request = drive.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    raw = buf.getvalue().decode("utf-8", errors="replace").strip()
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except Exception:
        return {}

def save_config_to_drive(
    drive: Resource,
    data: Dict[str, Any],
    file_id: Optional[str] = None,
    parent_folder_id: Optional[str] = None
) -> str:
    """
    Write JSON config to Google Drive.
    - If file_id provided, update that file.
    - Else upsert (update if found by name, otherwise create) DEFAULT_CONFIG_FILENAME.
    Returns the Drive file ID.
    """
    payload = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(payload), mimetype=DEFAULT_MIMETYPE, resumable=True)

    if file_id:
        updated = drive.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        return updated["id"]

    # Try to find an existing file by name to update
    resp = drive.files().list(
        q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
        orderBy="modifiedTime desc",
        pageSize=1,
        fields="files(id,name)"
    ).execute() or {}
    files = resp.get("files", [])
    if files:
        fid = files[0]["id"]
        updated = drive.files().update(fileId=fid, media_body=media).execute()
        return updated["id"]

    # Create a new file
    meta = {"name": DEFAULT_CONFIG_FILENAME}
    if parent_folder_id:
        meta["parents"] = [parent_folder_id]

    created = drive.files().create(
        body=meta,
        media_body=media,
        fields="id,name"
    ).execute()
    return created["id"]
```

---
### file: favtrip/drive_utils.py

```python
from __future__ import annotations
import io
from datetime import datetime, timedelta, timezone
from googleapiclient.http import MediaIoBaseUpload


def find_latest_sheet(drive_svc, folder_id: str):
    q = (
        f"'{folder_id}' in parents and "
        "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    )
    resp = drive_svc.files().list(
        q=q, orderBy="createdTime desc", pageSize=1,
        fields="files(id,name,createdTime)"
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def upload_to_drive(drive_svc, data: bytes, name: str, mime: str, folder_id: str, to_sheet: bool=False):
    meta = {"name": name, "parents": [folder_id]}
    if to_sheet:
        meta["mimeType"] = "application/vnd.google-apps.spreadsheet"
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
    return drive_svc.files().create(
        body=meta, media_body=media, fields="id,name,mimeType,webViewLink"
    ).execute()

def _rfc3339(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")

def trash_file(drive, file_id: str):
    return drive.files().update(fileId=file_id, body={"trashed": True}).execute()

def cleanup_folder_by_age(drive, folder_id: str, days: int, logger=None):
    if days <= 0:
        return 0

    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    cutoff_str = _rfc3339(cutoff)

    q = (
        f"'{folder_id}' in parents and trashed=false "
        f"and createdTime < '{cutoff_str}'"
    )

    trashed = 0
    page_token = None

    while True:
        resp = drive.files().list(
            q=q,
            pageSize=1000,
            orderBy="createdTime asc",
            fields="nextPageToken, files(id,name,createdTime)",
            pageToken=page_token
        ).execute() or {}

        for f in resp.get("files", []):
            try:
                trash_file(drive, f["id"])
                trashed += 1
                if logger:
                    logger.info(f"Trashed file: {f['name']} ({f['id']})")
            except Exception as e:
                if logger:
                    logger.warn(f"Failed to trash {f['id']}: {e}")

        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    return trashed
```

---
### file: favtrip/gmail_utils.py

```python
from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {ts} – {location}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Automated")
    msg.add_alternative(
        f"<p>Hi team,</p><p>Manager Report ({location})</p>"
        f"<a href='{pdf_link}'>Backup Link</a>", subtype="html"
    )
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)

```

---
### file: favtrip/google_client.py

```python
from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail
```

---
### file: favtrip/logger.py

```python
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass
```

---
### file: favtrip/pipeline.py

```python
from __future__ import annotations
import pandas
import csv
import io
import re
import requests
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

from io import BytesIO
from openpyxl import load_workbook, Workbook


from .config import Config
from .google_client import get_credentials, services
from .sheets_utils import (
    delete_sheet, copy_sheet_as, copy_first_sheet_as, refresh_sheets_with_prefix,
    get_value, first_gid,
    get_first_sheet_meta, get_values_2d, add_blank_sheet,
    add_or_replace_sheet, put_values_2d, _force_column_as_text
)
from .drive_utils import find_latest_sheet, upload_to_drive, _rfc3339, trash_file, cleanup_folder_by_age
from .gmail_utils import send_email, email_manager_report

CSV_MIME = "text/csv"


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


import requests
from io import BytesIO
from openpyxl import Workbook


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}&gid={gid}"
    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content


def timestamp_now(tz: str, fmt: str) -> str:
    return datetime.now(ZoneInfo(tz)).strftime(fmt)

class IncomingDataValidationError(Exception):
    """Raised when the incoming report is not 1 or 2 full weeks as configured."""
    pass

_DOW_MAP = {
    "Monday": 0, "Tuesday": 1, "Wednesday": 2, "Thursday": 3,
    "Friday": 4, "Saturday": 5, "Sunday": 6, "Any": None,
}

def _parse_sheet_date(cell: str | int | float):
    """
    Parse a Google Sheets date cell into a date (drops time if present).
    Accepts:
      - Google serial numbers (days since 1899-12-30)
      - Date strings: YYYY-MM-DD, MM/DD/YYYY, MM/DD/YY
      - DateTime strings: with 12h or 24h time, with or without seconds, AM/PM
      - ISO strings (date or datetime)

    Returns: datetime.date or None if unparseable.
    """
    from datetime import datetime, timedelta

    if cell is None or cell == "":
        return None

    # --- 1) Numeric serial (Google Sheets) ---
    try:
        if isinstance(cell, (int, float)) or (isinstance(cell, str) and cell.replace(".", "", 1).isdigit()):
            serial = float(cell)
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=serial)).date()
    except Exception:
        pass

    s = str(cell).strip()

    # Quick strip for weird whitespace
    s = " ".join(s.split())

    # If a timezone suffix or trailing text exists, try to isolate the datetime token
    # (We keep it simple: split on two spaces or take first token that contains '/')
    if " " in s and "/" in s:
        # Nothing fancy; the format tries below will accept the full string if they match
        pass

    # --- 2) Try common datetime formats (12h and 24h), with or without seconds ---
    dt_formats = [
        "%m/%d/%Y %I:%M:%S %p",  # 03/01/2026 12:03:45 AM
        "%m/%d/%Y %I:%M %p",     # 03/01/2026 12:03 AM
        "%m/%d/%Y %H:%M:%S",     # 03/01/2026 00:03:45
        "%m/%d/%Y %H:%M",        # 03/01/2026 00:03
        "%Y-%m-%d %H:%M:%S",     # 2026-03-01 00:03:45
        "%Y-%m-%d %H:%M",        # 2026-03-01 00:03
        "%Y-%m-%dT%H:%M:%S",     # 2026-03-01T00:03:45
        "%Y-%m-%dT%H:%M",        # 2026-03-01T00:03
    ]
    for fmt in dt_formats:
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue

    # --- 3) Try date-only formats ---
    date_formats = [
        "%Y-%m-%d",   # 2026-03-01
        "%m/%d/%Y",   # 03/01/2026
        "%m/%d/%y",   # 03/01/26
    ]
    for fmt in date_formats:
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue

    # --- 4) Last resort: Python ISO parser (handles 'YYYY-MM-DD' and full ISO datetimes) ---
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        pass

    # --- 5) If still not parsed, try taking only the date token before a space ---
    try:
        token = s.split(" ")[0]
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
            try:
                return datetime.strptime(token, fmt).date()
            except Exception:
                continue
    except Exception:
        pass

    return None

def _find_header_and_date_col(values2d):
    """
    Find the header row whose first cell == 'Store', and the 'Date' column index.
    Returns (header_row_ix, date_col_ix) or (None, None).
    """
    header_ix = None
    for r, row in enumerate(values2d):
        c0 = (row[0].strip() if row and isinstance(row[0], str) else row[0] if row else "")
        if str(c0).strip().lower() == "store":
            header_ix = r
            break
    if header_ix is None:
        return None, None
    headers = [str(h).strip() for h in values2d[header_ix]]
    date_col_ix = None
    for c, h in enumerate(headers):
        if h.lower() == "date":
            date_col_ix = c
            break
    return header_ix, date_col_ix

def _collect_unique_dates(values2d, header_ix, date_cix):
    dates = []
    for r in range(header_ix + 1, len(values2d)):
        row = values2d[r]
        if date_cix >= len(row):
            continue
        d = _parse_sheet_date(row[date_cix])
        if d:
            dates.append(d)
    return sorted(set(dates))

def _check_week_boundaries(unique_dates, start_dow, end_dow):
    """Validate first/last weekday (unless set to Any). Return (earliest, latest)."""
    if not unique_dates:
        raise IncomingDataValidationError("No dates found in incoming report.")
    earliest, latest = unique_dates[0], unique_dates[-1]
    s_ok = (_DOW_MAP[start_dow] is None) or (earliest.weekday() == _DOW_MAP[start_dow])
    e_ok = (_DOW_MAP[end_dow]   is None) or (latest.weekday()   == _DOW_MAP[end_dow])
    error_text = None
    if not (s_ok and e_ok):
        error_text = f"Please only upload 1 or 2 full weeks of data. The first day of week included in the report should be {start_dow} and the last day of week included in the report should be {end_dow}"
        raise IncomingDataValidationError(
            error_text
        )
    return earliest, latest, error_text

def _plan_weeks(unique_dates):
    """
    Decide if we have one or two weeks by count of unique calendar days.
    Returns ('one', set7) or ('two', (set7_oldest, set7_newest)).
    """
    if len(unique_dates) == 7:
        return "one", set(unique_dates)
    if len(unique_dates) == 14:
        return "two", (set(unique_dates[:7]), set(unique_dates[7:]))
    # Not 7 or 14
    raise IncomingDataValidationError(
        "Please only upload 1 or 2 full weeks of data. The first day of week included in the report should be XXX and the last day of week included in the report should be YYY"
    )

def _trim_header_if_needed(svc, spreadsheet_id: str, sheet_id: int, values2d, header_ix):
    """Ensure header is at row 0 by deleting rows above it."""
    if header_ix and header_ix > 0:
        delete_rows_range(svc, spreadsheet_id, sheet_id, 0, header_ix)

def _filter_rows_to_dates(svc, spreadsheet_id: str, sheet_id: int, values2d, header_ix, date_cix, keep_dates_set):
    """Delete all non-header rows whose Date is not in keep_dates_set."""
    bad_rows = []
    for r in range(header_ix + 1, len(values2d)):
        row = values2d[r]
        d = _parse_sheet_date(row[date_cix] if date_cix < len(row) else None)
        if (d is None) or (d not in keep_dates_set):
            bad_rows.append(r)
    delete_row_indices(svc, spreadsheet_id, sheet_id, bad_rows)


import re

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _clean_emails(items):
    """
    Accepts a list or a comma-separated string and returns a list of valid emails.
    Trailing commas and blanks are removed. Invalid tokens are dropped silently.
    """
    if items is None:
        return []
    if isinstance(items, str):
        items = [p.strip() for p in items.split(",")]
    return [e for e in (p.strip() for p in items) if e and _EMAIL_RE.match(e)]

def _fallback_recipients(hint, *candidates):
    """
    Return the first non-empty, valid recipient list from the provided candidates.
    If all candidates are empty/invalid, raise a friendly error.
    """
    for c in candidates:
        cleaned = _clean_emails(c)
        if cleaned:
            return cleaned
    # Nothing usable found:
    raise ValueError(
        f"No valid recipients available for: {hint}. "
        f"Please provide at least one email in the UI or .env "
        f"(TO_RECIPIENTS, DEFAULT_ORDER_RECIPIENTS, or per-report-key)."
    )


@dataclass
class RunResult:
    ok: bool
    elapsed_seconds: int
    location: str
    timestamp: str
    manager_pdf_link: str | None
    full_order_link: str | None


def run_pipeline(cfg: Config, logger=None) -> RunResult:
    import time
    start = time.perf_counter()

    if logger:
        logger.info("Authorizing with Google APIs…")
    creds = get_credentials(cfg.SCOPES, cfg.REDIRECT_PORT, cfg.FORCE_REAUTH)
    sheets_svc, drive_svc, gmail_svc = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    if logger:
        logger.info("Google services ready")

    # Step 1: latest incoming
    if logger:
        logger.info("Finding latest incoming spreadsheet…")
    latest = find_latest_sheet(drive_svc, cfg.INCOMING_FOLDER_ID)
    if not latest:
        raise SystemExit("No incoming report found.")
    new_report_id = latest["id"]
    if logger:
        logger.info(f"Latest incoming: {latest['name']} ({new_report_id})")

    # ---- NEW: Validate incoming weeks & plan actions (no workbook changes yet) ----
    if logger:
        logger.info("Validating incoming report (header, dates, week boundaries)…")
    first_title, first_sid = get_first_sheet_meta(sheets_svc, new_report_id)
    values = get_values_2d(sheets_svc, new_report_id, first_title, "A:Z")

    h_ix, d_cix = _find_header_and_date_col(values)
    if h_ix is None or d_cix is None:
        raise IncomingDataValidationError(
            "Unable to locate header ('Store' in A1) and/or 'Date' column in the incoming report."
        )

    unique_dates = _collect_unique_dates(values, h_ix, d_cix)

    if logger:
        logger.info(f"Found {len(unique_dates)} unique date(s) in incoming report")

    check_outputs = _check_week_boundaries(unique_dates, cfg.START_DAY_OF_WEEK, cfg.END_DAY_OF_WEEK)
    plan_kind, plan_payload = _plan_weeks(unique_dates)

    # Step 2: prep calculations workbook (branch by plan)
    if logger:
        logger.info("Preparing calculations workbook…")

    if plan_kind == "two":
        # Two weeks → build values in memory and write each in a single call
        if logger:
            logger.info("Detected 2 weeks; writing 'Last Week' (oldest 7) and 'Current Week' (newest 7) without row deletions")

        # Source header & body (we already loaded 'values' from the first sheet)
        header = [str(h) for h in values[h_ix]]
        body_rows = values[h_ix + 1 :]

        def _slice_rows(rows, date_cix, keep_dates: set):
            out = []
            for row in rows:
                d = _parse_sheet_date(row[date_cix] if date_cix < len(row) else None)
                if d and d in keep_dates:
                    out.append(row)
            return out

        keep_oldest7, keep_newest7 = plan_payload  # sets of dates from _plan_weeks
        last_week_rows = _slice_rows(body_rows, d_cix, keep_oldest7)
        current_week_rows = _slice_rows(body_rows, d_cix, keep_newest7)

        # Create fresh target sheets
        add_or_replace_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
        add_or_replace_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")

        # Force column 'Scan Code' to be text with a prefixed apostrophe
        last_week_rows = _force_column_as_text(header, last_week_rows, "Scan Code")
        current_week_rows = _force_column_as_text(header, current_week_rows, "Scan Code")

        # Bulk write (header + rows) → 1 write per sheet
        put_values_2d(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week", [header] + last_week_rows)
        put_values_2d(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", [header] + current_week_rows)

    elif plan_kind == "one" and cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK:
        # One week + rollover ON → current behavior
        if logger:
            logger.info("Detected 1 week; auto-rollover enabled → copying old Current→Last and inserting new Current")
        delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
        try:
            copy_sheet_as(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", "Last Week")
            if logger:
                logger.info("Copied old 'Current Week' to 'Last Week'")
        except Exception:
            if logger:
                logger.warn("No 'Current Week' sheet exists to copy")
        delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
        copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")

        # Trim header for Current Week
        meta = sheets_svc.spreadsheets().get(spreadsheetId=cfg.CALC_SPREADSHEET_ID).execute()
        cw_sid = next(s["properties"]["sheetId"] for s in meta["sheets"] if s["properties"]["title"] == "Current Week")
        _trim_header_if_needed(sheets_svc, cfg.CALC_SPREADSHEET_ID, cw_sid, values, h_ix)

    else:
        # One week + rollover OFF → Current Week only; Last Week blank
        if logger:
            logger.info("Detected 1 week; auto-rollover disabled → Current only, Last Week blank")
        delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
        delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
        add_blank_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
        copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")

        meta = sheets_svc.spreadsheets().get(spreadsheetId=cfg.CALC_SPREADSHEET_ID).execute()
        cw_sid = next(s["properties"]["sheetId"] for s in meta["sheets"] if s["properties"]["title"] == "Current Week")
        _trim_header_if_needed(sheets_svc, cfg.CALC_SPREADSHEET_ID, cw_sid, values, h_ix)

    # Refresh reference sheets (unchanged)
    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ')…")
    refresh_sheets_with_prefix(sheets_svc, cfg.CALC_SPREADSHEET_ID, prefix="REFR: ", logger=logger)

    # Step 3: read location code
    location = get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.LOCATION_NAMED_RANGE)
    ts = timestamp_now(cfg.TIMESTAMP_TZ, cfg.TIMESTAMP_FMT)
    if logger:
        logger.info(f"Location: {location}; Timestamp: {ts}")

    # Step 4A: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_MANAGER_PDF, "pdf")
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 4B: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_ORDER_CSV, "csv")

    # Step 4C: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_{ts}_{location}_FULL.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_{ts}_{location}_FULL.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4D: Create per-report-key outputs (CSV) and email

    # --- Parse the master CSV into rows of dicts ---

    # master_csv_bytes = export_sheet(..., "csv")
    text = master_csv_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))

    rows_list = list(reader)
    if not rows_list:
        raise RuntimeError("CSV has no rows.")

    headers = [h.strip() for h in rows_list[0]]
    if not headers:
        raise RuntimeError("CSV has no header.")

    # Find 'report_key' column, case-insensitive
    lower_idx = {h.lower(): i for i, h in enumerate(headers)}
    if "report_key" not in lower_idx:
        raise RuntimeError("Report_Key column missing.")
    report_idx = lower_idx["report_key"]

    # Materialize rows as list[dict]
    rows = []
    for row in rows_list[1:]:
        rows.append({headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))})

    # Group by report key
    groups: dict[str, list[dict]] = {}
    for r in rows:
        key = (str(r.get(headers[report_idx]) or "").strip()) or "UNASSIGNED"
        groups.setdefault(key.upper(), []).append(r)


    
    for key, key_rows in groups.items():
        if not cfg.USE_ALL_REPORT_KEYS and key.upper() not in (cfg.REPORT_KEY_RUN_LIST or []):
            continue

        # Build CSV text in memory
        sio = io.StringIO()
        w = csv.writer(sio, lineterminator="\n")
        w.writerow(headers)
        for rr in key_rows:
            w.writerow([rr.get(h, "") for h in headers])

        key_csv_bytes = sio.getvalue().encode("utf-8")
        tag = clean_tag(key)  # already upper

        csv_name = f"Order_Report_{ts}_{location}_{tag}.csv"

        # Upload CSV to Drive; conversion to Google Sheet happens via to_sheet=True
        created = upload_to_drive(
            drive_svc, key_csv_bytes, csv_name,
            CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True
        )
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)


        # Export the Google Sheet as PDF (unchanged behavior)
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{ts}_{location}_{tag}.pdf"

        # Prefer per‑key recipients; else default; else TO
        candidates = None
        if cfg.REPORT_KEY_RECIPIENTS:
            # The keys in UI are upper-cased; normalize here too
            candidates = cfg.REPORT_KEY_RECIPIENTS.get(tag)

        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.DEFAULT_ORDER_RECIPIENTS,
            cfg.TO_RECIPIENTS,
        )

        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – {tag}"
        msg["From"] = "me"
        msg["To"] = ", ".join(recipients)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)

        # Keep the Google Sheet link; PDF remains the attached artifact
        msg.set_content(
            f"Hi {key} team,\nYour order report is ready. \n"
            f"Google Sheet: {created.get('webViewLink')}\n"
            f"Attached: {pdfname}\n—Automated"
        )
        msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=pdfname)
        if cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL:
            msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)

        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info(f"Emailed {tag}")

    # Step 4E: Send Manager Report (guarded by cfg.EMAIL_MANAGER_REPORT)
    if getattr(cfg, "EMAIL_MANAGER_REPORT", True):
        to_list = _fallback_recipients("Manager Report (TO_RECIPIENTS)", cfg.TO_RECIPIENTS)
        cc_list = _clean_emails(cfg.CC_RECIPIENTS)
        email_manager_report(
            gmail_svc, "me", to_list, cc_list,
            pdf_name, pdf_bytes, manager_link, ts, location
        )
        if logger:
            logger.info("Manager email sent")
    else:
        if logger:
            logger.info("Manager email skipped by configuration (EMAIL_MANAGER_REPORT = False)")

    

    # Step 4F: Send Full Order if needed
    full_link = full_created.get('webViewLink')
    if cfg.SEND_SEPARATE_FULL_ORDER_EMAIL:
        to_full = _fallback_recipients("FULL order", cfg.DEFAULT_ORDER_RECIPIENTS, cfg.TO_RECIPIENTS)
        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – FULL"
        msg["From"] = "me"
        msg["To"] = ", ".join(to_full)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi team,\nFULL order report is ready.\nSheet: {full_link}\nAttached: {full_pdf_name}\n—Automated")
        msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info("FULL order email sent")
    else:
        if logger:
            logger.info("Separate full order email disabled")

    
    # Step 4G: File Cleanup

    try:
        if logger:
            logger.info("Cleaning up used incoming file…")
        trash_file(drive_svc, new_report_id)

        if logger:
            logger.info("Cleaning old incoming files…")
        cleanup_folder_by_age(
            drive_svc,
            cfg.INCOMING_FOLDER_ID,
            cfg.FAILED_INPUT_TIME_TO_LIFE,
            logger
        )

        if logger:
            logger.info("Cleaning old output files…")
        for folder in [
            cfg.MANAGER_REPORT_FOLDER_ID,
            cfg.ORDER_REPORT_FOLDER_ID
        ]:
            cleanup_folder_by_age(
                drive_svc,
                folder,
                cfg.OUTPUT_TIME_TO_LIFE,
                logger
            )

    except Exception as e:
        if logger:
            logger.warn(f"Housekeeping failed: {e}")

    elapsed = int(time.perf_counter() - start)
    if logger:
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        logger.info(f"Run completed in {h:02d}:{m:02d}:{s:02d}")

    return RunResult(True, elapsed, location, ts, manager_link, full_link)

```

---
### file: favtrip/sheets_utils.py

```python
from __future__ import annotations
import random
import time
from typing import Any, Dict, List


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def delete_sheet(svc, spreadsheet_id: str, title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), title)
    if s:
        body = {"requests": [{"deleteSheet": {"sheetId": s["sheetId"]}}]}
        svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def copy_sheet_as(svc, spreadsheet_id: str, src_title: str, new_title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), src_title)
    if not s:
        return None
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=spreadsheet_id,
        sheetId=s["sheetId"],
        body={"destinationSpreadsheetId": spreadsheet_id}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return new_id


def copy_first_sheet_as(svc, src_spreadsheet: str, dest_spreadsheet: str, new_title: str):
    meta = svc.spreadsheets().get(spreadsheetId=src_spreadsheet).execute()
    first_id = meta["sheets"][0]["properties"]["sheetId"]
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=src_spreadsheet,
        sheetId=first_id,
        body={"destinationSpreadsheetId": dest_spreadsheet}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=dest_spreadsheet, body=body).execute()
    return new_id


def refresh_sheets_with_prefix(svc, spreadsheet_id: str, prefix: str = "REFR: ", retries: int = 5, logger=None):
    sheets = list_sheets(svc, spreadsheet_id)
    targets = [s["properties"] for s in sheets if s["properties"]["title"].startswith(prefix)]
    for idx, t in enumerate(targets, start=1):
        body = {"requests": [{
            "findReplace": {
                "find": "=",
                "replacement": "=",
                "includeFormulas": True,
                "sheetId": t["sheetId"]
            }
        }]}
        attempt = 0
        while True:
            try:
                svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                if logger:
                    logger.info(f"[{idx}/{len(targets)}] Recalc OK: {t['title']}")
                break
            except Exception:
                attempt += 1
                if attempt > retries:
                    if logger:
                        logger.warn(f"FAILED recalc for {t['title']}")
                    break
                time.sleep(1 + random.random())


def get_value(svc, spreadsheet_id: str, sheet_title: str, named_range: str) -> str:
    try:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=named_range
        ).execute().get("values", [])
    except Exception:
        vals = []
    if not vals:
        try:
            vals = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_title}'!A1:A"
            ).execute().get("values", [])
        except Exception:
            vals = []
    return vals[0][0] if vals and vals[0] else "UNKNOWN"


def first_gid(svc, spreadsheet_id: str) -> int:
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return meta["sheets"][0]["properties"]["sheetId"]

# --- Additional helpers for row inspection/edits ---

def get_first_sheet_meta(svc, spreadsheet_id: str):
    """Return (first_sheet_title, first_sheet_id)."""
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    first = meta["sheets"][0]["properties"]
    return first["title"], first["sheetId"]

def get_values_2d(svc, spreadsheet_id: str, sheet_title: str, a1_range: str = "A:Z"):
    """Fetch a 2D values array from a sheet title + A1 range."""
    rng = f"'{sheet_title}'!{a1_range}"
    res = svc.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=rng).execute()
    return res.get("values", [])

def delete_rows_range(svc, spreadsheet_id: str, sheet_id: int, start_row_index: int, end_row_index: int):
    """Delete [start_row_index, end_row_index) (0‑based; end exclusive)."""
    if end_row_index <= start_row_index:
        return
    body = {"requests": [{
        "deleteDimension": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": start_row_index,
                "endIndex": end_row_index,
            }
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

def delete_row_indices(svc, spreadsheet_id: str, sheet_id: int, row_indices_desc: list[int]):
    """Delete multiple absolute row indices (0‑based) in descending order."""
    for r in sorted(row_indices_desc, reverse=True):
        delete_rows_range(svc, spreadsheet_id, sheet_id, r, r+1)

def add_blank_sheet(svc, spreadsheet_id: str, title: str, rows: int = 1000, cols: int = 26):
    """Create a blank sheet with a given title."""
    body = {"requests": [{
        "addSheet": {"properties": {"title": title, "gridProperties": {"rowCount": rows, "columnCount": cols}}}
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

def add_or_replace_sheet(svc, spreadsheet_id: str, title: str, rows: int = 2000, cols: int = 50):
    """
    Remove any existing sheet with 'title' and add a blank one.
    """
    try:
        delete_sheet(svc, spreadsheet_id, title)
    except Exception:
        # if not present, ignore
        pass
    add_blank_sheet(svc, spreadsheet_id, title, rows, cols)

def put_values_2d(svc, spreadsheet_id: str, sheet_title: str, values: list[list]):
    """
    Write a 2D array to 'A1' of 'sheet_title' in a single update.
    """
    rng = f"'{sheet_title}'!A1"
    svc.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": values}
    ).execute()

def _force_column_as_text(header: list[str], rows: list[list], header_name: str) -> list[list]:
    """
    For the column matching header_name, coerce every non-blank value to a string
    prefixed with a single apostrophe, so Google Sheets stores it as text.
    """
    idx = None
    for i, h in enumerate(header):
        if str(h).strip().lower() == header_name.strip().lower():
            idx = i
            break
    if idx is None:
        return rows  # header not found; nothing to do

    out = []
    for r in rows:
        r2 = list(r)
        if idx < len(r2) and r2[idx] not in (None, ""):
            # ensure string and prefix with apostrophe
            r2[idx] = "'" + str(r2[idx])
        out.append(r2)
    return out

```

---
### file: last_run.log

```
[2026-03-03 23:16:33] INFO: Authorizing with Google APIs…
[2026-03-03 23:16:33] INFO: Google services ready
[2026-03-03 23:16:33] INFO: Finding latest incoming spreadsheet…
[2026-03-03 23:16:34] INFO: Latest incoming: Testing Sales Report - Week 5 (1kNjmEbljdUIqJUjwh8e2-plfLWRZSGHxMHxL50-2ce4)
[2026-03-03 23:16:34] INFO: Preparing calculations workbook…
[2026-03-03 23:16:42] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 23:16:50] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 23:16:50] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 23:17:17] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 23:17:22] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 23:18:15] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 23:18:17] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-11-18-PM
[2026-03-03 23:18:17] INFO: Exporting Manager Report (PDF)…
[2026-03-03 23:18:21] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/1UILYk_poamIwe6JrFQUGZIz9az9iBsdU/view?usp=drivesdk
[2026-03-03 23:18:21] INFO: Exporting Master Order (CSV)…
[2026-03-03 23:18:34] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1bOcle5eMseFc43TZMnXLqgHR7_DcPs13ntb6PSBians/edit?usp=drivesdk
[2026-03-03 23:18:42] INFO: Emailed COFFEE
[2026-03-03 23:18:43] INFO: Manager email sent
[2026-03-03 23:18:43] INFO: Separate full order email disabled
[2026-03-03 23:18:43] INFO: Run completed in 00:02:09

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: requirements.txt

```text
-r documentation/requirements.txt
```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```

---
### file: ui_streamlit.py

```python
import os
import time
import threading
import json
import base64
import hashlib
import secrets
import re

import streamlit as st
from streamlit.components.v1 import html
from streamlit.components.v1 import html as _html_listener

from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials

from favtrip.google_client import load_valid_token, services, clear_token
from favtrip.config_store import save_config_to_drive
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.drive_utils import upload_to_drive


# =========================
# Constants & Simple Helpers
# =========================

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _split_emails(csv_str: str):
    return [e.strip() for e in (csv_str or "").split(",") if e.strip()]


def _parse_emails(csv_str: str):
    return _split_emails(csv_str)


def _invalid_emails(csv_str: str):
    return [e for e in _parse_emails(csv_str) if not EMAIL_RE.match(e)]


def _analyze_rk_rows(rows):
    """
    Validate the 'Per-Report-Key Recipients' editor rows.
    Returns (issues: List[str], preview_lines: List[str], rk_map: Dict[str, List[str]])
    """
    issues, preview, rk_map = [], [], {}
    seen, dupes = set(), set()

    for idx, r in enumerate(rows or [], start=1):
        raw_key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
        emails_csv = r.get("Emails (comma)") or ""
        if not raw_key and not emails_csv:
            # allow a blank template row
            continue

        # uppercase flag
        if raw_key != raw_key.upper():
            issues.append(f"Row {idx}: key '{raw_key}' is not ALL CAPS.")

        # duplicate detection
        if raw_key:
            if raw_key in seen:
                dupes.add(raw_key)
            else:
                seen.add(raw_key)

        # email validation
        bads = _invalid_emails(emails_csv)
        if bads:
            issues.append(f"Row {idx}: invalid emails → {', '.join(bads)}")

        # mapping + preview
        if raw_key:
            emails = _parse_emails(emails_csv)
            if emails:
                rk_map[raw_key] = emails
            preview.append(f"{raw_key} → {', '.join(emails) if emails else emails_csv}")

    if dupes:
        issues.append(f"Duplicate keys detected: {', '.join(sorted(dupes))}")
    return issues, preview, rk_map


def _b64url(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")


def _pkce_pair() -> tuple[str, str]:
    """
    Generate a high-entropy PKCE code_verifier and its S256 code_challenge.
    RFC 7636 requires 43–128 chars; this approach yields a URL-safe value.
    """
    verifier = _b64url(secrets.token_bytes(64))        # ~86 chars, URL-safe, no padding
    challenge = _b64url(hashlib.sha256(verifier.encode("ascii")).digest())
    return verifier, challenge


def _redirect_base() -> str:
    """
    Always return a non-empty redirect base that exactly matches your OAuth client's
    Authorized redirect URI. Prefer Secrets; normalize to one trailing slash.
    """
    base = (st.secrets.get("APP_BASE_URL", "") or "").strip()
    if not base:
        # Fallback to request (often available), still normalized
        try:
            base = (st.request.url_root or "").strip()
        except Exception:
            base = ""
    if not base:
        st.error("OAuth redirect base is not set. Define APP_BASE_URL in Secrets.")
        st.stop()
    return base.rstrip("/") + "/"


def _parse_state(state_b64: str) -> dict:
    # Add padding back for base64 decoding if needed
    padding = "=" * ((4 - len(state_b64) % 4) % 4)
    raw = base64.urlsafe_b64decode(state_b64 + padding)
    return json.loads(raw.decode("utf-8"))


def _infer_media_mime(name: str) -> str:
    n = (name or "").lower()
    if n.endswith(".csv"):
        return "text/csv"
    if n.endswith(".xlsx"):
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return "application/octet-stream"


def _get_drive_service_or_raise(cfg):
    creds = load_valid_token(cfg.SCOPES)
    if not creds:
        raise RuntimeError("Google authorization required. Please sign in first.")
    _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    return drive


def _rerun():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()


# =========================
# OAuth (Web / PKCE)
# =========================

def start_web_oauth(scopes):
    """
    Build an authorization URL that:
      - uses a stable redirect_uri (from Secrets)
      - uses explicit PKCE (S256)
      - embeds the code_verifier inside the state (base64url(JSON))
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    redirect = _redirect_base()

    # Explicit PKCE (stateless across redirect)
    code_verifier, code_challenge = _pkce_pair()

    # CSRF token + verifier encoded into state that Google will return unchanged.
    state_obj = {
        "csrf": _b64url(secrets.token_bytes(16)),
        "v": code_verifier,
        "r": redirect,
    }
    state_b64 = _b64url(json.dumps(state_obj).encode("utf-8"))

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
        state=state_b64,
        code_challenge=code_challenge,
        code_challenge_method="S256",
    )

    # Keep only minimal context; state carries the verifier.
    st.session_state["_oauth_redirect"] = redirect
    return auth_url


def finish_web_oauth(code: str, state_b64: str, scopes):
    """
    Recreate a Flow with the same redirect_uri and exchange code + code_verifier for tokens.
    (No UI side effects here.)
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    state_obj = _parse_state(state_b64)
    code_verifier = state_obj.get("v")
    redirect = state_obj.get("r") or st.session_state.get("_oauth_redirect") or _redirect_base()

    if not code_verifier:
        st.error("OAuth state did not include a PKCE code_verifier.")
        st.stop()

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    flow.fetch_token(code=code, code_verifier=code_verifier)

    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds

# --- OAuth Redirect Handler (Web/PKCE only) ---


# =========================
# UI Sections
# =========================

def render_run_form(cfg):
    err = None
    # --- STATE INIT ---
    if "incoming_uploader_version" not in st.session_state:
        st.session_state.incoming_uploader_version = 0
    if "incoming_selected_name" not in st.session_state:
        st.session_state.incoming_selected_name = None
    if "incoming_uploaded_ok" not in st.session_state:
        st.session_state.incoming_uploaded_ok = False

    uploader_key = f"incoming_upload_v{st.session_state.incoming_uploader_version}"

    # =========================
    # UPLOAD CARD (outside form)
    # =========================
    
    with st.container(border=True):
        st.subheader("Upload Live Items Report")
        st.caption("Please upload a live items report from Modisoft. The report must either be 1 or 2 full weeks of data.")
        # Use the SAME column grid as the run form header: [4, 1, 1]
        up_col, _, upbtn_col = st.columns([4, 1, 1])

        with up_col:
            incoming_file = st.file_uploader(
                "Upload Current Week Sales Report",
                type=["xlsx", "csv"],
                key=uploader_key,
                help="Please upload the current 'Live Items Report' from Modisoft as an XLSX or CSV file.",
                label_visibility="collapsed",
                accept_multiple_files=False,
            )

        # Track selection to manage gating (must click Upload Now successfully before Run)
        file_selected = incoming_file is not None
        if file_selected and st.session_state.get("incoming_selected_name") != incoming_file.name:
            st.session_state.incoming_selected_name = incoming_file.name
            st.session_state.incoming_uploaded_ok = False

        with upbtn_col:
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)
            
            upload_clicked = st.button(
                "⬆️ Upload Now",
                use_container_width=True,
                disabled=(not file_selected) or st.session_state.get("freeze_ui", False),
                type="primary",
                key="upload_submit",
            )

            st.markdown('</div>', unsafe_allow_html=True)

        # --- Handle the upload action immediately ---
        if upload_clicked:
            if not cfg.INCOMING_FOLDER_ID:
                st.error("Incoming Folder ID is empty. Set it under **Advanced → Incoming Folder ID**.")
            elif incoming_file is None:
                st.warning("Choose a .xlsx or .csv file first.")
            else:
                try:
                    drive = _get_drive_service_or_raise(cfg)
                    media_mime = _infer_media_mime(incoming_file.name)
                    base_name = os.path.splitext(incoming_file.name)[0]
                    nice_name = f"{base_name} (uploaded via UI)"
                    created = upload_to_drive(
                        drive,
                        data=incoming_file.getvalue(),
                        name=nice_name,
                        mime=media_mime,
                        folder_id=cfg.INCOMING_FOLDER_ID,
                        to_sheet=True,
                    )
                    link = created.get("webViewLink", "")

                    st.session_state.incoming_uploaded_ok = True
                    st.session_state.incoming_uploader_version += 1
                    st.session_state.incoming_selected_name = None
           

                    st.success("✅ Uploaded to Incoming as a Google Sheet.")
                    if link:
                        st.link_button("Open uploaded Sheet", link, use_container_width=True)
                    st.caption("This will be treated as the latest incoming report on the next run.")
                    _rerun()
                except Exception as e:
                    st.error(f"Upload failed: {e}")


    # =========================
    # RUN FORM CARD
    # =========================
    # When upload is OK, make the Run button green by adding .ft-run-green to the page segment
    
    run_form_wrapper_classes = "ft-card ft-row"
    had_ok = st.session_state.get("incoming_uploaded_ok", False)
    if had_ok:
        run_form_wrapper_classes += " ft-run-green"

    
    # OPEN the wrapper with real HTML (no entities)
    st.markdown(f'<div class="{run_form_wrapper_classes}">', unsafe_allow_html=True)



    with st.form("run_form"):
        # Header row uses the same columns to align the Run button with Upload button above
        tl, _, col_run = st.columns([4, 1, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")

        # --- Unified gating logic ---
        # A) If a file is currently selected but not uploaded -> disable Run
        # B) If no file selected and we have a prior successful upload -> enable Run
        # C) Otherwise (no prior upload or ambiguous state) -> disable Run
        file_selected = st.session_state.get("incoming_selected_name") is not None
        had_ok = st.session_state.get("incoming_uploaded_ok", False)

        run_enabled = (had_ok and not file_selected)
        run_disabled = not run_enabled

        with col_run:
            # Right-align and full-width, matching Upload Now
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)
            submitted = st.form_submit_button(
                "▶️ Run Pipeline",
                use_container_width=True,
                disabled=run_disabled or st.session_state.get("freeze_ui", False),
                type="primary",
                key="run_submit"
            )
            st.markdown('</div>', unsafe_allow_html=True)

        # ----- Main options -----

        # Recipients
        st.markdown("##### Recipients")
        col1, col2 = st.columns([1, 1])
        with col1:
            to = st.text_input(
                "To (comma)", value=",".join(cfg.TO_RECIPIENTS or []),
                help="Fallback recipients for Manager & Order emails."
            )
        with col2:
            cc = st.text_input(
                "CC (comma)", value=",".join(cfg.CC_RECIPIENTS or []),
                help="Optional CC added to all emails."
            )

        # Report Keys
        st.markdown("##### Report Keys")
        colk1, colk2 = st.columns([0.45, 1.55])
        with colk1:
            use_all = st.toggle(
                "Use all keys from CSV",
                value=cfg.USE_ALL_REPORT_KEYS,
                help="ON: process every key found. OFF: only the keys you list."
            )
        with colk2:
            report_keys = st.text_input(
                "Keys to run (comma)",
                value=",".join(cfg.REPORT_KEY_RUN_LIST or []),
                help="Used when 'Use all keys' is OFF. Example: COFFEE,GROCERY"
            )

        # Email Behavior
        st.markdown("##### Email Behavior")
        cole1, cole2, cole3 = st.columns([1, 1, 1])
        with cole1:
            include_full = st.toggle(
                "Attach FULL order in each email",
                value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL
            )
        with cole2:
            send_full = st.toggle(
                "Send separate FULL order email",
                value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL
            )
        with cole3:
            email_mgr = st.toggle(
                "Email Manager Report",
                value=getattr(cfg, "EMAIL_MANAGER_REPORT", True),
                help="When ON, the Manager Report email is sent. When OFF, it is skipped."
            )

        # Per‑Report‑Key Recipients
        with st.expander("Per‑Report‑Key Recipients (optional)", expanded=False):
            st.caption("Map **REPORT KEY (ALL CAPS)** → **Emails (comma)**.")
            rows = []
            if cfg.REPORT_KEY_RECIPIENTS:
                for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({"REPORT KEY (ALL CAPS)": k, "Emails (comma)": ",".join(v or [])})
            else:
                rows = [{"REPORT KEY (ALL CAPS)": "", "Emails (comma)": ""}]
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                use_container_width=True,
                key="rk_editor",
            )

            rk_issues, rk_preview, rk_map_preview = _analyze_rk_rows(edited_rows)
            if rk_preview:
                with st.expander("Row template preview"):
                    st.code("\n".join(rk_preview), language="text")
            if rk_issues:
                st.warning("Per‑report‑key recipient issues:\n\n- " + "\n- ".join(rk_issues))

        # Advanced
        with st.expander("Advanced (IDs, GIDs, Timezone, Redirect Port)", expanded=False):
            st.markdown("##### Google Drive / Sheets IDs")
            ga1, ga2 = st.columns([1, 1])
            with ga1:
                calc_id = st.text_input("Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID)
                mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID)
            with ga2:
                incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID)
                order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID)

            st.markdown("##### GIDs & Named Ranges")
            gb1, gb2 = st.columns([1, 1])
            with gb1:
                gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF))
                loc_sheet = st.text_input("Location Sheet Title", value=cfg.LOCATION_SHEET_TITLE)
            with gb2:
                gid_order = st.text_input("Order CSV gid", value=str(cfg.GID_ORDER_CSV))
                loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE)

            st.markdown("##### Time & OAuth")
            gc1, gc2 = st.columns([1, 1])
            with gc1:
                tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ)
                tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT)
                raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                redirect_port = st.number_input(
                    "Redirect Port (0 = auto)",
                    min_value=0, max_value=65535,
                    value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                    help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535."
                )
                output_ttl = st.number_input(
                    "Output Time-To-Life (days)",
                    min_value=0,
                    max_value=3650,
                    value=int(cfg.OUTPUT_TIME_TO_LIFE),
                    help="Delete Manager/Order report files older than this many days after a successful run."
                    )

                failed_input_ttl = st.number_input(
                    "Failed Input Time-To-Life (days)",
                    min_value=0,
                    max_value=3650,
                    value=int(cfg.FAILED_INPUT_TIME_TO_LIFE),
                    help="Delete old unused incoming files older than this many days."
                    )

            with gc2:                                
                # --- New advanced intake settings ---
                use_rollover = st.toggle(
                    'Use automatic rollover if only 1 week is uploaded',
                    value=cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK,
                    help='If this is on, when only 1 week is uploaded, the most recent previously uploaded data will become the "Last Week" data; If this is off then the "Last Week" data will be left blank'
                )
                _days = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Any"]
                start_dow = st.selectbox(
                    "Start day of week", _days, index=_days.index(cfg.START_DAY_OF_WEEK),
                    help="The day of week that the uploaded data should start at, any other day will raise an error"
                )
                end_dow = st.selectbox(
                    "End day of week", _days, index=_days.index(cfg.END_DAY_OF_WEEK),
                    help="The day of week that the uploaded data should end at, any other day will raise an error"
                )


        save_drive_defaults = st.checkbox("Update defaults", value=False)

        # ----- Submission handling -----
        def _split_emails(csv_str: str):
            return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

        if submitted:
            # Apply per-run config
            cfg.TO_RECIPIENTS = _split_emails(to)
            cfg.CC_RECIPIENTS = _split_emails(cc)
            cfg.USE_ALL_REPORT_KEYS = use_all
            cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

            cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
            cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full
            cfg.EMAIL_MANAGER_REPORT = bool(email_mgr)

            cfg.CALC_SPREADSHEET_ID = calc_id
            cfg.INCOMING_FOLDER_ID = incoming_id
            cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
            cfg.ORDER_REPORT_FOLDER_ID = order_folder
            cfg.REDIRECT_PORT = int(redirect_port)

            cfg.GID_MANAGER_PDF = gid_mgr
            cfg.GID_ORDER_CSV = gid_order
            cfg.LOCATION_SHEET_TITLE = loc_sheet
            cfg.LOCATION_NAMED_RANGE = loc_range
            cfg.TIMESTAMP_TZ = tz
            cfg.TIMESTAMP_FMT = tfmt
            
            cfg.OUTPUT_TIME_TO_LIFE = int(output_ttl)
            cfg.FAILED_INPUT_TIME_TO_LIFE = int(failed_input_ttl)

            cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK = bool(use_rollover)
            cfg.START_DAY_OF_WEEK = start_dow
            cfg.END_DAY_OF_WEEK = end_dow


            # Per-key recipients from editor
            
            rk_map = {}
            for r in edited_rows:
                key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
                emails_csv = r.get("Emails (comma)") or ""
                emails = _parse_emails(emails_csv)
                if key and emails:
                    rk_map[key] = emails
            cfg.REPORT_KEY_RECIPIENTS = rk_map

            # --- ADD: warnings before kicking off the run ---
            # 1) Warn if use_all is OFF and no explicit keys provided
            requested_keys = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]
            if not use_all and not requested_keys:
                st.warning(
                    "You left **Use all Report Keys** OFF but provided **no keys** to run. "
                    "No per‑key outputs will be generated unless you add keys.",
                    icon="⚠️"
                )

            # 2) Warn if no recipients anywhere (TO, DEFAULT_ORDER, or per‑key map)
            any_to = bool(_parse_emails(to) or (cfg.TO_RECIPIENTS or []))
            any_default = bool(cfg.DEFAULT_ORDER_RECIPIENTS or [])
            any_per_key = bool(cfg.REPORT_KEY_RECIPIENTS)
            if not (any_to or any_default or any_per_key):
                st.warning(
                    "No recipients are defined: **TO**, **DEFAULT_ORDER_RECIPIENTS**, and **Per‑Report‑Key** are all empty. "
                    "Emails will not be sent.",
                    icon="⚠️"
                )

            # 3) Surface per-key table issues detected earlier
            if rk_issues:
                st.warning(
                    "Per‑report‑key recipient issues detected above. These may prevent emails from sending correctly:\n\n- "
                    + "\n- ".join(rk_issues),
                    icon="⚠️"
                )
            # --- END ADD ---


            # --- Save edited defaults to Drive JSON (optional) ---
            if save_drive_defaults:
                try:
                    # Ensure we have a user token first
                    creds = load_valid_token(cfg.SCOPES)
                    if not creds:
                        st.error("Not authenticated. Please complete Google sign‑in first (top of page).")
                    else:
                        # Drive service
                        _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                        # What we persist (the fields you asked to move out of Secrets)
                        drive_defaults = {
                            "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
                            "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
                            "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
                            "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,

                            "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
                            "GID_ORDER_CSV": cfg.GID_ORDER_CSV,

                            "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
                            "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,

                            "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
                            "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,

                            "OUTPUT_TIME_TO_LIFE": cfg.OUTPUT_TIME_TO_LIFE,
                            "FAILED_INPUT_TIME_TO_LIFE": cfg.FAILED_INPUT_TIME_TO_LIFE,

                            "TO_RECIPIENTS": cfg.TO_RECIPIENTS,   # lists are fine; JSON keeps types
                            "CC_RECIPIENTS": cfg.CC_RECIPIENTS,

                            "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
                            "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,

                            "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,

                            "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,

                            "INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL": cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL,
                            "SEND_SEPARATE_FULL_ORDER_EMAIL": cfg.SEND_SEPARATE_FULL_ORDER_EMAIL,
                            "EMAIL_MANAGER_REPORT": cfg.EMAIL_MANAGER_REPORT,

                            "USE_AUTO_ROLLOVER_IF_ONE_WEEK" : cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK,
                            "START_DAY_OF_WEEK" : cfg.START_DAY_OF_WEEK,
                            "END_DAY_OF_WEEK" : cfg.END_DAY_OF_WEEK
                        }

                        # If you have CONFIG_FILE_ID in Secrets, we update that exact file.
                        # Otherwise we'll upsert a file named 'favtrip_config.json' and return its id.
                        CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
                        new_id = save_config_to_drive(
                            drive,
                            drive_defaults,
                            file_id=CONFIG_FILE_ID or None,   # update/pin if set, else upsert by name
                            # parent_folder_id=None,          # optional: set a folder id to create under
                        )

                        st.success(f"Saved defaults to Drive config (file id: {new_id}).")
                        if not CONFIG_FILE_ID:
                            st.info(
                                "Tip: add this ID to Streamlit Secrets as `CONFIG_FILE_ID` to pin the same file for all runs:\n"
                                f"`{new_id}`"
                            )
                except Exception as e:
                    st.error(f"Failed to save defaults to Drive: {e}")

            # If user checked "Force Google re-auth for this run", kick them into auth gating first.
            if cfg.FORCE_REAUTH:
                clear_token()
                try:
                    flow, url = start_web_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                    st.session_state.oauth_flow = flow
                    st.session_state.oauth_url = url
                    st.session_state.auth_required = True
                    st.info("Re-auth required for this run. Open the URL shown in the Authentication panel.")
                    _rerun()
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")
            else:
                # --- Live run with timer + last log (no full log after completion) ---
                # Write all logs to last_run.log; overwrite on each run
                logger = StatusLogger(print_to_console=True, file_path="last_run.log", overwrite=True)
                result_holder = {"value": None, "error": None}

                def _runner():
                    try:
                        result_holder["value"] = run_pipeline(cfg, logger=logger)
                    # Catch BaseException so SystemExit / KeyboardInterrupt are captured too
                    except BaseException as e:
                        result_holder["error"] = e

                t0 = time.perf_counter()
                th = threading.Thread(target=_runner, daemon=True)
                th.start()

                with st.status("Running pipeline…", expanded=True) as status:
                    timer_ph = st.empty()
                    lastlog_ph = st.empty()
                    while th.is_alive():
                        elapsed = int(time.perf_counter() - t0)
                        timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                        lastlog_ph.markdown(f"**Last:** {logger.last_line()}")
                        time.sleep(0.5)

                    th.join()
                    elapsed = int(time.perf_counter() - t0)
                    timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                    lastlog_ph.markdown(f"**Last:** {logger.last_line()}")

                    if result_holder["error"]:
                        err = result_holder["error"]
                        err_text = str(err)

                        # Always lock out Run after a failure until a new upload occurs
                        st.session_state.incoming_uploaded_ok = False

                        # --- Special case: "Please only upload 1 or 2 full weeks of data"
                        weeks_err = "Please only upload 1 or 2 full weeks of data" in err_text
                        if weeks_err:
                            # Route to the simple locked UI (error + Retry)
                            st.session_state["file_error"] = err_text
                            st.session_state["incoming_locked"] = True
                            status.update(label="❌ Invalid file (1–2 full weeks required)", state="error")
                            _rerun()
                            st.stop()

                        # --- Generic error path: show details here and freeze the UI
                        st.error("Run failed.")
                        try:
                            with st.expander("Error details", expanded=True):
                                st.exception(err)
                        except Exception:
                            # Fallback if exception rendering fails
                            with st.expander("Error details", expanded=True):
                                st.write(err_text)

                        # Freeze the UI until the user clicks Retry (after the form)
                        st.session_state["freeze_ui"] = True
                        status.update(label="❌ Failed", state="error")
                        

                    else:
                        result = result_holder["value"]
                        if result is None:
                            st.error("Run finished without returning a result. Check logs and inputs (IDs, Drive access).")
                            status.update(label="⚠️ No result", state="error")
                        else:
                            st.write("### Outputs")
                            col1, col2, col3 = st.columns(3)
                            col1.metric("Location", result.location)
                            col2.metric("Timestamp", result.timestamp)
                            mm = result.elapsed_seconds
                            col3.metric("Elapsed", f"{mm//3600:02d}:{(mm%3600)//60:02d}:{mm%60:02d}")

                            if getattr(result, "manager_pdf_link", None):
                                st.success(f"Manager PDF: {result.manager_pdf_link}")
                            if getattr(result, "full_order_link", None):
                                st.success(f"Full Order Sheet: {result.full_order_link}")

                                
                            if os.path.exists("last_run.log"):
                                with open("last_run.log", "rb") as f:
                                    st.session_state["last_run_log"] = f.read()
                                    st.session_state["last_run_timestamp"] = result.timestamp


                            st.session_state.incoming_uploaded_ok = False
                            st.session_state.incoming_uploader_version += 1
                            st.session_state.incoming_selected_name = None

                            status.update(label="✅ Completed", state="complete")
                            time.sleep(10)
                            _rerun()


    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.get("offer_log_download", False) and "last_run_log" in st.session_state:
        st.download_button(
            "⬇️ Download full log (last_run.log)",
            st.session_state["last_run_log"],
            file_name=f"last_run_{st.session_state['last_run_timestamp']}.log",
            mime="text/plain",
            use_container_width=True
            )

    # Post-form controls for generic errors
    # If the UI is frozen due to a non-file error, show a single Retry button here.
    if st.session_state.get("freeze_ui", False) and not st.session_state.get("incoming_locked", False):
        retry_cols = st.columns([1, 3, 1])
        with retry_cols[1]:
            if st.button("Retry", type="secondary", use_container_width=True):
                # Unfreeze, but force a new file upload before the next run
                st.session_state["freeze_ui"] = False
                st.session_state["incoming_uploaded_ok"] = False
                st.session_state["incoming_selected_name"] = None
                st.session_state["incoming_uploader_version"] += 1  # resets uploader widget
                _rerun()





# =========================
# App Entrypoint
# =========================

#st.title("🧾 FavTrip Reporting Pipeline")


st.set_page_config(
    page_title="FT Reporting",
    page_icon="🧾",          # emoji or path/URL to an image
    layout="wide",           # "centered" or "wide"
    initial_sidebar_state="expanded",  # "auto", "expanded", "collapsed"
    menu_items={
        "Get Help": "mailto:ryan-morrow@uiowa.edu",
        "Report a bug": "https://github.com/ryan-j-morrow/favtrip_reporting/issues",
        "About": "FavTrip Reporting Pipeline",
    },
)


st.markdown("""
<style>
/* Make only buttons inside a .ft-run-green scope look green */
.ft-run-green .stButton > button {
    background-color: #22c55e !important;  /* green */
    border-color: #16a34a !important;
    color: #ffffff !important;
}
/* Optional: dim disabled state a bit */
.ft-run-green .stButton > button:disabled {
    opacity: .6 !important;
    cursor: not-allowed !important;
}
/* Utility: right-align containers where used */
.ft-right-btn { display:flex; justify-content:flex-end; }
</style>
""", unsafe_allow_html=True)


cfg = Config.load()

# --- Finish OAuth inline when redirect comes back (this is in the NEW TAB) ---
params = st.query_params
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        # Token is saved locally in this new tab's app process
        st.success("✅ Google authentication complete.")

        # Remove code/state from URL
        st.query_params.clear()

        # No messaging back to opener and NO window.close().
        # This tab becomes the main app; just rerun to flip UI.
        st.toast("Signed in. Loading the app…")
        st.rerun()
    except Exception as e:
        st.error(f"OAuth error: {e}")

# If a valid token now exists (in this tab), reload cfg (Drive overlays, etc.)
st.session_state.auth_required = (load_valid_token(cfg.SCOPES) is None)
if not st.session_state.auth_required:
    cfg = Config.load()

# Session state defaults
if "incoming_selected_name" not in st.session_state:
    st.session_state.incoming_selected_name = None
if "incoming_uploaded_ok" not in st.session_state:
    st.session_state.incoming_uploaded_ok = False
if "offer_log_download" not in st.session_state:
    st.session_state.offer_log_download = False
if "offer_log_download" not in st.session_state:
    st.session_state.offer_log_download = False


# Sidebar (always visible)
with st.sidebar:
    st.header("Utilities")

    if st.button("Google Sign Out", type="secondary", use_container_width=True):
        clear_token()
        for key in ["auth_required", "oauth_flow", "oauth_url", "auth_checked"]:
            if key in st.session_state:
                del st.session_state[key]
        try:
            st.rerun()
        except AttributeError:
            st.experimental_rerun()
    
    st.checkbox(
        "Offer full log download after completion",
        key="offer_log_download",
        help="If enabled, a 'Download last_run.log' button appears when a run finishes.")

# Auth gate
if st.session_state.auth_required:
    # ----------------------------
    # Authentication panel (shown only if auth required)
    # ----------------------------
    if st.session_state.auth_required:
        with st.expander("Google Authentication", expanded=True):
            st.caption(
                "Authentication is required before running. "
                "Click **Sign in with Google** to open the consent screen (it will open in a new tab)."
            )

            sign_in_ph = st.empty()
            clicked = sign_in_ph.button("Sign in with Google", type="primary", use_container_width=True)

            if clicked:
                try:
                    auth_url = start_web_oauth(cfg.SCOPES)
                    sign_in_ph.empty()

                    # Friendly message in this (original) tab
                    st.markdown(
                        """
                        <div style="
                            display:flex;align-items:center;justify-content:center;
                            height:55vh;text-align:center;
                            font-family: system-ui, Segoe UI, Roboto, Helvetica, Arial, sans-serif;">
                        <div>
                            <h2 style="margin-bottom:0.5rem;">You're being signed in…</h2>
                            <p style="font-size:1.05rem;opacity:.9;">
                            A new browser tab was opened for Google sign‑in.<br/>
                            <strong>After it completes, continue in that tab.</strong>
                            </p>
                        </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    # Optional: refresh this tab when user returns (not required)
                    html(
                        """
                        <script>
                        document.addEventListener("visibilitychange", function() {
                            if (!document.hidden) { location.reload(); }
                        });
                        </script>
                        """,
                        height=0,
                    )

                    # Open Google auth in a NEW tab (this will ultimately become the main app)
                    html(
                        f"""
                        <script>
                        window.open({json.dumps(auth_url)}, "_blank", "noopener");
                        </script>
                        """,
                        height=0,
                    )

                    st.stop()
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")

            with st.expander("Having trouble?", expanded=False):
                st.write(
                    "- The Google authorization page opens in a **new browser tab**.\n"
                    "- After completing consent, the **new tab** will load the app.\n"
                    "- If you renamed your Streamlit app or URL, ensure the Google OAuth "
                    "Authorized redirect URI matches exactly (including trailing slash)."
                )
            st.stop()

# Optional lock if invalid incoming file was detected
locked = st.session_state.get("incoming_locked", False)
if locked:

    err_msg = st.session_state.get("file_error", "An unknown error occurred.")
    st.error(err_msg)

    if st.button("Retry", type="secondary"):
        st.session_state["incoming_locked"] = False
        _rerun()
    st.stop()

else:
    # Ensure config reflects Drive-based overrides after auth
    cfg = Config.load()
    render_run_form(cfg)
```

---
### file: web_url_credentials.json

```json
{"web":{"client_id":"674901450584-i6ncesf794ecg6u7olebqkag8gavbuc3.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["https://favtripreporting-dev1.streamlit.app/"]}}
```

```

---
### file: documentation/README.md

```markdown
# FavTrip Reporting Pipeline (Refactored)

This is a refactor of the original one-file notebook script into a small, testable package with:

- Configuration via `.env` (with a **per-run UI override** and an **optional** "Update defaults" toggle)
- A local **web UI** built with Streamlit (`ui_streamlit.py`)
- A CLI entrypoint (`cli.py`)
- Clear runtime logging and completion messages

## Quick start

1. **Prereqs**
   - Python 3.10+
   - A Google Cloud OAuth Client ID (Desktop) and its `credentials.json` in the working directory
   - `token.json` will be created after your first auth flow

2. **Install**

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scriptsctivate
pip install -r requirements.txt
cp .env.example .env  # fill in your IDs
```

3. **Run the web UI**

```bash
streamlit run ui_streamlit.py
```

4. **Run from CLI**

```bash
python cli.py --report-keys GROCERY,COFFEE --to you@example.com
```

## Packaging

### Windows `.exe` (CLI)

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline   --add-data "credentials.json;."   --add-data ".env;."   cli.py
```
- Place `token.json` next to the `.exe` after you authorize once (or allow it to be created on first run).

### macOS app (CLI or windowed)

Using **PyInstaller**:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name FavTripPipeline cli.py
# For a GUI stub that launches the Streamlit app in your browser:
pyinstaller --noconfirm --onefile --windowed --name FavTripPipelineUI launcher_streamlit.py
```

> Note: Streamlit runs a local server; packaging it as a single app is possible but you still need `credentials.json` and `.env` available. For a double-click experience, create a small `launcher_streamlit.py` that calls `os.system('streamlit run ui_streamlit.py')` and package that.

### macOS via `py2app` (alternative)

```bash
pip install py2app
python setup_py2app.py py2app
```

See comments in `setup_py2app.py`.

## What the UI exposes by default

- IDs: Calculations spreadsheet, incoming folder, manager & order folders
- Recipients: To/CC, report-key list, flags for emailing behavior
- Auth toggles: force re-auth, redirect port
- Advanced: GIDs, location sheet/range, timezone/format

A small **"Update defaults in .env"** checkbox persists edits; otherwise values apply **only to this run**.

## Notes
- The pipeline still relies on Google OAuth user credentials (`credentials.json` / `token.json`). Keep those files alongside your app.
- The Gmail API sends messages as the signed-in user (`me`).
- Exported PDFs/CSVs are uploaded to the Drive folders you specify.


```

---
### file: documentation/generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd "C:\Users\rjrul\OneDrive - University of Iowa\000 Current Semester\004 BAIS 4150 BAIS Capstone\favtrip_reporting"
#python documentation/generate_code_bundle.py
```

---
### file: documentation/git_workflow.txt

```text

---
1. Create a dev branch

# Switch to dev branch
    git checkout dev

# Push dev branch to GitHub (the -u flag sets upstream tracking)
    git push -u origin dev

---
2. Do all development on dev
Make changes normally, then stage, commit, and push:

    git add .
    git commit -m "your message here"
    git push

All pushes go to dev (not main).

---
3. Push dev → main when ready (production release)
Use a Pull Request on GitHub:
1. Go to the repo on GitHub.
2. You will see a banner offering “Compare & Pull Request” (dev → main).
3. Open the Pull Request.
4. Review and click “Merge Pull Request”.

After merging, update your local main branch:

    git checkout main
    git pull

---
4. Keep dev updated with the latest main
After merging into main, update dev:

    git checkout dev
    git pull origin main

Or:

    git merge main

This prevents dev from drifting behind main.

---
5. Summary
- main = production, always stable
- dev = development, experimental work
- Merge dev → main only after testing

```

---
### file: documentation/requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit
openpyxl
pandas
```

---
### file: favtrip/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: favtrip/config.py

```python
from __future__ import annotations

import os
import json
from dataclasses import dataclass, asdict
from typing import Any, Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: Any = None) -> Any:
    """
    Read from Streamlit secrets if present, else env var, else default.
    Does not raise if key missing; returns `default`.
    """
    try:
        import streamlit as st  # imported lazily to avoid hard dependency
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets.get(key, default)
    except Exception:
        pass
    return os.getenv(key, default)

_TRUE = {"1", "true", "yes", "on", "y", "t"}

def _coerce_bool(v: Any, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default

def _coerce_csv(v: Any) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]

def _coerce_json(v: Any) -> Dict[str, Any]:
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}

# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str
    USER_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"
    TEMPLATE_UPDATE_RANGE: str = "_update"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True
    EMAIL_MANAGER_REPORT: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300

    # Advanced intake settings
    USE_AUTO_ROLLOVER_IF_ONE_WEEK: bool = True
    START_DAY_OF_WEEK: str = "Sunday"    # Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Any
    END_DAY_OF_WEEK: str = "Saturday"    # Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Any

    # Cleanup
    OUTPUT_TIME_TO_LIFE: int = 30
    FAILED_INPUT_TIME_TO_LIFE: int = 1
    USER_TIME_TO_LIFE: int = 90

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        """
        Load config from Streamlit secrets (preferred on cloud) or from env/.env (local dev),
        then overlay any values found in a Drive-backed JSON config (optional).
        Secrets may be typed (bool/list/dict), so we coerce safely.
        """
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)

        cfg = Config(
            CALC_SPREADSHEET_ID=str(_get_secret("CALC_SPREADSHEET_ID", "")),
            INCOMING_FOLDER_ID=str(_get_secret("INCOMING_FOLDER_ID", "")),
            MANAGER_REPORT_FOLDER_ID=str(_get_secret("MANAGER_REPORT_FOLDER_ID", "")),
            ORDER_REPORT_FOLDER_ID=str(_get_secret("ORDER_REPORT_FOLDER_ID", "")),
            USER_FOLDER_ID=str(_get_secret("USER_FOLDER_ID", "")),

            GID_MANAGER_PDF=str(_get_secret("GID_MANAGER_PDF", "1921812573")),
            GID_ORDER_CSV=str(_get_secret("GID_ORDER_CSV", "1875928148")),
            LOCATION_SHEET_TITLE=str(_get_secret("LOCATION_SHEET_TITLE", "REFR: Values")),
            LOCATION_NAMED_RANGE=str(_get_secret("LOCATION_NAMED_RANGE", "_locations")),
            TEMPLATE_UPDATE_RANGE=str(_get_secret("TEMPLATE_UPDATE_RANGE", "_update")),
            TIMESTAMP_TZ=str(_get_secret("TIMESTAMP_TZ", "America/Chicago")),
            TIMESTAMP_FMT=str(_get_secret("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p")),

            OUTPUT_TIME_TO_LIFE=int(_get_secret("OUTPUT_TIME_TO_LIFE", 30)),
            FAILED_INPUT_TIME_TO_LIFE=int(_get_secret("FAILED_INPUT_TIME_TO_LIFE", 1)),
            USER_TIME_TO_LIFE=int(_get_secret("USER_TIME_TO_LIFE", 1)),

            TO_RECIPIENTS=_coerce_csv(_get_secret("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_coerce_csv(_get_secret("CC_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_coerce_bool(_get_secret("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=[s.upper() for s in _coerce_csv(_get_secret("REPORT_KEY_RUN_LIST", ""))],
            REPORT_KEY_RECIPIENTS=_coerce_json(_get_secret("REPORT_KEY_RECIPIENTS", "{}")),
            DEFAULT_ORDER_RECIPIENTS=_coerce_csv(_get_secret("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_coerce_bool(
                _get_secret("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_coerce_bool(
                _get_secret("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),
            EMAIL_MANAGER_REPORT=_coerce_bool(_get_secret("EMAIL_MANAGER_REPORT", "true")),

            SCOPES=_coerce_csv(
                _get_secret(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,"
                    "https://www.googleapis.com/auth/spreadsheets,"
                    "https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_coerce_bool(_get_secret("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(str(_get_secret("REDIRECT_PORT", "58285")) or "58285"),
            HTTP_TIMEOUT_SECONDS=int(str(_get_secret("HTTP_TIMEOUT_SECONDS", "300")) or "300"),

            USE_AUTO_ROLLOVER_IF_ONE_WEEK=_coerce_bool(_get_secret("USE_AUTO_ROLLOVER_IF_ONE_WEEK", "true")),
            START_DAY_OF_WEEK=str(_get_secret("START_DAY_OF_WEEK", "Sunday")),
            END_DAY_OF_WEEK=str(_get_secret("END_DAY_OF_WEEK", "Saturday")),
        )

        # Optional overlay from Drive JSON config (if creds + file present)
        try:
            import streamlit as st
            from favtrip.google_client import load_valid_token, services
            from favtrip.config_store import load_config_from_drive

            CONFIG_FILE_ID = ""
            if hasattr(st, "secrets"):
                CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

            creds = load_valid_token(cfg.SCOPES)
            if creds:
                _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
                overrides = load_config_from_drive(drive, CONFIG_FILE_ID or None)  # {} if not found
                if isinstance(overrides, dict) and overrides:
                    for k, v in overrides.items():
                        if hasattr(cfg, k):
                            setattr(cfg, k, v)
        except Exception:
            # Fail-open: if Drive/token not ready yet, just return base cfg
            pass

        return cfg

    # -------------------------------------------------------------------------
    # .env serialization (optional helper)
    # -------------------------------------------------------------------------
    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")

```

---
### file: favtrip/config_store.py

```python
# favtrip/config_store.py
from __future__ import annotations
import io
import json
from typing import Any, Dict, Optional
from googleapiclient.discovery import Resource
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

DEFAULT_CONFIG_FILENAME = "favtrip_config.json"
DEFAULT_MIMETYPE = "application/json"

def load_config_from_drive(drive: Resource, file_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Read the JSON config stored in Google Drive.
    If file_id is None, try to discover the newest file named DEFAULT_CONFIG_FILENAME.
    Returns {} if the file doesn't exist or is empty/invalid JSON.
    """
    # Discover by name if a specific id wasn't provided
    if not file_id:
        resp = drive.files().list(
            q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
            orderBy="modifiedTime desc",
            pageSize=1,
            fields="files(id,name,modifiedTime)"
        ).execute() or {}
        files = resp.get("files", [])
        if not files:
            return {}
        file_id = files[0]["id"]

    # Stream download the file
    request = drive.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    raw = buf.getvalue().decode("utf-8", errors="replace").strip()
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except Exception:
        return {}

def save_config_to_drive(
    drive: Resource,
    data: Dict[str, Any],
    file_id: Optional[str] = None,
    parent_folder_id: Optional[str] = None
) -> str:
    """
    Write JSON config to Google Drive.
    - If file_id provided, update that file.
    - Else upsert (update if found by name, otherwise create) DEFAULT_CONFIG_FILENAME.
    Returns the Drive file ID.
    """
    payload = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(payload), mimetype=DEFAULT_MIMETYPE, resumable=True)

    if file_id:
        updated = drive.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        return updated["id"]

    # Try to find an existing file by name to update
    resp = drive.files().list(
        q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
        orderBy="modifiedTime desc",
        pageSize=1,
        fields="files(id,name)"
    ).execute() or {}
    files = resp.get("files", [])
    if files:
        fid = files[0]["id"]
        updated = drive.files().update(fileId=fid, media_body=media).execute()
        return updated["id"]

    # Create a new file
    meta = {"name": DEFAULT_CONFIG_FILENAME}
    if parent_folder_id:
        meta["parents"] = [parent_folder_id]

    created = drive.files().create(
        body=meta,
        media_body=media,
        fields="id,name"
    ).execute()
    return created["id"]
```

---
### file: favtrip/drive_utils.py

```python
from __future__ import annotations
import io
from datetime import datetime, timedelta, timezone
from googleapiclient.http import MediaIoBaseUpload


def find_latest_sheet(drive_svc, folder_id: str):
    q = (
        f"'{folder_id}' in parents and "
        "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    )
    resp = drive_svc.files().list(
        q=q, orderBy="createdTime desc", pageSize=1,
        fields="files(id,name,createdTime)"
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def upload_to_drive(drive_svc, data: bytes, name: str, mime: str, folder_id: str, to_sheet: bool=False):
    meta = {"name": name, "parents": [folder_id]}
    if to_sheet:
        meta["mimeType"] = "application/vnd.google-apps.spreadsheet"
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
    return drive_svc.files().create(
        body=meta, media_body=media, fields="id,name,mimeType,webViewLink"
    ).execute()

def _rfc3339(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")

def trash_file(drive, file_id: str):
    return drive.files().update(fileId=file_id, body={"trashed": True}).execute()

def cleanup_folder_by_age(drive, folder_id: str, days: int, logger=None):
    if days <= 0:
        return 0

    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    cutoff_str = _rfc3339(cutoff)

    q = (
        f"'{folder_id}' in parents and trashed=false "
        f"and createdTime < '{cutoff_str}'"
    )

    trashed = 0
    page_token = None

    while True:
        resp = drive.files().list(
            q=q,
            pageSize=1000,
            orderBy="createdTime asc",
            fields="nextPageToken, files(id,name,createdTime)",
            pageToken=page_token
        ).execute() or {}

        for f in resp.get("files", []):
            try:
                trash_file(drive, f["id"])
                trashed += 1
                if logger:
                    logger.info(f"Trashed file: {f['name']} ({f['id']})")
            except Exception as e:
                if logger:
                    logger.warn(f"Failed to trash {f['id']}: {e}")

        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    return trashed
```

---
### file: favtrip/gmail_utils.py

```python
from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {ts} – {location}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Automated")
    msg.add_alternative(
        f"<p>Hi team,</p><p>Manager Report ({location})</p>"
        f"<a href='{pdf_link}'>Backup Link</a>", subtype="html"
    )
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)

```

---
### file: favtrip/google_client.py

```python
from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail
```

---
### file: favtrip/logger.py

```python
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass
```

---
### file: favtrip/pipeline.py

```python
from __future__ import annotations
import pandas
import csv
import io
import re
import requests
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

from io import BytesIO
from openpyxl import load_workbook, Workbook


from .config import Config
from .google_client import get_credentials, services
from .sheets_utils import (
    delete_sheet, copy_sheet_as, copy_first_sheet_as, refresh_sheets_with_prefix,
    get_value, first_gid,
    get_first_sheet_meta, get_values_2d, add_blank_sheet,
    add_or_replace_sheet, put_values_2d, _force_column_as_text, delete_row_indices, delete_rows_range
)
from .drive_utils import find_latest_sheet, upload_to_drive, _rfc3339, trash_file, cleanup_folder_by_age
from .gmail_utils import send_email, email_manager_report

CSV_MIME = "text/csv"


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


import requests
from io import BytesIO
from openpyxl import Workbook


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}&gid={gid}"
    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content


def timestamp_now(tz: str, fmt: str) -> str:
    return datetime.now(ZoneInfo(tz)).strftime(fmt)

class IncomingDataValidationError(Exception):
    """Raised when the incoming report is not 1 or 2 full weeks as configured."""
    pass

_DOW_MAP = {
    "Monday": 0, "Tuesday": 1, "Wednesday": 2, "Thursday": 3,
    "Friday": 4, "Saturday": 5, "Sunday": 6, "Any": None,
}

def _parse_sheet_date(cell: str | int | float):
    """
    Parse a Google Sheets date cell into a date (drops time if present).
    Accepts:
      - Google serial numbers (days since 1899-12-30)
      - Date strings: YYYY-MM-DD, MM/DD/YYYY, MM/DD/YY
      - DateTime strings: with 12h or 24h time, with or without seconds, AM/PM
      - ISO strings (date or datetime)

    Returns: datetime.date or None if unparseable.
    """
    from datetime import datetime, timedelta

    if cell is None or cell == "":
        return None

    # --- 1) Numeric serial (Google Sheets) ---
    try:
        if isinstance(cell, (int, float)) or (isinstance(cell, str) and cell.replace(".", "", 1).isdigit()):
            serial = float(cell)
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=serial)).date()
    except Exception:
        pass

    s = str(cell).strip()

    # Quick strip for weird whitespace
    s = " ".join(s.split())

    # If a timezone suffix or trailing text exists, try to isolate the datetime token
    # (We keep it simple: split on two spaces or take first token that contains '/')
    if " " in s and "/" in s:
        # Nothing fancy; the format tries below will accept the full string if they match
        pass

    # --- 2) Try common datetime formats (12h and 24h), with or without seconds ---
    dt_formats = [
        "%m/%d/%Y %I:%M:%S %p",  # 03/01/2026 12:03:45 AM
        "%m/%d/%Y %I:%M %p",     # 03/01/2026 12:03 AM
        "%m/%d/%Y %H:%M:%S",     # 03/01/2026 00:03:45
        "%m/%d/%Y %H:%M",        # 03/01/2026 00:03
        "%Y-%m-%d %H:%M:%S",     # 2026-03-01 00:03:45
        "%Y-%m-%d %H:%M",        # 2026-03-01 00:03
        "%Y-%m-%dT%H:%M:%S",     # 2026-03-01T00:03:45
        "%Y-%m-%dT%H:%M",        # 2026-03-01T00:03
    ]
    for fmt in dt_formats:
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue

    # --- 3) Try date-only formats ---
    date_formats = [
        "%Y-%m-%d",   # 2026-03-01
        "%m/%d/%Y",   # 03/01/2026
        "%m/%d/%y",   # 03/01/26
    ]
    for fmt in date_formats:
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue

    # --- 4) Last resort: Python ISO parser (handles 'YYYY-MM-DD' and full ISO datetimes) ---
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        pass

    # --- 5) If still not parsed, try taking only the date token before a space ---
    try:
        token = s.split(" ")[0]
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
            try:
                return datetime.strptime(token, fmt).date()
            except Exception:
                continue
    except Exception:
        pass

    return None

def _find_header_and_date_col(values2d):
    """
    Find the header row whose first cell == 'Store', and the 'Date' column index.
    Returns (header_row_ix, date_col_ix) or (None, None).
    """
    header_ix = None
    for r, row in enumerate(values2d):
        c0 = (row[0].strip() if row and isinstance(row[0], str) else row[0] if row else "")
        if str(c0).strip().lower() == "store":
            header_ix = r
            break
    if header_ix is None:
        return None, None
    headers = [str(h).strip() for h in values2d[header_ix]]
    date_col_ix = None
    for c, h in enumerate(headers):
        if h.lower() == "date":
            date_col_ix = c
            break
    return header_ix, date_col_ix

def _collect_unique_dates(values2d, header_ix, date_cix):
    dates = []
    for r in range(header_ix + 1, len(values2d)):
        row = values2d[r]
        if date_cix >= len(row):
            continue
        d = _parse_sheet_date(row[date_cix])
        if d:
            dates.append(d)
    return sorted(set(dates))

def _check_week_boundaries(unique_dates, start_dow, end_dow):
    """Validate first/last weekday (unless set to Any). Return (earliest, latest)."""
    if not unique_dates:
        raise IncomingDataValidationError("No dates found in incoming report.")
    earliest, latest = unique_dates[0], unique_dates[-1]
    s_ok = (_DOW_MAP[start_dow] is None) or (earliest.weekday() == _DOW_MAP[start_dow])
    e_ok = (_DOW_MAP[end_dow]   is None) or (latest.weekday()   == _DOW_MAP[end_dow])
    error_text = None
    if not (s_ok and e_ok):
        error_text = f"Please only upload 1 or 2 full weeks of data. The first day of week included in the report should be {start_dow} and the last day of week included in the report should be {end_dow}"
        raise IncomingDataValidationError(
            error_text
        )
    return earliest, latest, error_text

def _plan_weeks(unique_dates):
    """
    Decide if we have one or two weeks by count of unique calendar days.
    Returns ('one', set7) or ('two', (set7_oldest, set7_newest)).
    """
    if len(unique_dates) == 7:
        return "one", set(unique_dates)
    if len(unique_dates) == 14:
        return "two", (set(unique_dates[:7]), set(unique_dates[7:]))
    # Not 7 or 14
    raise IncomingDataValidationError(
        "Please only upload 1 or 2 full weeks of data. The first day of week included in the report should be XXX and the last day of week included in the report should be YYY"
    )

def _trim_header_if_needed(svc, spreadsheet_id: str, sheet_id: int, values2d, header_ix):
    """Ensure header is at row 0 by deleting rows above it."""
    if header_ix and header_ix > 0:
        delete_rows_range(svc, spreadsheet_id, sheet_id, 0, header_ix)

def _filter_rows_to_dates(svc, spreadsheet_id: str, sheet_id: int, values2d, header_ix, date_cix, keep_dates_set):
    """Delete all non-header rows whose Date is not in keep_dates_set."""
    bad_rows = []
    for r in range(header_ix + 1, len(values2d)):
        row = values2d[r]
        d = _parse_sheet_date(row[date_cix] if date_cix < len(row) else None)
        if (d is None) or (d not in keep_dates_set):
            bad_rows.append(r)
    delete_row_indices(svc, spreadsheet_id, sheet_id, bad_rows)


import re

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _clean_emails(items):
    """
    Accepts a list or a comma-separated string and returns a list of valid emails.
    Trailing commas and blanks are removed. Invalid tokens are dropped silently.
    """
    if items is None:
        return []
    if isinstance(items, str):
        items = [p.strip() for p in items.split(",")]
    return [e for e in (p.strip() for p in items) if e and _EMAIL_RE.match(e)]

def _fallback_recipients(hint, *candidates):
    """
    Return the first non-empty, valid recipient list from the provided candidates.
    If all candidates are empty/invalid, raise a friendly error.
    """
    for c in candidates:
        cleaned = _clean_emails(c)
        if cleaned:
            return cleaned
    # Nothing usable found:
    raise ValueError(
        f"No valid recipients available for: {hint}. "
        f"Please provide at least one email in the UI or .env "
        f"(TO_RECIPIENTS, DEFAULT_ORDER_RECIPIENTS, or per-report-key)."
    )


@dataclass
class RunResult:
    ok: bool
    elapsed_seconds: int
    location: str
    timestamp: str
    manager_pdf_link: str | None
    full_order_link: str | None


def run_pipeline(cfg: Config, logger=None) -> RunResult:
    import time
    start = time.perf_counter()

    if logger:
        logger.info("Authorizing with Google APIs…")
    creds = get_credentials(cfg.SCOPES, cfg.REDIRECT_PORT, cfg.FORCE_REAUTH)
    sheets_svc, drive_svc, gmail_svc = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    if logger:
        logger.info("Google services ready")

    # Step 1: latest incoming
    if logger:
        logger.info("Finding latest incoming spreadsheet…")
    latest = find_latest_sheet(drive_svc, cfg.INCOMING_FOLDER_ID)
    if not latest:
        raise SystemExit("No incoming report found.")
    new_report_id = latest["id"]
    if logger:
        logger.info(f"Latest incoming: {latest['name']} ({new_report_id})")

    # ---- NEW: Validate incoming weeks & plan actions (no workbook changes yet) ----
    if logger:
        logger.info("Validating incoming report (header, dates, week boundaries)…")
    first_title, first_sid = get_first_sheet_meta(sheets_svc, new_report_id)
    values = get_values_2d(sheets_svc, new_report_id, first_title, "A:Z")

    h_ix, d_cix = _find_header_and_date_col(values)
    if h_ix is None or d_cix is None:
        raise IncomingDataValidationError(
            "Unable to locate header ('Store' in A1) and/or 'Date' column in the incoming report."
        )

    unique_dates = _collect_unique_dates(values, h_ix, d_cix)

    if logger:
        logger.info(f"Found {len(unique_dates)} unique date(s) in incoming report")

    check_outputs = _check_week_boundaries(unique_dates, cfg.START_DAY_OF_WEEK, cfg.END_DAY_OF_WEEK)
    plan_kind, plan_payload = _plan_weeks(unique_dates)

    # Step 2: prep calculations workbook (branch by plan)
    if logger:
        logger.info("Preparing calculations workbook…")

    if plan_kind == "two":
        # Two weeks → build values in memory and write each in a single call
        if logger:
            logger.info("Detected 2 weeks; writing 'Last Week' (oldest 7) and 'Current Week' (newest 7) without row deletions")

        # Source header & body (we already loaded 'values' from the first sheet)
        header = [str(h) for h in values[h_ix]]
        body_rows = values[h_ix + 1 :]

        def _slice_rows(rows, date_cix, keep_dates: set):
            out = []
            for row in rows:
                d = _parse_sheet_date(row[date_cix] if date_cix < len(row) else None)
                if d and d in keep_dates:
                    out.append(row)
            return out

        keep_oldest7, keep_newest7 = plan_payload  # sets of dates from _plan_weeks
        last_week_rows = _slice_rows(body_rows, d_cix, keep_oldest7)
        current_week_rows = _slice_rows(body_rows, d_cix, keep_newest7)

        # Create fresh target sheets
        add_or_replace_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
        add_or_replace_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")

        # Force column 'Scan Code' to be text with a prefixed apostrophe
        last_week_rows = _force_column_as_text(header, last_week_rows, "Scan Code")
        current_week_rows = _force_column_as_text(header, current_week_rows, "Scan Code")

        # Bulk write (header + rows) → 1 write per sheet
        put_values_2d(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week", [header] + last_week_rows)
        put_values_2d(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", [header] + current_week_rows)

    elif plan_kind == "one" and cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK:
        # One week + rollover ON → current behavior
        if logger:
            logger.info("Detected 1 week; auto-rollover enabled → copying old Current→Last and inserting new Current")
        delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
        try:
            copy_sheet_as(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", "Last Week")
            if logger:
                logger.info("Copied old 'Current Week' to 'Last Week'")
        except Exception:
            if logger:
                logger.warn("No 'Current Week' sheet exists to copy")
        delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
        copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")

        # Trim header for Current Week
        meta = sheets_svc.spreadsheets().get(spreadsheetId=cfg.CALC_SPREADSHEET_ID).execute()
        cw_sid = next(s["properties"]["sheetId"] for s in meta["sheets"] if s["properties"]["title"] == "Current Week")
        _trim_header_if_needed(sheets_svc, cfg.CALC_SPREADSHEET_ID, cw_sid, values, h_ix)

    else:
        # One week + rollover OFF → Current Week only; Last Week blank
        if logger:
            logger.info("Detected 1 week; auto-rollover disabled → Current only, Last Week blank")
        delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
        delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
        add_blank_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
        copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")

        meta = sheets_svc.spreadsheets().get(spreadsheetId=cfg.CALC_SPREADSHEET_ID).execute()
        cw_sid = next(s["properties"]["sheetId"] for s in meta["sheets"] if s["properties"]["title"] == "Current Week")
        _trim_header_if_needed(sheets_svc, cfg.CALC_SPREADSHEET_ID, cw_sid, values, h_ix)

    # Refresh reference sheets (unchanged)
    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ')…")
    refresh_sheets_with_prefix(sheets_svc, cfg.CALC_SPREADSHEET_ID, prefix="REFR: ", logger=logger)

    # Step 3: read location code
    location = get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.LOCATION_NAMED_RANGE)
    ts = timestamp_now(cfg.TIMESTAMP_TZ, cfg.TIMESTAMP_FMT)
    if logger:
        logger.info(f"Location: {location}; Timestamp: {ts}")

    # Step 4A: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_MANAGER_PDF, "pdf")
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 4B: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_ORDER_CSV, "csv")

    # Step 4C: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_{ts}_{location}_FULL.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_{ts}_{location}_FULL.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4D: Create per-report-key outputs (CSV) and email

    # --- Parse the master CSV into rows of dicts ---

    # master_csv_bytes = export_sheet(..., "csv")
    text = master_csv_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))

    rows_list = list(reader)
    if not rows_list:
        raise RuntimeError("CSV has no rows.")

    headers = [h.strip() for h in rows_list[0]]
    if not headers:
        raise RuntimeError("CSV has no header.")

    # Find 'report_key' column, case-insensitive
    lower_idx = {h.lower(): i for i, h in enumerate(headers)}
    if "report_key" not in lower_idx:
        raise RuntimeError("Report_Key column missing.")
    report_idx = lower_idx["report_key"]

    # Materialize rows as list[dict]
    rows = []
    for row in rows_list[1:]:
        rows.append({headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))})

    # Group by report key
    groups: dict[str, list[dict]] = {}
    for r in rows:
        key = (str(r.get(headers[report_idx]) or "").strip()) or "UNASSIGNED"
        groups.setdefault(key.upper(), []).append(r)


    
    for key, key_rows in groups.items():
        if not cfg.USE_ALL_REPORT_KEYS and key.upper() not in (cfg.REPORT_KEY_RUN_LIST or []):
            continue

        # Build CSV text in memory
        sio = io.StringIO()
        w = csv.writer(sio, lineterminator="\n")
        w.writerow(headers)
        for rr in key_rows:
            w.writerow([rr.get(h, "") for h in headers])

        key_csv_bytes = sio.getvalue().encode("utf-8")
        tag = clean_tag(key)  # already upper

        csv_name = f"Order_Report_{ts}_{location}_{tag}.csv"

        # Upload CSV to Drive; conversion to Google Sheet happens via to_sheet=True
        created = upload_to_drive(
            drive_svc, key_csv_bytes, csv_name,
            CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True
        )
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)


        # Export the Google Sheet as PDF (unchanged behavior)
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{ts}_{location}_{tag}.pdf"

        # Prefer per‑key recipients; else default; else TO
        candidates = None
        if cfg.REPORT_KEY_RECIPIENTS:
            # The keys in UI are upper-cased; normalize here too
            candidates = cfg.REPORT_KEY_RECIPIENTS.get(tag)

        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.DEFAULT_ORDER_RECIPIENTS,
            cfg.TO_RECIPIENTS,
        )

        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – {tag}"
        msg["From"] = "me"
        msg["To"] = ", ".join(recipients)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)

        # Keep the Google Sheet link; PDF remains the attached artifact
        msg.set_content(
            f"Hi {key} team,\nYour order report is ready. \n"
            f"Google Sheet: {created.get('webViewLink')}\n"
            f"Attached: {pdfname}\n—Automated"
        )
        msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=pdfname)
        if cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL:
            msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)

        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info(f"Emailed {tag}")

    # Step 4E: Send Manager Report (guarded by cfg.EMAIL_MANAGER_REPORT)
    if getattr(cfg, "EMAIL_MANAGER_REPORT", True):
        to_list = _fallback_recipients("Manager Report (TO_RECIPIENTS)", cfg.TO_RECIPIENTS)
        cc_list = _clean_emails(cfg.CC_RECIPIENTS)
        email_manager_report(
            gmail_svc, "me", to_list, cc_list,
            pdf_name, pdf_bytes, manager_link, ts, location
        )
        if logger:
            logger.info("Manager email sent")
    else:
        if logger:
            logger.info("Manager email skipped by configuration (EMAIL_MANAGER_REPORT = False)")

    

    # Step 4F: Send Full Order if needed
    full_link = full_created.get('webViewLink')
    if cfg.SEND_SEPARATE_FULL_ORDER_EMAIL:
        to_full = _fallback_recipients("FULL order", cfg.DEFAULT_ORDER_RECIPIENTS, cfg.TO_RECIPIENTS)
        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – FULL"
        msg["From"] = "me"
        msg["To"] = ", ".join(to_full)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi team,\nFULL order report is ready.\nSheet: {full_link}\nAttached: {full_pdf_name}\n—Automated")
        msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info("FULL order email sent")
    else:
        if logger:
            logger.info("Separate full order email disabled")

    
    # Step 4G: File Cleanup

    try:
        if logger:
            logger.info("Cleaning up used incoming file…")
        trash_file(drive_svc, new_report_id)

        if logger:
            logger.info("Cleaning old incoming files…")
        cleanup_folder_by_age(
            drive_svc,
            cfg.INCOMING_FOLDER_ID,
            cfg.FAILED_INPUT_TIME_TO_LIFE,
            logger
        )

        if logger:
            logger.info("Cleaning old output files…")
        for folder in [
            cfg.MANAGER_REPORT_FOLDER_ID,
            cfg.ORDER_REPORT_FOLDER_ID
        ]:
            cleanup_folder_by_age(
                drive_svc,
                folder,
                cfg.OUTPUT_TIME_TO_LIFE,
                logger
            )
        
        if logger:
            logger.info("Cleaning old calculation files…")
            cleanup_folder_by_age(
                drive_svc,
                cfg.USER_FOLDER_ID,
                cfg.USER_TIME_TO_LIFE,
                logger
            )

    except Exception as e:
        if logger:
            logger.warn(f"Housekeeping failed: {e}")

    elapsed = int(time.perf_counter() - start)
    if logger:
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        logger.info(f"Run completed in {h:02d}:{m:02d}:{s:02d}")

    return RunResult(True, elapsed, location, ts, manager_link, full_link)

```

---
### file: favtrip/sheets_utils.py

```python
from __future__ import annotations
import random
import time
from typing import Any, Dict, List


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def delete_sheet(svc, spreadsheet_id: str, title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), title)
    if s:
        body = {"requests": [{"deleteSheet": {"sheetId": s["sheetId"]}}]}
        svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def copy_sheet_as(svc, spreadsheet_id: str, src_title: str, new_title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), src_title)
    if not s:
        return None
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=spreadsheet_id,
        sheetId=s["sheetId"],
        body={"destinationSpreadsheetId": spreadsheet_id}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return new_id


def copy_first_sheet_as(svc, src_spreadsheet: str, dest_spreadsheet: str, new_title: str):
    meta = svc.spreadsheets().get(spreadsheetId=src_spreadsheet).execute()
    first_id = meta["sheets"][0]["properties"]["sheetId"]
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=src_spreadsheet,
        sheetId=first_id,
        body={"destinationSpreadsheetId": dest_spreadsheet}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=dest_spreadsheet, body=body).execute()
    return new_id


def refresh_sheets_with_prefix(svc, spreadsheet_id: str, prefix: str = "REFR: ", retries: int = 5, logger=None):
    sheets = list_sheets(svc, spreadsheet_id)
    targets = [s["properties"] for s in sheets if s["properties"]["title"].startswith(prefix)]
    for idx, t in enumerate(targets, start=1):
        body = {"requests": [{
            "findReplace": {
                "find": "=",
                "replacement": "=",
                "includeFormulas": True,
                "sheetId": t["sheetId"]
            }
        }]}
        attempt = 0
        while True:
            try:
                svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                if logger:
                    logger.info(f"[{idx}/{len(targets)}] Recalc OK: {t['title']}")
                break
            except Exception:
                attempt += 1
                if attempt > retries:
                    if logger:
                        logger.warn(f"FAILED recalc for {t['title']}")
                    break
                time.sleep(1 + random.random())


def get_value(svc, spreadsheet_id: str, sheet_title: str, named_range: str) -> str:
    try:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=named_range
        ).execute().get("values", [])
    except Exception:
        vals = []
    if not vals:
        try:
            vals = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_title}'!A1:A"
            ).execute().get("values", [])
        except Exception:
            vals = []
    return vals[0][0] if vals and vals[0] else "UNKNOWN"


def first_gid(svc, spreadsheet_id: str) -> int:
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return meta["sheets"][0]["properties"]["sheetId"]

# --- Additional helpers for row inspection/edits ---

def get_first_sheet_meta(svc, spreadsheet_id: str):
    """Return (first_sheet_title, first_sheet_id)."""
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    first = meta["sheets"][0]["properties"]
    return first["title"], first["sheetId"]

def get_values_2d(svc, spreadsheet_id: str, sheet_title: str, a1_range: str = "A:Z"):
    """Fetch a 2D values array from a sheet title + A1 range."""
    rng = f"'{sheet_title}'!{a1_range}"
    res = svc.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=rng).execute()
    return res.get("values", [])

def delete_rows_range(svc, spreadsheet_id: str, sheet_id: int, start_row_index: int, end_row_index: int):
    """Delete [start_row_index, end_row_index) (0‑based; end exclusive)."""
    if end_row_index <= start_row_index:
        return
    body = {"requests": [{
        "deleteDimension": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": start_row_index,
                "endIndex": end_row_index,
            }
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

def delete_row_indices(svc, spreadsheet_id: str, sheet_id: int, row_indices_desc: list[int]):
    """Delete multiple absolute row indices (0‑based) in descending order."""
    for r in sorted(row_indices_desc, reverse=True):
        delete_rows_range(svc, spreadsheet_id, sheet_id, r, r+1)

def add_blank_sheet(svc, spreadsheet_id: str, title: str, rows: int = 1000, cols: int = 26):
    """Create a blank sheet with a given title."""
    body = {"requests": [{
        "addSheet": {"properties": {"title": title, "gridProperties": {"rowCount": rows, "columnCount": cols}}}
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

def add_or_replace_sheet(svc, spreadsheet_id: str, title: str, rows: int = 2000, cols: int = 50):
    """
    Remove any existing sheet with 'title' and add a blank one.
    """
    try:
        delete_sheet(svc, spreadsheet_id, title)
    except Exception:
        # if not present, ignore
        pass
    add_blank_sheet(svc, spreadsheet_id, title, rows, cols)

def put_values_2d(svc, spreadsheet_id: str, sheet_title: str, values: list[list]):
    """
    Write a 2D array to 'A1' of 'sheet_title' in a single update.
    """
    rng = f"'{sheet_title}'!A1"
    svc.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": values}
    ).execute()

def _force_column_as_text(header: list[str], rows: list[list], header_name: str) -> list[list]:
    """
    For the column matching header_name, coerce every non-blank value to a string
    prefixed with a single apostrophe, so Google Sheets stores it as text.
    """
    idx = None
    for i, h in enumerate(header):
        if str(h).strip().lower() == header_name.strip().lower():
            idx = i
            break
    if idx is None:
        return rows  # header not found; nothing to do

    out = []
    for r in rows:
        r2 = list(r)
        if idx < len(r2) and r2[idx] not in (None, ""):
            # ensure string and prefix with apostrophe
            r2[idx] = "'" + str(r2[idx])
        out.append(r2)
    return out

```

---
### file: last_run.log

```
[2026-03-03 23:16:33] INFO: Authorizing with Google APIs…
[2026-03-03 23:16:33] INFO: Google services ready
[2026-03-03 23:16:33] INFO: Finding latest incoming spreadsheet…
[2026-03-03 23:16:34] INFO: Latest incoming: Testing Sales Report - Week 5 (1kNjmEbljdUIqJUjwh8e2-plfLWRZSGHxMHxL50-2ce4)
[2026-03-03 23:16:34] INFO: Preparing calculations workbook…
[2026-03-03 23:16:42] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 23:16:50] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 23:16:50] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 23:17:17] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 23:17:22] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 23:18:15] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 23:18:17] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-11-18-PM
[2026-03-03 23:18:17] INFO: Exporting Manager Report (PDF)…
[2026-03-03 23:18:21] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/1UILYk_poamIwe6JrFQUGZIz9az9iBsdU/view?usp=drivesdk
[2026-03-03 23:18:21] INFO: Exporting Master Order (CSV)…
[2026-03-03 23:18:34] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1bOcle5eMseFc43TZMnXLqgHR7_DcPs13ntb6PSBians/edit?usp=drivesdk
[2026-03-03 23:18:42] INFO: Emailed COFFEE
[2026-03-03 23:18:43] INFO: Manager email sent
[2026-03-03 23:18:43] INFO: Separate full order email disabled
[2026-03-03 23:18:43] INFO: Run completed in 00:02:09

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: requirements.txt

```text
-r documentation/requirements.txt
```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```

---
### file: ui_streamlit.py

```python
import os
import time
import threading
import json
import base64
import hashlib
import secrets
import re

import streamlit as st
from streamlit.components.v1 import html
from streamlit.components.v1 import html as _html_listener

from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials

from favtrip.google_client import load_valid_token, services, clear_token
from favtrip.config_store import save_config_to_drive
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.drive_utils import upload_to_drive


# =========================
# Constants & Simple Helpers
# =========================

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _split_emails(csv_str: str):
    return [e.strip() for e in (csv_str or "").split(",") if e.strip()]


def _parse_emails(csv_str: str):
    return _split_emails(csv_str)


def _invalid_emails(csv_str: str):
    return [e for e in _parse_emails(csv_str) if not EMAIL_RE.match(e)]


def _analyze_rk_rows(rows):
    """
    Validate the 'Per-Report-Key Recipients' editor rows.
    Returns (issues: List[str], preview_lines: List[str], rk_map: Dict[str, List[str]])
    """
    issues, preview, rk_map = [], [], {}
    seen, dupes = set(), set()

    for idx, r in enumerate(rows or [], start=1):
        raw_key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
        emails_csv = r.get("Emails (comma)") or ""
        if not raw_key and not emails_csv:
            # allow a blank template row
            continue

        # uppercase flag
        if raw_key != raw_key.upper():
            issues.append(f"Row {idx}: key '{raw_key}' is not ALL CAPS.")

        # duplicate detection
        if raw_key:
            if raw_key in seen:
                dupes.add(raw_key)
            else:
                seen.add(raw_key)

        # email validation
        bads = _invalid_emails(emails_csv)
        if bads:
            issues.append(f"Row {idx}: invalid emails → {', '.join(bads)}")

        # mapping + preview
        if raw_key:
            emails = _parse_emails(emails_csv)
            if emails:
                rk_map[raw_key] = emails
            preview.append(f"{raw_key} → {', '.join(emails) if emails else emails_csv}")

    if dupes:
        issues.append(f"Duplicate keys detected: {', '.join(sorted(dupes))}")
    return issues, preview, rk_map


def _b64url(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")


def _pkce_pair() -> tuple[str, str]:
    """
    Generate a high-entropy PKCE code_verifier and its S256 code_challenge.
    RFC 7636 requires 43–128 chars; this approach yields a URL-safe value.
    """
    verifier = _b64url(secrets.token_bytes(64))        # ~86 chars, URL-safe, no padding
    challenge = _b64url(hashlib.sha256(verifier.encode("ascii")).digest())
    return verifier, challenge


def _redirect_base() -> str:
    """
    Always return a non-empty redirect base that exactly matches your OAuth client's
    Authorized redirect URI. Prefer Secrets; normalize to one trailing slash.
    """
    base = (st.secrets.get("APP_BASE_URL", "") or "").strip()
    if not base:
        # Fallback to request (often available), still normalized
        try:
            base = (st.request.url_root or "").strip()
        except Exception:
            base = ""
    if not base:
        st.error("OAuth redirect base is not set. Define APP_BASE_URL in Secrets.")
        st.stop()
    return base.rstrip("/") + "/"


def _parse_state(state_b64: str) -> dict:
    # Add padding back for base64 decoding if needed
    padding = "=" * ((4 - len(state_b64) % 4) % 4)
    raw = base64.urlsafe_b64decode(state_b64 + padding)
    return json.loads(raw.decode("utf-8"))


def _infer_media_mime(name: str) -> str:
    n = (name or "").lower()
    if n.endswith(".csv"):
        return "text/csv"
    if n.endswith(".xlsx"):
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return "application/octet-stream"


def _get_drive_service_or_raise(cfg):
    creds = load_valid_token(cfg.SCOPES)
    if not creds:
        raise RuntimeError("Google authorization required. Please sign in first.")
    _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    return drive


def _rerun():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()


# =========================
# OAuth (Web / PKCE)
# =========================

def start_web_oauth(scopes):
    """
    Build an authorization URL that:
      - uses a stable redirect_uri (from Secrets)
      - uses explicit PKCE (S256)
      - embeds the code_verifier inside the state (base64url(JSON))
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    redirect = _redirect_base()

    # Explicit PKCE (stateless across redirect)
    code_verifier, code_challenge = _pkce_pair()

    # CSRF token + verifier encoded into state that Google will return unchanged.
    state_obj = {
        "csrf": _b64url(secrets.token_bytes(16)),
        "v": code_verifier,
        "r": redirect,
    }
    state_b64 = _b64url(json.dumps(state_obj).encode("utf-8"))

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
        state=state_b64,
        code_challenge=code_challenge,
        code_challenge_method="S256",
    )

    # Keep only minimal context; state carries the verifier.
    st.session_state["_oauth_redirect"] = redirect
    return auth_url


def finish_web_oauth(code: str, state_b64: str, scopes):
    """
    Recreate a Flow with the same redirect_uri and exchange code + code_verifier for tokens.
    (No UI side effects here.)
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    state_obj = _parse_state(state_b64)
    code_verifier = state_obj.get("v")
    redirect = state_obj.get("r") or st.session_state.get("_oauth_redirect") or _redirect_base()

    if not code_verifier:
        st.error("OAuth state did not include a PKCE code_verifier.")
        st.stop()

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    flow.fetch_token(code=code, code_verifier=code_verifier)

    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds

# --- OAuth Redirect Handler (Web/PKCE only) ---


# =========================
# UI Sections
# =========================

def render_run_form(cfg):
    err = None
    # --- STATE INIT ---
    if "incoming_uploader_version" not in st.session_state:
        st.session_state.incoming_uploader_version = 0
    if "incoming_selected_name" not in st.session_state:
        st.session_state.incoming_selected_name = None
    if "incoming_uploaded_ok" not in st.session_state:
        st.session_state.incoming_uploaded_ok = False

    uploader_key = f"incoming_upload_v{st.session_state.incoming_uploader_version}"

    # =========================
    # UPLOAD CARD (outside form)
    # =========================
    
    with st.container(border=True):
        st.subheader("Upload Live Items Report")
        st.caption("Please upload a live items report from Modisoft. The report must either be 1 or 2 full weeks of data.")
        # Use the SAME column grid as the run form header: [4, 1, 1]
        up_col, _, upbtn_col = st.columns([4, 1, 1])

        with up_col:
            incoming_file = st.file_uploader(
                "Upload Current Week Sales Report",
                type=["xlsx", "csv"],
                key=uploader_key,
                help="Please upload the current 'Live Items Report' from Modisoft as an XLSX or CSV file.",
                label_visibility="collapsed",
                accept_multiple_files=False,
            )

        # Track selection to manage gating (must click Upload Now successfully before Run)
        file_selected = incoming_file is not None
        if file_selected and st.session_state.get("incoming_selected_name") != incoming_file.name:
            st.session_state.incoming_selected_name = incoming_file.name
            st.session_state.incoming_uploaded_ok = False

        with upbtn_col:
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)
            
            upload_clicked = st.button(
                "⬆️ Upload Now",
                use_container_width=True,
                disabled=(not file_selected) or st.session_state.get("freeze_ui", False),
                type="primary",
                key="upload_submit",
            )

            st.markdown('</div>', unsafe_allow_html=True)

        # --- Handle the upload action immediately ---
        if upload_clicked:
            if not cfg.INCOMING_FOLDER_ID:
                st.error("Incoming Folder ID is empty. Set it under **Advanced → Incoming Folder ID**.")
            elif incoming_file is None:
                st.warning("Choose a .xlsx or .csv file first.")
            else:
                try:
                    drive = _get_drive_service_or_raise(cfg)
                    media_mime = _infer_media_mime(incoming_file.name)
                    base_name = os.path.splitext(incoming_file.name)[0]
                    nice_name = f"{base_name} (uploaded via UI)"
                    created = upload_to_drive(
                        drive,
                        data=incoming_file.getvalue(),
                        name=nice_name,
                        mime=media_mime,
                        folder_id=cfg.INCOMING_FOLDER_ID,
                        to_sheet=True,
                    )
                    link = created.get("webViewLink", "")

                    st.session_state.incoming_uploaded_ok = True
                    st.session_state.incoming_uploader_version += 1
                    st.session_state.incoming_selected_name = None
           

                    st.success("✅ Uploaded to Incoming as a Google Sheet.")
                    if link:
                        st.link_button("Open uploaded Sheet", link, use_container_width=True)
                    st.caption("This will be treated as the latest incoming report on the next run.")
                    _rerun()
                except Exception as e:
                    st.error(f"Upload failed: {e}")


    # =========================
    # RUN FORM CARD
    # =========================
    # When upload is OK, make the Run button green by adding .ft-run-green to the page segment
    
    run_form_wrapper_classes = "ft-card ft-row"
    had_ok = st.session_state.get("incoming_uploaded_ok", False)
    if had_ok:
        run_form_wrapper_classes += " ft-run-green"

    
    # OPEN the wrapper with real HTML (no entities)
    st.markdown(f'<div class="{run_form_wrapper_classes}">', unsafe_allow_html=True)



    with st.form("run_form"):
        # Header row uses the same columns to align the Run button with Upload button above
        tl, _, col_run = st.columns([4, 1, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")

        # --- Unified gating logic ---
        # A) If a file is currently selected but not uploaded -> disable Run
        # B) If no file selected and we have a prior successful upload -> enable Run
        # C) Otherwise (no prior upload or ambiguous state) -> disable Run
        file_selected = st.session_state.get("incoming_selected_name") is not None
        had_ok = st.session_state.get("incoming_uploaded_ok", False)

        run_enabled = (had_ok and not file_selected)
        run_disabled = not run_enabled

        with col_run:
            # Right-align and full-width, matching Upload Now
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)
            submitted = st.form_submit_button(
                "▶️ Run Pipeline",
                use_container_width=True,
                disabled=run_disabled or st.session_state.get("freeze_ui", False),
                type="primary",
                key="run_submit"
            )
            st.markdown('</div>', unsafe_allow_html=True)

        # ----- Main options -----

        # Recipients
        st.markdown("##### Recipients")
        col1, col2 = st.columns([1, 1])
        with col1:
            to = st.text_input(
                "To (comma)", value=",".join(cfg.TO_RECIPIENTS or []),
                help="Fallback recipients for Manager & Order emails."
            )
        with col2:
            cc = st.text_input(
                "CC (comma)", value=",".join(cfg.CC_RECIPIENTS or []),
                help="Optional CC added to all emails."
            )

        # Report Keys
        st.markdown("##### Report Keys")
        colk1, colk2 = st.columns([0.45, 1.55])
        with colk1:
            use_all = st.toggle(
                "Use all keys from CSV",
                value=cfg.USE_ALL_REPORT_KEYS,
                help="ON: process every key found. OFF: only the keys you list."
            )
        with colk2:
            report_keys = st.text_input(
                "Keys to run (comma)",
                value=",".join(cfg.REPORT_KEY_RUN_LIST or []),
                help="Used when 'Use all keys' is OFF. Example: COFFEE,GROCERY"
            )

        # General Behavior
        st.markdown("##### General Behavior")
        cole1, cole2, cole3, cole4 = st.columns([1, 1, 1, 1])
        with cole1:
            include_full = st.toggle(
                "Attach FULL order in each email",
                value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL
            )
        with cole2:
            send_full = st.toggle(
                "Send separate FULL order email",
                value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL
            )
        with cole3:
            email_mgr = st.toggle(
                "Email Manager Report",
                value=getattr(cfg, "EMAIL_MANAGER_REPORT", True),
                help="When ON, the Manager Report email is sent. When OFF, it is skipped."
            )
        with cole4:
            use_rollover = st.toggle(
                    'Use auto-rollover for single week uploads',
                    value=cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK,
                    help='If this is on, when only 1 week is uploaded, the most recent previously uploaded data will become the "Last Week" data; If this is off then the "Last Week" data will be left blank'
                )

        # Per‑Report‑Key Recipients
        with st.expander("Per‑Report‑Key Recipients (optional)", expanded=False):
            st.caption("Map **REPORT KEY (ALL CAPS)** → **Emails (comma)**.")
            rows = []
            if cfg.REPORT_KEY_RECIPIENTS:
                for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({"REPORT KEY (ALL CAPS)": k, "Emails (comma)": ",".join(v or [])})
            else:
                rows = [{"REPORT KEY (ALL CAPS)": "", "Emails (comma)": ""}]
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                use_container_width=True,
                key="rk_editor",
            )

            rk_issues, rk_preview, rk_map_preview = _analyze_rk_rows(edited_rows)
            if rk_preview:
                with st.expander("Row template preview"):
                    st.code("\n".join(rk_preview), language="text")
            if rk_issues:
                st.warning("Per‑report‑key recipient issues:\n\n- " + "\n- ".join(rk_issues))

        # Advanced
        with st.expander("Advanced (IDs, GIDs, Timezone, Redirect Port)", expanded=False):
            st.markdown("##### Google Drive / Sheets IDs")
            ga1, ga2 = st.columns([1, 1])
            with ga1:
                calc_id = st.text_input("Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID)
                mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID)
                user_folder = st.text_input("User Calculations Folder ID", value=cfg.USER_FOLDER_ID)
            with ga2:
                incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID)
                order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID)

            st.markdown("##### GIDs & Named Ranges")
            gb1, gb2 = st.columns([1, 1])
            with gb1:
                gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF))
                loc_sheet = st.text_input("Location Sheet Title", value=cfg.LOCATION_SHEET_TITLE)
            with gb2:
                gid_order = st.text_input("Order CSV gid", value=str(cfg.GID_ORDER_CSV))
                loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE)
                update_range = st.text_input("Update Timestamp Range", value=cfg.TEMPLATE_UPDATE_RANGE)

            st.markdown("##### Time & OAuth")
            gc1, gc2 = st.columns([1, 1])
            with gc1:
                tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ)
                tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT)
                raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                redirect_port = st.number_input(
                    "Redirect Port (0 = auto)",
                    min_value=0, max_value=65535,
                    value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                    help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535."
                )

                user_ttl = st.number_input(
                    "User Calculations Time-To-Life (days)",
                    min_value=0,
                    max_value=3650,
                    value=int(cfg.USER_TIME_TO_LIFE),
                    help="Delete old unused user calculations files older than this many days."
                    )

            with gc2:                                
                # --- New advanced intake settings ---
                _days = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Any"]
                start_dow = st.selectbox(
                    "Start day of week", _days, index=_days.index(cfg.START_DAY_OF_WEEK),
                    help="The day of week that the uploaded data should start at, any other day will raise an error"
                )
                end_dow = st.selectbox(
                    "End day of week", _days, index=_days.index(cfg.END_DAY_OF_WEEK),
                    help="The day of week that the uploaded data should end at, any other day will raise an error"
                )
                failed_input_ttl = st.number_input(
                    "Failed Input Time-To-Life (days)",
                    min_value=0,
                    max_value=3650,
                    value=int(cfg.FAILED_INPUT_TIME_TO_LIFE),
                    help="Delete old unused incoming files older than this many days."
                    )
                output_ttl = st.number_input(
                    "Output Time-To-Life (days)",
                    min_value=0,
                    max_value=3650,
                    value=int(cfg.OUTPUT_TIME_TO_LIFE),
                    help="Delete Manager/Order report files older than this many days after a successful run."
                    )


        save_drive_defaults = st.checkbox("Update defaults", value=False)

        # ----- Submission handling -----
        def _split_emails(csv_str: str):
            return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

        if submitted:
            # Apply per-run config
            cfg.TO_RECIPIENTS = _split_emails(to)
            cfg.CC_RECIPIENTS = _split_emails(cc)
            cfg.USE_ALL_REPORT_KEYS = use_all
            cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

            cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
            cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full
            cfg.EMAIL_MANAGER_REPORT = bool(email_mgr)

            cfg.CALC_SPREADSHEET_ID = calc_id
            cfg.INCOMING_FOLDER_ID = incoming_id
            cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
            cfg.ORDER_REPORT_FOLDER_ID = order_folder
            cfg.USER_FOLDER_ID = user_folder
            cfg.REDIRECT_PORT = int(redirect_port)

            cfg.GID_MANAGER_PDF = gid_mgr
            cfg.GID_ORDER_CSV = gid_order
            cfg.LOCATION_SHEET_TITLE = loc_sheet
            cfg.LOCATION_NAMED_RANGE = loc_range
            cfg.TEMPLATE_UPDATE_RANGE = update_range
            cfg.TIMESTAMP_TZ = tz
            cfg.TIMESTAMP_FMT = tfmt
            
            cfg.OUTPUT_TIME_TO_LIFE = int(output_ttl)
            cfg.FAILED_INPUT_TIME_TO_LIFE = int(failed_input_ttl)
            cfg.USER_TIME_TO_LIFE = int(user_ttl)

            cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK = bool(use_rollover)
            cfg.START_DAY_OF_WEEK = start_dow
            cfg.END_DAY_OF_WEEK = end_dow


            # Per-key recipients from editor
            
            rk_map = {}
            for r in edited_rows:
                key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
                emails_csv = r.get("Emails (comma)") or ""
                emails = _parse_emails(emails_csv)
                if key and emails:
                    rk_map[key] = emails
            cfg.REPORT_KEY_RECIPIENTS = rk_map

            # --- ADD: warnings before kicking off the run ---
            # 1) Warn if use_all is OFF and no explicit keys provided
            requested_keys = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]
            if not use_all and not requested_keys:
                st.warning(
                    "You left **Use all Report Keys** OFF but provided **no keys** to run. "
                    "No per‑key outputs will be generated unless you add keys.",
                    icon="⚠️"
                )

            # 2) Warn if no recipients anywhere (TO, DEFAULT_ORDER, or per‑key map)
            any_to = bool(_parse_emails(to) or (cfg.TO_RECIPIENTS or []))
            any_default = bool(cfg.DEFAULT_ORDER_RECIPIENTS or [])
            any_per_key = bool(cfg.REPORT_KEY_RECIPIENTS)
            if not (any_to or any_default or any_per_key):
                st.warning(
                    "No recipients are defined: **TO**, **DEFAULT_ORDER_RECIPIENTS**, and **Per‑Report‑Key** are all empty. "
                    "Emails will not be sent.",
                    icon="⚠️"
                )

            # 3) Surface per-key table issues detected earlier
            if rk_issues:
                st.warning(
                    "Per‑report‑key recipient issues detected above. These may prevent emails from sending correctly:\n\n- "
                    + "\n- ".join(rk_issues),
                    icon="⚠️"
                )
            # --- END ADD ---


            # --- Save edited defaults to Drive JSON (optional) ---
            if save_drive_defaults:
                try:
                    # Ensure we have a user token first
                    creds = load_valid_token(cfg.SCOPES)
                    if not creds:
                        st.error("Not authenticated. Please complete Google sign‑in first (top of page).")
                    else:
                        # Drive service
                        _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                        # What we persist (the fields you asked to move out of Secrets)
                        drive_defaults = {
                            "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
                            "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
                            "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
                            "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,
                            "USER_FOLDER_ID": cfg.USER_FOLDER_ID,

                            "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
                            "GID_ORDER_CSV": cfg.GID_ORDER_CSV,

                            "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
                            "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,
                            "TEMPLATE_UPDATE_RANGE": cfg.TEMPLATE_UPDATE_RANGE,

                            "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
                            "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,

                            "OUTPUT_TIME_TO_LIFE": cfg.OUTPUT_TIME_TO_LIFE,
                            "FAILED_INPUT_TIME_TO_LIFE": cfg.FAILED_INPUT_TIME_TO_LIFE,
                            "USER_TIME_TO_LIFE": cfg.USER_TIME_TO_LIFE,

                            "TO_RECIPIENTS": cfg.TO_RECIPIENTS,   # lists are fine; JSON keeps types
                            "CC_RECIPIENTS": cfg.CC_RECIPIENTS,

                            "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
                            "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,

                            "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,

                            "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,

                            "INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL": cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL,
                            "SEND_SEPARATE_FULL_ORDER_EMAIL": cfg.SEND_SEPARATE_FULL_ORDER_EMAIL,
                            "EMAIL_MANAGER_REPORT": cfg.EMAIL_MANAGER_REPORT,

                            "USE_AUTO_ROLLOVER_IF_ONE_WEEK" : cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK,
                            "START_DAY_OF_WEEK" : cfg.START_DAY_OF_WEEK,
                            "END_DAY_OF_WEEK" : cfg.END_DAY_OF_WEEK
                        }

                        # If you have CONFIG_FILE_ID in Secrets, we update that exact file.
                        # Otherwise we'll upsert a file named 'favtrip_config.json' and return its id.
                        CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
                        new_id = save_config_to_drive(
                            drive,
                            drive_defaults,
                            file_id=CONFIG_FILE_ID or None,   # update/pin if set, else upsert by name
                            # parent_folder_id=None,          # optional: set a folder id to create under
                        )

                        st.success(f"Saved defaults to Drive config (file id: {new_id}).")
                        if not CONFIG_FILE_ID:
                            st.info(
                                "Tip: add this ID to Streamlit Secrets as `CONFIG_FILE_ID` to pin the same file for all runs:\n"
                                f"`{new_id}`"
                            )
                except Exception as e:
                    st.error(f"Failed to save defaults to Drive: {e}")

            # If user checked "Force Google re-auth for this run", kick them into auth gating first.
            if cfg.FORCE_REAUTH:
                clear_token()
                try:
                    flow, url = start_web_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                    st.session_state.oauth_flow = flow
                    st.session_state.oauth_url = url
                    st.session_state.auth_required = True
                    st.info("Re-auth required for this run. Open the URL shown in the Authentication panel.")
                    _rerun()
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")
            else:
                # --- Live run with timer + last log (no full log after completion) ---
                # Write all logs to last_run.log; overwrite on each run
                logger = StatusLogger(print_to_console=True, file_path="last_run.log", overwrite=True)
                result_holder = {"value": None, "error": None}

                def _runner():
                    try:
                        result_holder["value"] = run_pipeline(cfg, logger=logger)
                    # Catch BaseException so SystemExit / KeyboardInterrupt are captured too
                    except BaseException as e:
                        result_holder["error"] = e

                t0 = time.perf_counter()
                th = threading.Thread(target=_runner, daemon=True)
                th.start()

                with st.status("Running pipeline…", expanded=True) as status:
                    timer_ph = st.empty()
                    lastlog_ph = st.empty()
                    while th.is_alive():
                        elapsed = int(time.perf_counter() - t0)
                        timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                        lastlog_ph.markdown(f"**Last:** {logger.last_line()}")
                        time.sleep(0.5)

                    th.join()
                    elapsed = int(time.perf_counter() - t0)
                    timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                    lastlog_ph.markdown(f"**Last:** {logger.last_line()}")

                    if result_holder["error"]:
                        err = result_holder["error"]
                        err_text = str(err)

                        # Always lock out Run after a failure until a new upload occurs
                        st.session_state.incoming_uploaded_ok = False

                        # --- Special case: "Please only upload 1 or 2 full weeks of data"
                        weeks_err = "Please only upload 1 or 2 full weeks of data" in err_text
                        if weeks_err:
                            # Route to the simple locked UI (error + Retry)
                            st.session_state["file_error"] = err_text
                            st.session_state["incoming_locked"] = True
                            status.update(label="❌ Invalid file (1–2 full weeks required)", state="error")
                            _rerun()
                            st.stop()

                        # --- Generic error path: show details here and freeze the UI
                        st.error("Run failed.")
                        try:
                            with st.expander("Error details", expanded=True):
                                st.exception(err)
                        except Exception:
                            # Fallback if exception rendering fails
                            with st.expander("Error details", expanded=True):
                                st.write(err_text)

                        # Freeze the UI until the user clicks Retry (after the form)
                        st.session_state["freeze_ui"] = True
                        status.update(label="❌ Failed", state="error")
                        

                    else:
                        result = result_holder["value"]
                        if result is None:
                            st.error("Run finished without returning a result. Check logs and inputs (IDs, Drive access).")
                            status.update(label="⚠️ No result", state="error")
                        else:
                            st.write("### Outputs")
                            col1, col2, col3 = st.columns(3)
                            col1.metric("Location", result.location)
                            col2.metric("Timestamp", result.timestamp)
                            mm = result.elapsed_seconds
                            col3.metric("Elapsed", f"{mm//3600:02d}:{(mm%3600)//60:02d}:{mm%60:02d}")

                            if getattr(result, "manager_pdf_link", None):
                                st.success(f"Manager PDF: {result.manager_pdf_link}")
                            if getattr(result, "full_order_link", None):
                                st.success(f"Full Order Sheet: {result.full_order_link}")

                                
                            if os.path.exists("last_run.log"):
                                with open("last_run.log", "rb") as f:
                                    st.session_state["last_run_log"] = f.read()
                                    st.session_state["last_run_timestamp"] = result.timestamp


                            st.session_state.incoming_uploaded_ok = False
                            st.session_state.incoming_uploader_version += 1
                            st.session_state.incoming_selected_name = None

                            status.update(label="✅ Completed", state="complete")
                            time.sleep(10)
                            _rerun()


    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.get("offer_log_download", False) and "last_run_log" in st.session_state:
        st.download_button(
            "⬇️ Download full log (last_run.log)",
            st.session_state["last_run_log"],
            file_name=f"last_run_{st.session_state['last_run_timestamp']}.log",
            mime="text/plain",
            use_container_width=True
            )

    # Post-form controls for generic errors
    # If the UI is frozen due to a non-file error, show a single Retry button here.
    if st.session_state.get("freeze_ui", False) and not st.session_state.get("incoming_locked", False):
        retry_cols = st.columns([1, 3, 1])
        with retry_cols[1]:
            if st.button("Retry", type="secondary", use_container_width=True):
                # Unfreeze, but force a new file upload before the next run
                st.session_state["freeze_ui"] = False
                st.session_state["incoming_uploaded_ok"] = False
                st.session_state["incoming_selected_name"] = None
                st.session_state["incoming_uploader_version"] += 1  # resets uploader widget
                _rerun()





# =========================
# App Entrypoint
# =========================

#st.title("🧾 FavTrip Reporting Pipeline")


st.set_page_config(
    page_title="FT Reporting",
    page_icon="🧾",          # emoji or path/URL to an image
    layout="wide",           # "centered" or "wide"
    initial_sidebar_state="expanded",  # "auto", "expanded", "collapsed"
    menu_items={
        "Get Help": "mailto:ryan-morrow@uiowa.edu",
        "Report a bug": "https://github.com/ryan-j-morrow/favtrip_reporting/issues",
        "About": "FavTrip Reporting Pipeline",
    },
)


st.markdown("""
<style>
/* Make only buttons inside a .ft-run-green scope look green */
.ft-run-green .stButton > button {
    background-color: #22c55e !important;  /* green */
    border-color: #16a34a !important;
    color: #ffffff !important;
}
/* Optional: dim disabled state a bit */
.ft-run-green .stButton > button:disabled {
    opacity: .6 !important;
    cursor: not-allowed !important;
}
/* Utility: right-align containers where used */
.ft-right-btn { display:flex; justify-content:flex-end; }
</style>
""", unsafe_allow_html=True)


cfg = Config.load()

# --- Finish OAuth inline when redirect comes back (this is in the NEW TAB) ---
params = st.query_params
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        # Token is saved locally in this new tab's app process
        st.success("✅ Google authentication complete.")

        # Remove code/state from URL
        st.query_params.clear()

        # No messaging back to opener and NO window.close().
        # This tab becomes the main app; just rerun to flip UI.
        st.toast("Signed in. Loading the app…")
        st.rerun()
    except Exception as e:
        st.error(f"OAuth error: {e}")

# If a valid token now exists (in this tab), reload cfg (Drive overlays, etc.)
st.session_state.auth_required = (load_valid_token(cfg.SCOPES) is None)
if not st.session_state.auth_required:
    cfg = Config.load()

# Session state defaults
if "incoming_selected_name" not in st.session_state:
    st.session_state.incoming_selected_name = None
if "incoming_uploaded_ok" not in st.session_state:
    st.session_state.incoming_uploaded_ok = False
if "offer_log_download" not in st.session_state:
    st.session_state.offer_log_download = False
if "offer_log_download" not in st.session_state:
    st.session_state.offer_log_download = False


# Sidebar (always visible)
with st.sidebar:
    st.header("Utilities")

    if st.button("Google Sign Out", type="secondary", use_container_width=True):
        clear_token()
        for key in ["auth_required", "oauth_flow", "oauth_url", "auth_checked"]:
            if key in st.session_state:
                del st.session_state[key]
        try:
            st.rerun()
        except AttributeError:
            st.experimental_rerun()
    
    st.checkbox(
        "Offer full log download after completion",
        key="offer_log_download",
        help="If enabled, a 'Download last_run.log' button appears when a run finishes.")

# Auth gate
if st.session_state.auth_required:
    # ----------------------------
    # Authentication panel (shown only if auth required)
    # ----------------------------
    if st.session_state.auth_required:
        with st.expander("Google Authentication", expanded=True):
            st.caption(
                "Authentication is required before running. "
                "Click **Sign in with Google** to open the consent screen (it will open in a new tab)."
            )

            sign_in_ph = st.empty()
            clicked = sign_in_ph.button("Sign in with Google", type="primary", use_container_width=True)

            if clicked:
                try:
                    auth_url = start_web_oauth(cfg.SCOPES)
                    sign_in_ph.empty()

                    # Friendly message in this (original) tab
                    st.markdown(
                        """
                        <div style="
                            display:flex;align-items:center;justify-content:center;
                            height:55vh;text-align:center;
                            font-family: system-ui, Segoe UI, Roboto, Helvetica, Arial, sans-serif;">
                        <div>
                            <h2 style="margin-bottom:0.5rem;">You're being signed in…</h2>
                            <p style="font-size:1.05rem;opacity:.9;">
                            A new browser tab was opened for Google sign‑in.<br/>
                            <strong>After it completes, continue in that tab.</strong>
                            </p>
                        </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    # Optional: refresh this tab when user returns (not required)
                    html(
                        """
                        <script>
                        document.addEventListener("visibilitychange", function() {
                            if (!document.hidden) { location.reload(); }
                        });
                        </script>
                        """,
                        height=0,
                    )

                    # Open Google auth in a NEW tab (this will ultimately become the main app)
                    html(
                        f"""
                        <script>
                        window.open({json.dumps(auth_url)}, "_blank", "noopener");
                        </script>
                        """,
                        height=0,
                    )

                    st.stop()
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")

            with st.expander("Having trouble?", expanded=False):
                st.write(
                    "- The Google authorization page opens in a **new browser tab**.\n"
                    "- After completing consent, the **new tab** will load the app.\n"
                    "- If you renamed your Streamlit app or URL, ensure the Google OAuth "
                    "Authorized redirect URI matches exactly (including trailing slash)."
                )
            st.stop()

# Optional lock if invalid incoming file was detected
locked = st.session_state.get("incoming_locked", False)
if locked:

    err_msg = st.session_state.get("file_error", "An unknown error occurred.")
    st.error(err_msg)

    if st.button("Retry", type="secondary"):
        st.session_state["incoming_locked"] = False
        _rerun()
    st.stop()

else:
    # Ensure config reflects Drive-based overrides after auth
    cfg = Config.load()
    render_run_form(cfg)
```

---
### file: web_url_credentials.json

```json
{"web":{"client_id":"674901450584-i6ncesf794ecg6u7olebqkag8gavbuc3.apps.googleusercontent.com","project_id":"favtripdev","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["https://favtripreporting-dev1.streamlit.app/"]}}
```
