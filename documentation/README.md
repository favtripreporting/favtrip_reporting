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

