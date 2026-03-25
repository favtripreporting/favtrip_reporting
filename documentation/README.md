# FavTrip Reporting Pipeline — Codebase Overview

## Purpose & High-Level Architecture

The FavTrip Reporting Pipeline is a **Google Workspace–integrated reporting system** designed to automate weekly store reporting based on Modisoft sales data. It accepts a “Live Items Report” input file, validates and processes one or two weeks of sales data, updates calculation workbooks, produces PDFs and Sheets, and distributes reports via Gmail.

The codebase is structured into **three clear layers**:

1. **User Interface Layer** – Handles user interaction, authentication, configuration, and orchestration.
2. **Core Functional Modules** – Implements all business logic and Google API interactions.
3. **Supporting Assets & Documentation** – Reference materials, examples, and developer tooling.

This README intentionally focuses only on the **UI entrypoint** and the **core functional modules**.  
For detailed behavior, contracts, and edge cases, refer to the **module-level docstrings** in each file.

---

## User Interface Layer

### `_user_interface_.py`

This file implements the **primary Streamlit web application** and is the main operational entrypoint for most users.

At a high level, the UI is responsible for:

- Handling **Google OAuth authentication** (PKCE-based, stateless across redirects).
- Accepting Modisoft report uploads (CSV/XLSX) and storing them in Google Drive.
- Exposing **runtime configuration controls**:
  - Email recipients
  - Report keys
  - Feature toggles
  - Advanced IDs, GIDs, and date validation rules
- Validating inputs early to prevent unsafe or invalid pipeline runs.
- Executing the backend pipeline and streaming **live status updates and timing**.
- Rendering outputs (Drive links, timestamps) and handling failure recovery.

Key architectural characteristics:

- The UI **does not contain business logic**.
- All processing is delegated to `core_functional_modules.pipeline.run_pipeline`.
- Configuration is assembled via a shared `Config` object and passed downstream.
- The UI can evolve independently of the pipeline without risking logic drift.

For details about OAuth flow, state management, upload gating, and UI locking behavior, see the module docstring in this file.

---

## Core Functional Modules

All core logic lives under `core_functional_modules/`. These modules are designed to be:

- UI-agnostic
- Composable and reusable
- Safe to invoke from both the Streamlit UI and the CLI

Only responsibilities are summarized here; implementation details live in docstrings.

---

### `config.py` — Central Configuration Model

- Defines the canonical `Config` dataclass used everywhere in the system.
- Loads configuration via a **layered merge**:
  1. Streamlit secrets (typed, preferred in cloud)
  2. Environment variables / `.env`
  3. Optional Google Drive–hosted JSON overrides
- Normalizes types (booleans, lists, dicts) for consistent behavior.
- Acts as the single source of truth for IDs, flags, recipients, and cleanup policies.

This module underpins all other core components.

---

### `config_store.py` — Drive-Backed Config Persistence

- Reads and writes JSON configuration files stored in Google Drive.
- Enables the UI’s “Update defaults” feature.
- Uses resilient, fail-open behavior so missing or malformed configs never break execution.

---

### `google_client.py` — OAuth & Google Service Bootstrapping

- Manages Google OAuth sign-in and token lifecycle (`token.json`).
- Supports both browser-assisted and manual (CLI-style) flows.
- Produces authenticated service clients for:
  - Google Drive
  - Google Sheets
  - Gmail

All Google API usage throughout the system flows through this module.

---

### `drive_utils.py` — Google Drive Operations

- Uploads files (raw or converted to Google Sheets).
- Finds the latest files in folders.
- Copies, renames, and trashes Drive resources.
- Performs **time-based cleanup** of old files.
- Handles Drive query escaping and RFC‑3339 timestamps.

---

### `sheets_utils.py` — Google Sheets Manipulation

- Implements all Sheet-level operations:
  - Copying, deleting, and renaming sheets
  - Writing 2D values
  - Refreshing formulas
  - Forcing text coercion on specific columns
- Provides chunked refresh utilities for large sheets.

This module isolates low-level Sheets API complexity from the pipeline.

---

### `gmail_utils.py` — Email Composition & Sending

- Builds MIME-compliant emails with:
  - PDF attachments
  - Optional full-order attachments
  - Plain-text and HTML bodies
- Sends messages through the Gmail API as the authenticated user.
- Assumes recipient selection has already been resolved upstream.

---

### `logger.py` — Run-Time Status Logging

- Lightweight status logger with:
  - In-memory event tracking
  - Optional file-backed logging (`last_run.log`)
- Designed for real-time UI updates and post-run log downloads.
- Avoids heavy logging frameworks for predictability and portability.

---

### `pipeline.py` — Orchestrated Processing Engine

This is the **core execution engine** of the system.

At a high level, it:

1. Authenticates and initializes Google services.
2. Locates the latest incoming report.
3. Validates date coverage and week boundaries.
4. Prepares or rolls forward calculation workbooks.
5. Refreshes reference sheets.
6. Generates PDFs and Sheets for:
   - Manager reports
   - Full orders
   - Per-report-key outputs
7. Routes emails using a prioritized fallback model.
8. Performs Drive cleanup and retention enforcement.

The pipeline is intentionally **UI-agnostic** and is safe to invoke from:
- The Streamlit UI
- The CLI
- Future automation or scheduling contexts

---

## Supporting Folders

### `documentation/`

This directory contains **developer-facing documentation and tooling**, including:

- Architectural and usage documentation
- Setup, packaging, and deployment notes
- Git workflow guidance
- Dependency lists
- Tooling used to generate the project snapshot markdown bundle

Nothing in this folder is required for runtime execution. It exists to support onboarding, maintenance, and knowledge transfer.

---

### `__dev_input_sales_files/`

This directory contains **example Modisoft sales reports** used for testing and validation.

The files intentionally cover multiple scenarios, including:

- One-week vs. two-week uploads
- Multiple stores
- Invalid or misaligned week boundaries
- Bad start or end dates

These files are **not consumed automatically** by the application and exist purely for manual testing and development.

---

## Files and Folders Not Covered Here

Any files or folders **not explicitly mentioned in this README** (for example: logs, generated outputs, cached tokens, per-user folders, or packaging artifacts) are:

- **Auto-created**
- **Auto-maintained**
- **Managed by the application at runtime**

⚠️ **Do not edit, delete, or manually modify these files unless you fully understand their lifecycle and side effects.**  
Incorrect changes may break authentication, corrupt Drive state, or cause data loss.

---

## How to Navigate the Codebase

- Start with `_user_interface_.py` to understand user flow and orchestration.
- Follow execution into `pipeline.run_pipeline`.
- Use module-level docstrings in `core_functional_modules/` for precise logic and guarantees.
- Treat `Config` as the shared contract that ties the system together.
