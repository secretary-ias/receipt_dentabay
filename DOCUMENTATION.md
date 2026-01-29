# Receipt Dentabay v5 – Developer Documentation

## Overview
- Desktop utility written in Python/Tkinter that gives the clinic temporary, read/write access to an existing DoctorAssist `clinicdb` MySQL instance.
- Primary goals: view the daily appointment schedule, settle visits by creating receipts directly in MySQL, review previously issued receipts, and generate/email/print PDF copies for patients.
- Uses ReportLab for PDF output, PyMySQL for database access, and a small theming layer to recreate the web UI styling inside Tkinter.

## Repository Layout
- `main.py` – entry point; loads `settings.json`, launches the Tkinter app.
- `app/` – application package.
  - `ui.py` – monolithic UI/controller layer (`ReceiptApp`) containing tab layouts, event handlers, and business logic.
  - `database.py` – thin data-access layer around DoctorAssist tables. Defines dataclasses for domain entities (patients, receipts, appointments, visit notes) and provides read/write helpers (`appointments_for_date`, `create_receipt`, `replace_receipt`, etc.).
  - `config.py` – settings dataclasses (`AppSettings`, `ClinicInfo`, `DatabaseInfo`, `EmailSettings`) plus `ConfigManager` for JSON persistence and path normalisation.
  - `receipt.py` – `ReceiptPDFGenerator`, payment dataclasses, and routines that lay out a receipt onto an A4 page.
  - `data_loader.py` – optional utility that can rebuild a lightweight SQLite cache from a MySQL connection or a DoctorAssist SQL dump (not wired into the UI yet).
  - `theme.py` – ttk style definitions and a `card()` helper for consistent layout.
  - `assets/` – default clinic logo(s) bundled with the app.
- `assets/` (root) – Windows icon and default logo used when settings do not point to a custom image.
- `clinicdb/` – extracted InnoDB table files (used only as a data source; the app connects to a running server).
- `data/clinic_cache.sqlite` – cached subset of clinic data if the importer is executed.
- `dist/`, `build/`, `*.spec` – PyInstaller artefacts for packaging a Windows executable.
- `settings.json` – runtime configuration edited through the Settings tab.

## Runtime Architecture
1. `main.py` resolves `settings.json`, instantiates `app.config.ConfigManager`, and creates `ReceiptApp`.
2. `ReceiptApp` initialises Tkinter styles, loads configuration, attempts to connect to MySQL via `ClinicDatabase`, and prompts the user for a username (optional).
3. A shared calendar selector controls the date context for three tabs (Schedule, Settlement, Receipts). Changing the date triggers fresh queries through `ClinicDatabase`.
4. Long-running or failure-prone operations (DB access, PDF rendering, SMTP) are wrapped in try/except blocks; user feedback is provided via Tk message boxes and the status bar.

## Tab Flow & Responsibilities
- **Schedule**
  - Lists appointments for the selected day (`ClinicDatabase.appointments_for_date`).
  - Colour-codes rows using the DoctorAssist status colour.
  - Shows a condensed timeline of visit notes (`medic_reports_for_patient`) and lets staff open the patient profile or launch the visit-note editor.
  - Visit notes are editable in a dedicated dialog that can pre-fill items from settlement data and pushes updates back through `update_medical_record`.

- **Settlement**
  - Pulls queue-ready appointments (`appointments_for_status(..., status_ids=(3,))`).
  - Displays patient contact info, latest visit note, and a treatment item list assembled either from the note's dental chart or existing receipts.
  - Users can add/edit/remove treatment line items sourced from `stock_items` (categories cached via `stock_categories()`), with unit prices defaulting from the catalogue but fully editable before saving.
  - `Mark as Paid` converts the in-memory list to `ReceiptDraftItem` records and calls either `create_receipt` or `replace_receipt`, persisting line items, totals, payment code, and linkage to the medic report.
  - Totals panel allows manual discount/rounding adjustments; payment methods are fetched via `payment_methods_list()`.

- **Receipts**
  - Lists receipts issued on the selected date (with optional patient filter) using `receipts_for_date`.
  - Displays totals, balance, and payment history (combining `get_receipt_items` and `receipt_payment_history`).
  - Supports editing existing receipts in place: highlight a line item to update quantity/price/remarks via the same dialog used in Settlement, or select a payment row to adjust the recorded amount.
  - Actions: generate PDF (`ReceiptPDFGenerator.generate` + `_rename_to_desired`), open via default OS viewer, email via Gmail SMTP (`smtplib.SMTP_SSL`), share through WhatsApp deep link, and print.
  - If instalment payments exist, the PDF includes progress tables using `PaymentProgress`.

- **Settings**
  - Edits clinic profile (name, address, contacts, logo), receipt output folder, and email template.
  - Switches data source between live MySQL and SQL backup. Paths are normalised and stored back into `settings.json`.
  - Offers quick actions to choose logo/output directory, browse for a backup file, save settings, and retry the database connection.

## Data Layer Highlights (`app/database.py`)
- Maintains a persistent PyMySQL connection; automatically downgrades charset if the server does not support `utf8mb4`.
- Dataclasses define typed payloads for appointments, patient profile, visit notes, receipts, and payment fragments.
- Provides fetch helpers for:
  - Appointments: `appointments_for_date`, `appointments_for_status`, individual appointment lookup.
  - Patient details: `get_patient_profile`, allergies, documents, deposits.
  - Receipts: `receipts_for_date`, `get_receipt_items`, `receipts_for_medic_report`, `receipt_items_for_patient_date`, `receipt_payment_history`.
  - Stock catalogue: `stock_categories`, `stock_items_by_category`, `stock_item_details`, dental notation metadata.
  - Visit notes: `medic_reports_for_patient`, `medic_report_for_appointment`, plus `update_medical_record`.
- Write helpers cover:
  - `create_receipt` / `replace_receipt` (receipt header + `receipt_items` lines).
  - `create_partial_payment` / `delete_partial_payment` / `update_partial_payment_amount`.
  - `record_deposit` and other supporting inserts.
- Includes cleanup utilities and defensive conversions (`_to_float`, `_to_datetime`) to cope with legacy column types.

## Receipt Generation Pipeline
1. UI collects the selected `ReceiptSummary` plus optional payment selection.
2. Fetches detailed line items (`get_receipt_items`) and payment history to construct `PaymentEntry` and `PaymentProgress` records.
3. `ReceiptPDFGenerator.generate` lays out header (clinic info + logo), patient block, items table, totals, and payment details. Output stored under `settings.receipt.output_directory`.
4. `_rename_to_desired` enforces file naming convention `receipt_{patient}_{receipt}_{YYYYMMDD}_pXX.pdf`, ensuring instalment copies are distinguishable.
5. Downstream actions (email/WhatsApp/print) consume the freshly generated PDF to avoid stale data.

## Configuration & Environment
- Runtime settings live in `settings.json`. The Settings tab keeps secret values (MySQL password, Gmail app password) editable without manual file edits.
- `ConfigManager` guarantees relative paths stay within the project directory (`receipts/` default) and materialises missing folders/logos on save.
- Ensure the following Python packages are installed in the active environment:
  - `pymysql`
  - `reportlab`
  - `tkinter` (bundled with CPython on Windows)
- Optional: run `pip install -r requirements.txt` if you maintain one; otherwise, install the two third-party libraries manually.

## Data Importer (Optional Workflow)
- `app/data_loader.ClinicDataImporter` can refresh a local SQLite cache (`data/clinic_cache.sqlite`) from either:
  - A DoctorAssist SQL backup (`backup_clinicdb.sql`), hashing the file to avoid unnecessary rebuilds.
  - Live MySQL credentials supplied via `ConfigManager.mysql_settings()`.
- The importer defines per-table plans (`TableImportPlan`) for receipts, receipt items, visit notes, patients, and related tables, normalising data before insert.
- The UI currently talks directly to MySQL; integrating the SQLite cache would require additional routing logic.

## Packaging
- PyInstaller spec files (`receipt_app.spec`, `ClinicReceiptPrinter.spec`) describe how to bundle the application into a single-file Windows executable, embedding assets and the `settings.json` template.
- `dist/` contains the latest packaged binaries; rebuild by activating the virtual environment and running `pyinstaller receipt_app.spec`.

## Operational Notes & Limitations
- The login dialog only validates usernames exist in the DoctorAssist `users` table; there is no password verification.
- Error handling is user-facing (message boxes/status bar) but logging is minimal—consider adding structured logs for support diagnostics.
- Settlement workflow assumes status ID `3` corresponds to patients waiting at the cashier; adjust if DoctorAssist changes enumerations.
- Email delivery is hard-coded for Gmail SMTP (`smtp.gmail.com:465`). Other providers would require Settings UI and DAL updates.
- `app/ui.py` is large (>170 KB); future refactors could split tabs into separate modules for readability and targeted testing.
