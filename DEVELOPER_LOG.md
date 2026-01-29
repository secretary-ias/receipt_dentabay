# Developer Log

## 2025-10-15 – Repository Audit & Documentation Pass
- Reviewed the full Tkinter application (`app/ui.py`) and mapped interactions between schedule, settlement, receipts, and settings tabs.
- Confirmed database layer (`app/database.py`) reads/writes directly against DoctorAssist MySQL, including receipt creation, partial payments, visit notes, and stock catalogue lookups.
- Documented the PDF generation workflow (`app/receipt.py`) and verified output naming aligns with the cashier's filing convention.
- Catalogued configuration touchpoints (`settings.json`, `ConfigManager`) and external dependencies (PyMySQL, ReportLab, Gmail SMTP).
- Produced `DOCUMENTATION.md` to capture architecture, runtime flow, and operational notes for new contributors.

## 2025-10-16 - Receipt Editing Enhancements
- Enabled manual unit price overrides across Settlement dialogs, visit note item composer, and receipt editing while keeping catalogue prices as defaults.
- Added receipt-tab editing for both line items and partial payments, including a new DAL helper `update_partial_payment_amount` to persist amount adjustments.
- Refreshed receipts UI state so totals, balances, and PDF exports reflect edits immediately without stale figures or NameError regressions.

## Observations & Follow-Up Ideas
- `app/ui.py` remains a single 170kB file; consider breaking it into tab-specific modules (`schedule.py`, `settlement.py`, etc.) or introducing a presenter/service layer to improve testability.
- Settlement queue assumes DoctorAssist status ID `3` means "Ready for payment"; make this configurable if multiple clinics use different workflows.
- Email support is Gmail-specific. To unblock alternative providers, abstract SMTP host/port and TLS options into settings.
- Data importer (`app/data_loader.py`) is powerful but unused in the UI—decide whether to expose a cache-refresh action or remove dead code.
- Introduce structured logging (e.g., `logging` module) for easier troubleshooting, especially around DB writes and external integrations (SMTP, WhatsApp, OS printing).
