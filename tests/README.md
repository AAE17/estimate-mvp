# SO+ tests

Run: `npm install` then `npm test` (Node 20+). No LibreOffice, no real Supabase, no internet needed.

- `kachu-preview.test.js` – 34 kachu bill (MB) cases: screen preview (index.html functions) vs independent hand calculation.
- `kachu-mb-excel.test.js` – same 34 cases through `POST /api/mb`: Excel cells, hidden empty sections, no stale template rows.
- `security.test.js` – private APIs need login (401), expired account 403, logs admin-only.
- `fixtures/kachu-cases.json` – the cases. Add a new case here when a new bug is found.
- `helpers/` – fake Supabase + server starter (`SOFFICE_DISABLED=1`), preview runner, hand calculation.

Rule: every change must keep `npm test` green. GitHub Actions runs it on every push/PR.
