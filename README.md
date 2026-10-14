# Sunil Poultry — v2 (Google Sheet backend)

Files
- `index.html` – the whole site (set `API_URL` inside, see DEPLOY.md)
- `sw.js`, `manifest.webmanifest`, `icon*.png/svg` – installable phone app (PWA)
- `photo1.png` – put your farmer photo here (not included in the upload I received)
- `apps-script/Code.gs` – backend (Google Apps Script) → paste into your Sheet's Apps Script
- `apps-script/SeedData.gs` – sample data that `setupAndSeed()` writes into the Sheet
- `apps-script/appsscript.json` – project manifest (time zone, web-app access)
- `DEPLOY.md` – step-by-step setup
