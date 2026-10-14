# Sunil Poultry — Google Sheet setup & deployment

Everything the site saves goes into a Google Sheet. Photos go into a Google Drive folder
(the Sheet keeps their IDs). You do this setup **once** (about 10 minutes).

## 1. Create the Sheet and paste the code
1. Go to sheets.google.com → **Blank spreadsheet** → name it e.g. `Sunil Poultry Data`.
2. Menu **Extensions → Apps Script**.
3. Delete the sample code. Paste the contents of **`apps-script/Code.gs`** into `Code.gs`.
4. Click **＋ (Files) → Script**, name it `SeedData`, paste **`apps-script/SeedData.gs`**.
5. (Recommended) **Project Settings ⚙ → tick "Show appsscript.json"**, open it, and replace its
   contents with **`apps-script/appsscript.json`** (sets the India time zone and public web-app access).
6. At the top of `Code.gs`, change `OWNER_INITIAL_PASSWORD` (default `Sunil@2026`).
   You can change it later inside the app: **Account → Change owner password**.

## 2. Create the tabs and load the sample data
1. In the function drop-down choose **`setupAndSeed`** → **Run**.
2. Google asks for permission (Sheets + Drive). Choose your account →
   **Advanced → Go to project (unsafe) → Allow**. (It is your own script.)
3. Check the Sheet: tabs **Farmers, Batches, Reports, Dispatches, BagFlow, Settings** are filled with the
   12 sample farmers, 18 batches, 286 daily reports and feed-bag entries.
   Running it again does nothing if data already exists.

## 3. Publish the backend
1. **Deploy → New deployment → ⚙ Select type → Web app**.
2. *Execute as:* **Me**  ·  *Who has access:* **Anyone**.
3. **Deploy** → copy the **Web app URL** (ends with `/exec`).

## 4. Connect the website
1. Open **`index.html`** in a text editor, find
   `const CONFIG={API_URL:'PASTE_YOUR_WEB_APP_URL_HERE'};` and paste your URL.
2. Upload **every file** in this folder (`index.html, sw.js, manifest.webmanifest, icons`, and your
   **`photo1.png`**) to the root of your GitHub repo → Settings → Pages → `main` / root.
   (Replace the old site files.)

## 5. Sign in
| Role | ID | Password |
|---|---|---|
| Owner | `owner` | `Sunil@2026` (or what you set in step 1) |
| Sample farmers | `F1` … `F12` (or their phone number) | `farmer@123` |

New farmers: **Farmers → ＋ Add farmer** → you must enter a password. The app then offers a ready WhatsApp
message with the farmer code and password.

## How the new features work
- **One report per batch per day.** After submitting, the farmer sees a green “Today's report submitted”
  banner at the top (and a strip on every page). If the farmer entered something wrong, the owner opens
  **Reports → ↺ Allow re-entry**, which deletes that report so it can be submitted again.
- **Photos.** Feed-bag and weight photos are required; the mortality photo is required when deaths > 0.
  Photos are compressed on the phone and stamped with date and time. To make photos optional set
  `REQUIRE_PHOTOS: false` in `Code.gs`.
- **Farmer password change.** On the login screen (Farmer tab) or Profile → farmer enters code, full name,
  phone and village exactly as the owner saved them, then a new password. Wrong details are refused
  (5 tries per 15 minutes). Owner can also set a new password any time under Farmers → Edit.
- **Feed bags.** Owner **assigns** bags → farmer sees them → farmer **accepts** (or says “not received”).
  Farmer can **request** bags → owner approves (can change the number) or declines → farmer accepts.
  Stock in hand = accepted bags − bags used in daily reports. Owner sees every farmer's stock and a ledger.
- **Farmer photo.** Add a photo when adding/editing a farmer. The sample farmer F1 (Ramesh Yadav) uses
  the file `photo1.png` from your site folder; if the file is missing, initials are shown instead.

## Updating the code later
After editing `Code.gs`: **Deploy → Manage deployments → ✎ → Version: New version → Deploy**.
The web-app URL stays the same.

## Good to know
- Do not type into the Sheet tabs while farmers are working; use the app. (Viewing is fine.)
- Passwords are stored hashed (SHA-256 with a salt) — you cannot read them from the Sheet.
- Data from the old local-storage version is not migrated; the Sheet starts with the sample data.
- Every request after login carries a signed token; the public `/exec` URL alone gives no access to data.

## What's new in v3 (front-end only — no Apps Script redeploy needed)
- **Analytics** page (owner): period/farmer/status filters, mortality, feed, weight and reporting-rate stats, daily charts, deaths by bird age, farmer comparison and batch comparison tables.
- **Loading register** page (owner): filter by farmer, batch, dates and text; trip/bird totals, destination chart, CSV export, record a loading from one place.
- **Filters** on Farmers, Batches and Reports (search, status, sort, dates, batch, deaths-only) with live summary stats and a Clear button.
- **More details:** batch details now show age, peak death day, feed per 1000 birds, weight gain/day, reporting rate and 14-day charts; new **Farmer → Details** view.
- **Loading states:** skeleton screens when opening Analytics/Loading, spinners on buttons and the refresh icon.
Just replace `index.html` and `sw.js` in your repo (keep your `API_URL` line in index.html).
