# AASLT Control Dashboard

Static GitHub Pages dashboard for West Point Air Assault control operations. It presents a compiled operations database as a calendar, split Air Assault timelines, editable work ownership, next-day CM/MRE meal status, source-cell explorer, data diagnostics, S4 support lanes, and receipt register.

## Run Locally

From the repository root:

```sh
python3 -m http.server 8000
```

Open `http://localhost:8000/`. The dashboard is static; there is no backend service or build step.

## Update Data

The generated dashboard data lives in `data/aaslt-db.json`. To refresh it from the source workbooks, place the current workbook files where `scripts/extract_workbooks.py` expects them, then run from the repository root:

```sh
python3 scripts/extract_workbooks.py
```

Review the generated summary in the terminal, then commit the updated `data/aaslt-db.json` if it looks correct. The script also writes `data/schedule.json` as a legacy compatibility export, but the site reads the database file.

`data/aaslt-db.json` is the single static database for GitHub Pages. It contains normalized tables for events, source workbooks, worksheets, copied cells, LRTC-seeded work items, LRTC-seeded S4 support items, mess notes, and data diagnostics. The source Excel files are build inputs only; the published dashboard does not fetch or parse the workbooks directly. The extractor prefers `Total AASLT Cadre LRTC (2).xlsx` in Downloads as the current LRTC source, with older LRTC filenames only used as a local fallback if the newest workbook is missing.

The mess matrix extraction intentionally ignores Air Assault headcounts in the operational view. It treats the mess matrix as the authority for meal type. Cadet Mess (`CM`) meals use fixed meal times: 07:30 breakfast, 12:00 lunch, and 18:00 dinner. MRE meals (`M`, `M+`, `HM+`) use the LRTC schedule time when one is available, then fall back to planning defaults. The next-day meal board stays focused on whether each simultaneous Air Assault class is eating in the Cadet Mess or receiving MREs, with the source cells still available in the source explorer.

LRTC cells are the authority for timed timeline items, ownership work, and support requirements. Track inference uses explicit AA/WP labels, day numbers, and the LRTC color conventions where shared admin or medic columns need to attach a row to one of two simultaneous Air Assault tracks.

The dashboard generates a timed `08:00-11:00` MARNE medical FLA support card on each class `Day -2`. These generated cards are timed medical events, not all-day events.

PE/PT events such as Clean PE, PE1, PE2, PE3, and AA PT are normalized to South Dock.

Untimed source rows are hidden from the operational dashboard. This removes context-only cells such as `DAY X WP###`, location headers, and untimed classroom notes while preserving the day labels for the track headers and still using mess-matrix rows to generate timed meal events.

LRTC cadence, admin, and support cells are split before publishing. Trainee timeline cells stay on the Air Assault timeline; cadre/admin cells seed the work queue; medic-column FLA cells seed S4 FLA requests and, where timed, medical timeline cards.

This provided build intentionally embeds the copied source workbook cells in the static database because the operator requested that behavior. If the dashboard later needs secure uploads, role-based access, or private storage, the same data flow can be replaced or extended with SharePoint, Drive, or Supabase-backed storage.

## Publish With GitHub Pages

This folder includes `.github/workflows/pages.yml`, which publishes the static site from the repository root with GitHub Actions. Push the files to a GitHub repository named however you want, then enable **Settings -> Pages -> Source: GitHub Actions**.

From the repository root, commit and push the static site files:

```sh
git init
git branch -M main
git add .
git commit -m "Publish AASLT control dashboard"
git remote add origin https://github.com/YOUR-USER/YOUR-REPO.git
git push origin main
```

GitHub Pages will serve `index.html`, `data/aaslt-db.json`, `firebase-config.js`, and the other static assets directly.

## Firebase Sync

The dashboard can sync shared edits through Firebase. GitHub Pages hosts the static files; Firebase stores the live shared state.

1. In Firebase Console, create or open a Firebase project.
2. Add a Web app and copy its Firebase config object.
3. Enable Cloud Firestore.
4. Firebase Storage currently requires upgrading this project out of Spark, so this build leaves receipt files local-only unless Storage is later enabled.
5. Paste the Web app config into `firebase-config.js` and set `enabled: true`.
6. In Firestore rules, use `firestore.rules`.
7. If Storage is later enabled, put the bucket name back into `firebase-config.js` and use `storage.rules` if receipt files should be public to dashboard users.

The Firebase web config is not a password; it is expected to be public in browser apps. The included rules are intentionally open so anyone who can reach the site can read and write dashboard state. That matches the current requested operating model, but it is not secure against unwanted edits if the URL spreads.

Synced through Firestore:

- work queue items
- timeline event edits
- S4 support items
- deleted source-backed items, so deleted seed tasks do not reappear after reload
- receipt metadata/register

Synced through Firebase Storage when configured:

- uploaded hand receipt files, only if Firebase Storage is enabled; on Spark, receipt files stay local in the browser and the shared register syncs as metadata

## Local State

The dashboard still keeps a local fallback copy in the browser. If `firebase-config.js` is disabled or Firebase is unreachable, work queue items, timeline event edits, S4 support lanes, deleted-source records, and receipts continue working locally. Once Firebase sync is configured, the shared Firestore state becomes the live team state and local storage acts as an offline/fallback cache.
