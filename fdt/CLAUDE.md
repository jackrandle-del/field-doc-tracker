# GreenCert Field Documentation App

## Purpose
Mobile field documentation app for green building certification inspectors 
on multifamily new construction (MFNC) projects. Targets two certification 
programs: Energy Star MFNC (Rev03–Rev05) and EarthCraft (V6 and V7).

## Core Philosophy
Only include items that are VISUALLY VERIFIABLE in the field — excluding 
anything requiring instrumentation, measurement equipment, or percentage-
tolerance testing. Explicit exceptions: delivery temperature at faucets, 
sone ratings for fans. When adding new checklist items, apply this filter 
first — if it needs a meter, gauge, or calculated result, it doesn't belong.

## Architecture
- Single-file React app: `src/App.js` (Create React App, no bundler config needed)
- Checklist items defined as arrays (e.g. EARTHCRAFT_CERTIFIED_V7), registered 
  in CHECKLIST_REGISTRY keyed by `programId||version||revision`
- Photos stored in IndexedDB (not localStorage — size limits), synced on demand to a
  per-project SharePoint folder (Microsoft OAuth — see `getValidToken`/`SP_SITE`)
- Projects/records stored in Firestore (`src/firebase.js`), shared across the whole team —
  every team member sees every project and can update any checklist item. `App()` subscribes
  via `onSnapshot`; writes happen at four call sites (`updateRecord`, `deleteProject`, the two
  `ProjectForm` `onSave` handlers) rather than one blanket save. Offline persistence is on, so
  the app keeps working with no signal (e.g. a TA in the field) and syncs once reconnected — no
  custom conflict handling, since simultaneous multi-device editing of the same project isn't
  a requirement.
- Access is gated by one shared team login (Firebase Authentication, email/password — not
  per-person accounts), fully separate from the per-user Microsoft/SharePoint OAuth used for
  photo uploads. Local naming: `auth`/`setAuth` is the SharePoint token; `fbAuth` (aliased
  import) and `teamUser`/`setTeamUser` are the Firebase/team-login state — don't confuse them.
- EarthCraft Multifamily support includes uploading a populated workbook (`.xlsx`, parsed via
  SheetJS) to auto-populate a project's optional/bonus checklist items and auto-pass mandatory
  items the workbook already confirms — see [[project_earthcraft_optional_points]] and
  [[project_earthcraft_workbook_structure]] in memory for the full design/parsing details.

## Deployment
- GitHub: jackrandle-del/field-doc-tracker (this code is in the `fdt` subfolder)
- Vercel: auto-deploys from `main` branch to field-doc-tracker-6n.vercel.app
- IMPORTANT: Vercel requires env var CI=false (CRA fails builds on ESLint 
  warnings otherwise, e.g. unused vars). Any new unused variable/import will 
  break production builds — check for lint warnings before pushing.
- Requires six `REACT_APP_FIREBASE_*` env vars (API key, auth domain, project ID, storage
  bucket, messaging sender ID, app ID) set in Vercel for Production, Preview, and Development —
  CRA bakes these in at build time, so a build that ran before they were saved needs a fresh
  deploy, not just the vars added. Firestore itself needs security rules requiring
  `request.auth != null` on the `projects` and `records` collections (Firebase console →
  Firestore Database → Rules).
- Vercel's own Deployment Protection (a login wall Vercel puts in front of deployments,
  separate from this app's TeamLogin) must stay OFF — our own team login already gates access,
  and technical advisors don't have Vercel accounts to get past a second one.

## Known cleanup needed (not urgent)
- `programLabel`, `programColor`, `pickVersion` (unused dead code) and the 
  no-op `isMRF` in `ProjectDashboard` were removed (2026-08).
- SP_SITE and getValidToken are NOT unused — both are core to the SharePoint 
  photo upload feature. SP_FOLDER no longer exists (removed when uploads 
  moved from one shared folder to a folder linked per project).
- Still open: a missing dependency (`record.photos`) in the photo-loading 
  useEffect in `ItemDetail` — intentional (only want to reload from IndexedDB 
  when `photoKey` changes, not on every record update), but still flagged by 
  ESLint's exhaustive-deps rule.

## Workflow
- I review proposed changes before they're applied — show me a plan first 
  for anything non-trivial, then implement after I confirm
- Commit and push only after I explicitly approve
- I'm not a developer by background — explain technical tradeoffs in plain terms
