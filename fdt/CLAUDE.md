# Doc Tracker

## Purpose
Mobile field documentation app for green building certification inspectors 
on multifamily new construction (MFNC) projects. Targets two certification 
programs: Energy Star MFNC (Rev03–Rev05) and EarthCraft (V6.5 and V7 — true pre-6.5 V6 is
out of scope, it only applies to older projects that won't live in this app).

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

- **Single Family portal — shown to users as "Audits"** (in progress, behind a per-browser switch): a second, much simpler portal
  next to multifamily. A browser opts in with `?sf=1` (`?sf=0` opts out; localStorage `fdt_sf_v1`);
  without it the app is exactly the multifamily tool and Single Family houses are hidden. With it, a
  chooser appears after login (last choice remembered in `fdt_portal_v1`, a "Single Family /
  Multifamily" switch button in the top bar). A house is a project doc with `type: "sf"` (absent = MF),
  `sfFolder: {id, name, path}`, and `sfItems: [{id, text}]`; its one category is "Documentation",
  holding a built-in photo-only "General uploads" item plus TA-named items (statuses are relabeled
  Verified / Issue found / N/A but stored as pass/fail/na). Reuses `ProjectDashboard`, `ChecklistView`
  and `ItemDetail` via `isSF()` / `getProjectCategories()` / `getProjectItems()`. Photos sync to the
  OneDrive of `viridiantdata@viridiant.org` (`SF_ONEDRIVE_URL`, folder `SF_ROOT_PATH`), into a NEW
  `Doc Tracker Documentation` folder inside the house's linked folder, then dated `MM.DD.YYYY`
  subfolders; uploads use rename-on-conflict so nothing is ever overwritten. Existing folders are
  never modified. Final Testing / MRF / programs / workbook are intentionally not part of it.
  Later phase: TA-provided per-house checklists populate `sfItems`.
  **Simplified (current UI):** opening a house shows just `SfPhotoBoard` — a "+ Add photos" button and a
  grid of photos, each with an optional per-photo note under it (stored as `note` on the photo's metadata
  entry via `photoMetaForSave`). All photos live on the `sf_general` record; the category/item/status
  screens and `sfItems` still exist in code and data but aren't reachable from the Single Family UI.
- Folder linking uses a shared tap-down `FolderBrowser` (read-only, folders only). Single Family starts at
  `SF_ROOT_PATH`; the multifamily New/Edit project form starts at the top of the MFProjects library and
  fills in the same plain-text `sharePointFolder` path as before (the typed/pasted box stays as a fallback).

## Deployment
- GitHub: jackrandle-del/field-doc-tracker (this code is in the `fdt` subfolder)
- Vercel: auto-deploys from `main` branch, live at **https://www.doctracker.org** (the link to
  hand TAs — the bare `doctracker.org` redirects to the `www` version, so use `www` directly to
  avoid the extra hop). `field-doc-tracker-6n.vercel.app` still works as a fallback alias.
- IMPORTANT: Vercel requires env var CI=false (CRA fails builds on ESLint 
  warnings otherwise, e.g. unused vars). Any new unused variable/import will 
  break production builds — check for lint warnings before pushing.
- The Microsoft OAuth app (`50108c90-8844-4fbc-96af-d4fb7e7fa4ca`, registered as a
  **Single-page application** platform — PKCE public client, no server/secret) must have
  whichever domain the app is actually served from listed as an authorized redirect URI, or
  SharePoint connect fails with AADSTS50011. `getRedirectUri()` in `src/App.js` builds this from
  `window.location.origin` at runtime, so moving domains again needs a matching Azure AD update
  but no code change.
- Server-side SharePoint sign-in (`api/ms-token.js`; **on by default; a browser can opt out with `?serverauth=0`**):
  Microsoft caps sign-ins for browser-only (SPA) apps at 24 hours, so this path has a Vercel
  Function add the client secret and gets ~90-day sign-ins instead. Needs a `MS_CLIENT_SECRET` env
  var in Vercel (no `REACT_APP_` prefix — that would expose it), a **Web**-platform redirect URI
  `https://www.doctracker.org/auth/callback` on the Azure app (the SPA redirect URI stays for the
  default flow), and `/auth/callback` + `/api/ms-token` allowed in any Vercel Firewall allowlist
  rule. **The client secret expires (max 24 months) — renew it before then or every server-side
  sign-in fails.** `GET /api/ms-token` returns `secretConfigured` as a setup check.
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
- **IMPORTANT — the app is live in production with real TAs/PMs using it during their actual
  workflow (since 2026-09-16).** Merging to `main` deploys immediately and affects everyone
  currently using the tool, not just test data. Do NOT merge a branch into `main` (or push
  directly to `main`) without my explicit authorization, given AFTER we've gone through QA
  together on the branch/change first. Work and test on a feature branch and wait for that
  go-ahead — don't assume urgency justifies skipping this, even for a fix. (Earlier in the
  project, immediate merge-to-main after every small fix was the norm, justified by SharePoint
  OAuth only working on the production domain — that pattern is retired now that this is live.)
