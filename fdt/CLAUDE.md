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
- Photos stored in IndexedDB (not localStorage — size limits)
- Records/projects stored in localStorage

## Deployment
- GitHub: jackrandle-del/field-doc-tracker (this code is in the `fdt` subfolder)
- Vercel: auto-deploys from `main` branch to field-doc-tracker-6n.vercel.app
- IMPORTANT: Vercel requires env var CI=false (CRA fails builds on ESLint 
  warnings otherwise, e.g. unused vars). Any new unused variable/import will 
  break production builds — check for lint warnings before pushing.

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
