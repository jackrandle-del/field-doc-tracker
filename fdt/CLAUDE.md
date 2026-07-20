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
Unused variables in App.js: SP_SITE, SP_FOLDER, getValidToken, programLabel, 
programColor, pickVersion, isMRF. Also a missing dependency (record.photo) 
in a useEffect.

## Workflow
- I review proposed changes before they're applied — show me a plan first 
  for anything non-trivial, then implement after I confirm
- Commit and push only after I explicitly approve
- I'm not a developer by background — explain technical tradeoffs in plain terms
