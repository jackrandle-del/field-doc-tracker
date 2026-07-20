import React, { useState, useEffect, useRef } from "react";

// ── INDEXEDDB PHOTO STORAGE ───────────────────────────────────────────────────
// Photos are stored here instead of localStorage — handles large binary data
// without the size constraints that cause silent save failures.
const IDB_NAME = 'fdt_photos_v1';
const IDB_STORE = 'photos';

const openPhotoDB = () => new Promise((resolve, reject) => {
  const req = indexedDB.open(IDB_NAME, 1);
  req.onupgradeneeded = e => e.target.result.createObjectStore(IDB_STORE);
  req.onsuccess = e => resolve(e.target.result);
  req.onerror = e => reject(e.target.error);
});

const idbSavePhoto = async (key, dataUrl) => {
  const db = await openPhotoDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, 'readwrite');
    tx.objectStore(IDB_STORE).put(dataUrl, key);
    tx.oncomplete = () => resolve();
    tx.onerror = e => reject(e.target.error);
  });
};

const idbGetPhoto = async (key) => {
  const db = await openPhotoDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, 'readonly');
    const req = tx.objectStore(IDB_STORE).get(key);
    req.onsuccess = e => resolve(e.target.result || null);
    req.onerror = e => reject(e.target.error);
  });
};

const idbDeletePhoto = async (key) => {
  const db = await openPhotoDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, 'readwrite');
    tx.objectStore(IDB_STORE).delete(key);
    tx.oncomplete = () => resolve();
    tx.onerror = e => reject(e.target.error);
  });
};

// ── SHAREPOINT AUTH (OAuth2 PKCE) ─────────────────────────────────────────────
const CLIENT_ID   = "50108c90-8844-4fbc-96af-d4fb7e7fa4ca";
const TENANT_ID   = "ba936175-44a3-4888-b75e-6f814421a09c";
const SP_SITE     = "https://ecva.sharepoint.com/sites/MFProjects";
const SP_FOLDER   = "Shared Documents/Field Documentation Tracker - Folder Link";
const SCOPES      = "Files.ReadWrite.All Sites.ReadWrite.All User.Read offline_access";
const AUTH_KEY    = "fdt_auth_v1";

const getRedirectUri = () => window.location.origin + window.location.pathname.replace(/\/$/, '');

// PKCE helpers
const genVerifier = () => {
  const arr = new Uint8Array(32);
  crypto.getRandomValues(arr);
  return btoa(String.fromCharCode(...arr)).replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');
};
const genChallenge = async (v) => {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(v));
  return btoa(String.fromCharCode(...new Uint8Array(digest))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');
};

const startLogin = async () => {
  const verifier = genVerifier();
  const challenge = await genChallenge(verifier);
  sessionStorage.setItem('fdt_verifier', verifier);
  const p = new URLSearchParams({
    client_id: CLIENT_ID, response_type: 'code',
    redirect_uri: getRedirectUri(), scope: SCOPES,
    code_challenge: challenge, code_challenge_method: 'S256',
    response_mode: 'query',
  });
  window.location.href = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?${p}`;
};

const exchangeCode = async (code) => {
  const verifier = sessionStorage.getItem('fdt_verifier');
  sessionStorage.removeItem('fdt_verifier');
  const body = new URLSearchParams({
    client_id: CLIENT_ID, grant_type: 'authorization_code',
    code, redirect_uri: getRedirectUri(),
    code_verifier: verifier, scope: SCOPES,
  });
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body,
  });
  return res.json();
};

const refreshAuth = async (refreshToken) => {
  const body = new URLSearchParams({
    client_id: CLIENT_ID, grant_type: 'refresh_token',
    refresh_token: refreshToken, scope: SCOPES,
  });
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body,
  });
  return res.json();
};

const fetchUserInfo = async (token) => {
  const res = await fetch('https://graph.microsoft.com/v1.0/me', {
    headers: { Authorization: `Bearer ${token}` }
  });
  return res.json();
};

const loadAuth = () => { try { return JSON.parse(localStorage.getItem(AUTH_KEY)) || null; } catch { return null; } };
const saveAuth = (a) => localStorage.setItem(AUTH_KEY, JSON.stringify(a));
const clearAuth = () => localStorage.removeItem(AUTH_KEY);

// Returns a valid access token, refreshing if needed. Returns null if not authenticated.
const getValidToken = async (auth, setAuth) => {
  if (!auth) return null;
  if (auth.expiresAt > Date.now() + 300000) return auth.accessToken; // still valid
  try {
    const result = await refreshAuth(auth.refreshToken);
    if (result.access_token) {
      const updated = { ...auth, accessToken: result.access_token,
        expiresAt: Date.now() + result.expires_in * 1000,
        refreshToken: result.refresh_token || auth.refreshToken };
      saveAuth(updated); setAuth(updated);
      return updated.accessToken;
    }
  } catch {}
  clearAuth(); setAuth(null); return null;
};

// ── SHAREPOINT PHOTO UPLOAD (Microsoft Graph) ─────────────────────────────────
// Resolves SP_SITE to a Graph site id once per session, then reuses it.
let _cachedSiteId = null;
const getSharePointSiteId = async (token) => {
  if (_cachedSiteId) return _cachedSiteId;
  const url = new URL(SP_SITE);
  const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${url.hostname}:${url.pathname}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const json = await res.json();
  if (!json.id) throw new Error(json.error?.message || "Could not resolve SharePoint site");
  _cachedSiteId = json.id;
  return _cachedSiteId;
};

const extFromDataUrl = (dataUrl) => {
  const m = /^data:image\/(\w+)/.exec(dataUrl);
  const type = (m?.[1] || "jpeg").toLowerCase();
  return type === "jpeg" ? "jpg" : type;
};

// Strips characters SharePoint disallows in file/folder names
const sanitizeSpName = (name) => name.replace(/[\\/:*?"<>|]/g, "-").trim();

// SP_FOLDER is expressed relative to the "Shared Documents" library, which is the
// default document library's drive root — so that prefix is dropped for Graph.
const SP_FOLDER_PATH = SP_FOLDER.replace(/^Shared Documents\/?/, "");

const uploadPhotoToSharePoint = async (token, fileName, dataUrl) => {
  const siteId = await getSharePointSiteId(token);
  const blob = await (await fetch(dataUrl)).blob();
  const pathSegments = `${SP_FOLDER_PATH}/${fileName}`.split("/").map(encodeURIComponent).join("/");
  const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${pathSegments}:/content`, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": blob.type || "image/jpeg" },
    body: blob,
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Upload failed (${res.status})`);
  }
  return res.json();
};

const EARTHCRAFT_CERTIFIED_V7 = [
  // ── SITE PLANNING ──────────────────────────────────────────────────────────
    { id: "ec_sp2_7", pointNumber: "SP 2.7", tier: "ALL", text: "Outdoor community gathering space provided on site", category: "Site Planning" },
  { id: "ec_du1_5",  pointNumber: "DU 1.5",  tier: "ALL",  text: "Maintain 2\" clearance between wall siding and roof surface", category: "Durability & Moisture Management" },
  { id: "ec_du1_6",  pointNumber: "DU 1.6",  tier: "ALL",  text: "Install level air conditioner condensing unit pad", category: "Durability & Moisture Management" },
  { id: "ec_du1_7",  pointNumber: "DU 1.7",  tier: "ALL",  text: "Roof drip edge with ≥ 1/4\" overhang", category: "Durability & Moisture Management" },
  { id: "ec_nc_du2_5", pointNumber: "DU 2.5",  tier: "ALL",  text: "Do not install wet or water-damaged building materials", category: "Durability & Moisture Management" },
  { id: "ec_du2_8",  pointNumber: "DU 2.8",  tier: "ALL",  text: "Design for or install additional dehumidification: rough-in electrical/plumbing for dehumidifier OR install whole-unit ENERGY STAR dehumidifier", category: "Durability & Moisture Management" },
  { id: "ec_iaq1",   pointNumber: "IAQ 1",   tier: "ALL",  text: "No unvented combustion fireplaces, appliances, or space heaters; all combustion appliances mechanically drafted or direct-vented (EarthCraft IAQ 1 / Energy Star 10.1–10.3)", category: "Indoor Air Quality", mergedWith: ["es_10_1","es_10_2","es_10_3"] },
  { id: "ec_iaq1_2", pointNumber: "IAQ 1.2", tier: "ALL",  text: "Sealed-combustion or electric water heater installed within thermal envelope; no unit-level atmospherically vented water heaters or furnaces (EarthCraft IAQ 1.1–1.2 / Energy Star ES 5.0)", category: "Indoor Air Quality", mergedWith: ["es_10_1","ec_es5_0_ref"] },
  { id: "ec_iaq1_3", pointNumber: "IAQ 1.3", tier: "ALL",  text: "Carbon monoxide detector installed if combustion appliances exist (one per unit)", category: "Indoor Air Quality" },
  { id: "ec_iaq2",   pointNumber: "IAQ 2",   tier: "ALL",  text: "Protect all ducts and indoor coils until floor/wall finishing is complete", category: "Indoor Air Quality" },
  { id: "ec_iaq2_1", pointNumber: "IAQ 2.1", tier: "ALL",  text: "Filter is easily accessible for property maintenance; MERV 6+ minimum installed in each ducted system; all return and outdoor air passes through filter prior to distribution (EarthCraft IAQ 2.1–2.3 / Energy Star 9.1)", category: "Indoor Air Quality", mergedWith: ["es_9_1"] },
  { id: "ec_iaq2_5", pointNumber: "IAQ 2.5", tier: "ALL",  text: "No carpet in below-grade units", category: "Indoor Air Quality" },
  { id: "ec_es5_1",  pointNumber: "ES 5.1",  tier: "ALL",  text: "Heat trap on all storage water heaters; confirm presence by visual inspection or AHRI certificate (EarthCraft ES 5.1 / Energy Star 11.3)", category: "Water Efficiency", mergedWith: ["es_11_3"] },
  { id: "ec_es5_3",  pointNumber: "ES 5.3",  tier: "ALL",  text: "Pipe insulation on first 2' of hot and cold water pipes at water heater", category: "Water Efficiency" },
  { id: "ec_we1_0",  pointNumber: "WE 1.0",  tier: "ALL",  text: "Meet National Energy Policy Act low-flow standards for all fixtures", category: "Water Efficiency" },
  { id: "ec_we1_1",  pointNumber: "WE 1.1",  tier: "ALL",  text: "Detect and repair all leaks at water-using fixtures, appliances, and equipment", category: "Water Efficiency" },
  { id: "ec_we1_2",  pointNumber: "WE 1.2",  tier: "ALL",  text: "Low-flow fixtures throughout: WaterSense toilet ≤1.28 gpf; WaterSense urinal ≤0.5 gpf; WaterSense lavatory faucet ≤1.5 gpm; WaterSense showerhead ≤2.0 gpm (EarthCraft WE 1.2 / Energy Star 13.2)", category: "Water Efficiency", mergedWith: ["es_13_2"] },
  { id: "ec_du2_6",  pointNumber: "DU 2.6",  tier: "ALL",  text: "Newly installed and existing plants maintain distance ≥2' from building at maturity", category: "Durability & Moisture Management" },
  { id: "ec_v7_re1_0", pointNumber: "RE 1.0", text: "Limit framing at all windows and doors", category: "Resource Efficiency", tier: "ALL" },
  { id: "ec_v7_re1_1", pointNumber: "RE 1.1", text: "Engineered roof framing (90%)", category: "Resource Efficiency", tier: "ALL" },
  { id: "ec_v7_du1_6", pointNumber: "DU 1.6", text: "Continuous foundation termite flashing (required if slab edge is insulated)", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du1_10", pointNumber: "DU 1.10", text: "Drain pan installed for all water heaters and washing machines", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du2_6", pointNumber: "DU 2.6", text: "Capillary break between foundation and framing at all exterior walls", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du2_7", pointNumber: "DU 2.7", text: "Drainage board and damp proofing installed for all below-grade walls", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_iaq1_1", pointNumber: "IAQ 1.1", text: "All fireplaces have outdoor combustion air supply; masonry-built fireplaces have gasketed doors", category: "Indoor Air Quality", tier: "ALL" },
  { id: "ec_v7_iaq2_2", pointNumber: "IAQ 2.2", text: "Rodent and corrosion-proof screens with mesh ≤0.5\" provided for all openings not fully sealed or caulked", category: "Indoor Air Quality", tier: "ALL" },
  { id: "ec_v7_iaq2_3", pointNumber: "IAQ 2.3", text: "All outdoor supply air crosses a filter prior to distribution", category: "Indoor Air Quality", tier: "ALL" },
  { id: "ec_v7_be3_9", pointNumber: "BE 3.9", text: "Slab edge insulation ≥ R-10", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be3_10", pointNumber: "BE 3.10", text: "Insulation installation quality: Grade I throughout OR Grade II with continuous insulated sheathing ≥ R-3 (100% coverage)", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_we1_3", pointNumber: "WE 1.3", text: "Hot water recirculation system uses manual demand or presence sensor controls", category: "Water Efficiency", tier: "ALL" }
];

const EARTHCRAFT_GOLD_V7 = [
  { id: "ec_du1_5",  pointNumber: "DU 1.5",  tier: "ALL",  text: "Maintain 2\" clearance between wall siding and roof surface", category: "Durability & Moisture Management" },
  { id: "ec_du1_6",  pointNumber: "DU 1.6",  tier: "ALL",  text: "Install level air conditioner condensing unit pad", category: "Durability & Moisture Management" },
  { id: "ec_du1_7",  pointNumber: "DU 1.7",  tier: "ALL",  text: "Roof drip edge with ≥ 1/4\" overhang", category: "Durability & Moisture Management" },
  { id: "ec_nc_du2_5", pointNumber: "DU 2.5",  tier: "ALL",  text: "Do not install wet or water-damaged building materials", category: "Durability & Moisture Management" },
  { id: "ec_du2_8",  pointNumber: "DU 2.8",  tier: "ALL",  text: "Design for or install additional dehumidification: rough-in electrical/plumbing for dehumidifier OR install whole-unit ENERGY STAR dehumidifier", category: "Durability & Moisture Management" },
  { id: "ec_iaq1",   pointNumber: "IAQ 1",   tier: "ALL",  text: "No unvented combustion fireplaces, appliances, or space heaters; all combustion appliances mechanically drafted or direct-vented (EarthCraft IAQ 1 / Energy Star 10.1–10.3)", category: "Indoor Air Quality", mergedWith: ["es_10_1","es_10_2","es_10_3"] },
  { id: "ec_iaq1_2", pointNumber: "IAQ 1.2", tier: "ALL",  text: "Sealed-combustion or electric water heater installed within thermal envelope; no unit-level atmospherically vented water heaters or furnaces (EarthCraft IAQ 1.1–1.2 / Energy Star ES 5.0)", category: "Indoor Air Quality", mergedWith: ["es_10_1","ec_es5_0_ref"] },
  { id: "ec_iaq1_3", pointNumber: "IAQ 1.3", tier: "ALL",  text: "Carbon monoxide detector installed if combustion appliances exist (one per unit)", category: "Indoor Air Quality" },
  { id: "ec_iaq2",   pointNumber: "IAQ 2",   tier: "ALL",  text: "Protect all ducts and indoor coils until floor/wall finishing is complete", category: "Indoor Air Quality" },
  { id: "ec_iaq2_1", pointNumber: "IAQ 2.1", tier: "ALL",  text: "Filter is easily accessible for property maintenance; MERV 6+ minimum installed in each ducted system; all return and outdoor air passes through filter prior to distribution (EarthCraft IAQ 2.1–2.3 / Energy Star 9.1)", category: "Indoor Air Quality", mergedWith: ["es_9_1"] },
  { id: "ec_iaq2_5", pointNumber: "IAQ 2.5", tier: "ALL",  text: "No carpet in below-grade units", category: "Indoor Air Quality" },
  { id: "ec_es5_1",  pointNumber: "ES 5.1",  tier: "ALL",  text: "Heat trap on all storage water heaters; confirm presence by visual inspection or AHRI certificate (EarthCraft ES 5.1 / Energy Star 11.3)", category: "Water Efficiency", mergedWith: ["es_11_3"] },
  { id: "ec_es5_3",  pointNumber: "ES 5.3",  tier: "ALL",  text: "Pipe insulation on first 2' of hot and cold water pipes at water heater", category: "Water Efficiency" },
  { id: "ec_we1_0",  pointNumber: "WE 1.0",  tier: "ALL",  text: "Meet National Energy Policy Act low-flow standards for all fixtures", category: "Water Efficiency" },
  { id: "ec_we1_1",  pointNumber: "WE 1.1",  tier: "ALL",  text: "Detect and repair all leaks at water-using fixtures, appliances, and equipment", category: "Water Efficiency" },
  { id: "ec_we1_2",  pointNumber: "WE 1.2",  tier: "ALL",  text: "Low-flow fixtures throughout: WaterSense toilet ≤1.28 gpf; WaterSense urinal ≤0.5 gpf; WaterSense lavatory faucet ≤1.5 gpm; WaterSense showerhead ≤2.0 gpm (EarthCraft WE 1.2 / Energy Star 13.2)", category: "Water Efficiency", mergedWith: ["es_13_2"] },
  { id: "ec_du2_6",  pointNumber: "DU 2.6",  tier: "ALL",  text: "Newly installed and existing plants maintain distance ≥2' from building at maturity", category: "Durability & Moisture Management" },
  { id: "ec_v7_re1_0", pointNumber: "RE 1.0", text: "Limit framing at all windows and doors", category: "Resource Efficiency", tier: "ALL" },
  { id: "ec_v7_re1_1", pointNumber: "RE 1.1", text: "Engineered roof framing (90%)", category: "Resource Efficiency", tier: "ALL" },
  { id: "ec_v7_du1_6", pointNumber: "DU 1.6", text: "Continuous foundation termite flashing (required if slab edge is insulated)", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du1_10", pointNumber: "DU 1.10", text: "Drain pan installed for all water heaters and washing machines", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du2_6", pointNumber: "DU 2.6", text: "Capillary break between foundation and framing at all exterior walls", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du2_7", pointNumber: "DU 2.7", text: "Drainage board and damp proofing installed for all below-grade walls", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_iaq1_1", pointNumber: "IAQ 1.1", text: "All fireplaces have outdoor combustion air supply; masonry-built fireplaces have gasketed doors", category: "Indoor Air Quality", tier: "ALL" },
  { id: "ec_v7_iaq2_2", pointNumber: "IAQ 2.2", text: "Rodent and corrosion-proof screens with mesh ≤0.5\" provided for all openings not fully sealed or caulked", category: "Indoor Air Quality", tier: "ALL" },
  { id: "ec_v7_iaq2_3", pointNumber: "IAQ 2.3", text: "All outdoor supply air crosses a filter prior to distribution", category: "Indoor Air Quality", tier: "ALL" },
  { id: "ec_v7_be3_9", pointNumber: "BE 3.9", text: "Slab edge insulation ≥ R-10", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be3_10", pointNumber: "BE 3.10", text: "Insulation installation quality: Grade I throughout OR Grade II with continuous insulated sheathing ≥ R-3 (100% coverage)", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_we1_3", pointNumber: "WE 1.3", text: "Hot water recirculation system uses manual demand or presence sensor controls", category: "Water Efficiency", tier: "ALL" },
  { id: "ec_du2_7",  pointNumber: "DU 2.7",  tier: "GOLD", text: "If installed, drain at outside perimeter edge of footing surrounded with 6\" clearstone and filter fabric", category: "Durability & Moisture Management" },
  { id: "ec_iaq2_6", pointNumber: "IAQ 2.6", tier: "GOLD", text: "Filters are ≥ MERV 8", category: "Indoor Air Quality" },
  { id: "ec_v7_re1_2", pointNumber: "RE 1.2", text: "Advanced framing: 2-stud corners where structurally feasible; ladder T-walls where structurally feasible; headers sized for actual loads", category: "Resource Efficiency", tier: "GOLD" },
  { id: "ec_v7_du2_9", pointNumber: "DU 2.9", text: "Additional dehumidification system installed: basement or sealed crawlspace system", category: "Durability & Moisture Management", tier: "GOLD" },
  { id: "ec_v7_iaq1_5", pointNumber: "IAQ 1.5", text: "If installed, all fireplaces meet indoor air quality guidelines and have gasketed doors", category: "Indoor Air Quality", tier: "GOLD" },
  { id: "ec_v7_be1_14", pointNumber: "BE 1.14", text: "Top plate sealed to drywall at all levels", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be3_11", pointNumber: "BE 3.11", text: "Corners insulated to ≥ R-6", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be3_12", pointNumber: "BE 3.12", text: "Headers insulated to ≥ R-3", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be3_13", pointNumber: "BE 3.13", text: "Fiberglass batts are unfaced and friction-fit throughout", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be5_0", pointNumber: "BE 5.0", text: "Ducts in unconditioned attic: buried in R-49 insulation OR ducts with R-8 insulation encapsulated in 1.5\" closed-cell foam and buried under ≥2\" blown insulation", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_es1_11", pointNumber: "ES 1.11", text: "HVAC equipment is ENERGY STAR qualified; for split systems, the pairing must be qualified", category: "Energy Efficient Systems", tier: "GOLD" }
];



// ─── PROGRAM CATALOG ─────────────────────────────────────────────────────────
// Each program entry has: id, label, color, versions[]
// Each version: { version, revisions[] }
const PROGRAM_CATALOG = [
  {
    id: "energy_star_mfnc",
    label: "Energy Star MFNC",
    color: "#0D9488",
    versions: [
      { version: "1 / 1.1 / 1.2", revisions: ["Rev. 03", "Rev. 04"] },
      { version: "1.1 / 1.2 / 1.3", revisions: ["Rev. 05"] },
    ],
  },
  {
    id: "earthcraft_certified",
    label: "EarthCraft Certified",
    color: "#2D6A4F",
    versions: [
      { version: "V6", revisions: ["New Construction"] },
          { version: "V7", revisions: ["New Construction"] },
    ],
  },
  {
    id: "earthcraft_gold",
    label: "EarthCraft Gold",
    color: "#1B4332",
    versions: [
      { version: "V6", revisions: ["New Construction"] },
      { version: "V7", revisions: ["New Construction"] },
    ],
  },
  {
    id: "earthcraft_sf2024_certified",
    label: "EarthCraft Certified - Southface",
    color: "#2D6A4F",
    versions: [{ version: "v2024", revisions: ["Southface"] }],
  },
  {
    id: "earthcraft_sf2024_gold",
    label: "EarthCraft Gold - Southface",
    color: "#1B4332",
    versions: [{ version: "v2024", revisions: ["Southface"] }],
  },
];

const CATEGORIES = [
  { id: "Site Planning",                    code: "SP"  },
  { id: "Construction Waste Management",    code: "CW"  },
  { id: "Resource Efficiency",              code: "RE"  },
  { id: "Durability & Moisture Management", code: "DU"  },
  { id: "Indoor Air Quality",               code: "IAQ" },
  { id: "High Performance Building Envelope", code: "BE" },
  { id: "Energy Efficient Systems",         code: "ES"  },
  { id: "Water Efficiency",                 code: "WE"  },
  { id: "Education & Operations",           code: "EO"  },
  { id: "Minimum Rated Features",           code: "MRF" },
];

// ─── ENERGY STAR MFNC v1/1.1/1.2 Rev.03 — 107 items ────────────────────────
const ENERGY_STAR_MFNC_V1_REV03 = [
  { id: "es_1_5_3", pointNumber: "1.5.3", text: "Heated plenums: bottom has at least R-13 insulation", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_1_6_1", pointNumber: "1.6.1", text: "Garages with space heating: walls insulated ≥ R-5ci (CZ 5-6), ≥ R-7.5ci (CZ 7), ≥ R-9.5ci (CZ 8)", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_1_3", pointNumber: "1.3", text: "All insulation achieves Grade I installation per ANSI/RESNET/ICC 301", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_2_1", pointNumber: "2.1", text: "Air barrier fully aligned: dropped ceilings/soffits below unconditioned attics, chase/dead space, and all other ceilings", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_2_2", pointNumber: "2.2", text: "Air barrier fully aligned: walls behind showers, tubs, staircases, and fireplaces", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_2_3", pointNumber: "2.3", text: "Air barrier fully aligned: architectural bump-outs, dead space, and all other exterior walls", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_2_4", pointNumber: "2.4", text: "Air barrier fully aligned: floors above garages, floors above unconditioned spaces, and cantilevered floors", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_2_5", pointNumber: "2.5", text: "Air barrier fully aligned: all other floors adjoining unconditioned space (rim/band joists at exterior wall or porch roof)", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_3_1", pointNumber: "3.1", text: "Insulated ceilings with attic above: Grade I insulation extends to inside face of exterior wall; ≥ R-21 (CZ 1-5), ≥ R-30 (CZ 6-8)", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_3_2", pointNumber: "3.2", text: "Attic access panels and drop-down stairs insulated ≥ R-10 or equipped with durable ≥ R-10 cover", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_3_3", pointNumber: "3.3", text: "Insulation beneath attic platforms (HVAC platforms, walkways) ≥ R-21 (CZ 1-5), ≥ R-30 (CZ 6-8)", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_3_4", pointNumber: "3.4", text: "Slabs on grade in CZ 4-8: 100% of slab edge insulated to ≥ R-5 per 2009 IECC Table 502.2(1), aligned with thermal boundary", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_3_5", pointNumber: "3.5", text: "Above-grade concrete slab edges (podiums, balconies) in CZ 4-8: 100% of slab edge insulated to ≥ R-5, aligned with thermal boundary", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_3_6", pointNumber: "3.6", text: "Concrete slab floors in CZ 4-8 above ambient/garages/unconditioned spaces: floor insulation meets U-factor per 2009 IECC Table 502.1.2", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_3_7_1", pointNumber: "3.7.1", text: "Above-grade walls and rim/band joists: continuous rigid insulation or insulated siding ≥ R-3 (CZ 1-4), ≥ R-5 (CZ 5-8)", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_3_7_2", pointNumber: "3.7.2", text: "Above-grade walls: Structural Insulated Panels OR Insulated Concrete Forms OR Double-wall framing", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_3_7_3a", pointNumber: "3.7.3a", text: "Advanced framing: corners insulated ≥ R-6 to edge", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_3_7_3b", pointNumber: "3.7.3b", text: "Advanced framing: headers above windows & doors insulated ≥ R-3 (2x4 framing) or ≥ R-5 (all other assemblies)", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_3_7_3c", pointNumber: "3.7.3c", text: "Advanced framing: interior/exterior wall intersections insulated to same R-value as rest of exterior wall", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_1", pointNumber: "4.1", text: "Ducts, flues, shafts, plumbing, piping, wiring, exhaust fans & penetrations to unconditioned space sealed with blocking/flashing as needed", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_2", pointNumber: "4.2", text: "Recessed lighting adjacent to unconditioned space: ICAT labeled and gasketed; exterior surface insulated ≥ R-10 in CZ 4-8", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_3", pointNumber: "4.3", text: "Continuous top plate or blocking at top of walls adjoining unconditioned space (including balloon-framed parapets) sealed", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_4", pointNumber: "4.4", text: "Drywall sealed to top plate at all unconditioned attic/wall interfaces using caulk, foam, drywall adhesive, or equivalent", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_5", pointNumber: "4.5", text: "Rough opening around windows & exterior doors sealed", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_6", pointNumber: "4.6", text: "Assemblies separating attached garages from occupiable space sealed; air barrier installed, sealed, and aligned", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_7", pointNumber: "4.7", text: "Doors adjacent to unconditioned space or ambient made substantially air-tight with doorsweep and weatherstripping", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_8", pointNumber: "4.8", text: "Attic access panels, roof hatches and drop-down stairs gasketed or equipped with durable gasketed covers", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_9", pointNumber: "4.9", text: "Unit entrance doors from corridor/stairwell made substantially air-tight with doorsweep and weatherstripping", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_5_9", pointNumber: "5.9", text: "All heating and cooling systems serving a dwelling unit have thermostatic controls within the unit", category: "Energy Efficient Systems", mandatory: true },

  { id: "es_5_10", pointNumber: "5.10", text: "Stair and elevator shaft vents equipped with motorized dampers; verified closed at time of inspection", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_5_11", pointNumber: "5.11", text: "Freeze protection systems include automatic controls to shut off when pipe wall/garage/plenum temperatures above 40°F", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_5_11_1", pointNumber: "5.11.1", text: "Heat tracing for freeze protection: controls based on pipe wall temperature and minimum R-3 pipe insulation", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_5_12", pointNumber: "5.12", text: "Snow/ice-melting systems: automatic controls to shut off when pavement above 50°F with no precipitation", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_5_13", pointNumber: "5.13", text: "Hydronic systems: terminal heating/cooling equipment separated from riser by control valve or terminal pump", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_5_14", pointNumber: "5.14", text: "Hydronic systems: terminal units equipped with pressure independent balancing valves or control valves", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_5_15", pointNumber: "5.15", text: "Hydronic systems: piping insulated per Item 4.42 of National HVAC Design Report including at planks and penetrations", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_5_16", pointNumber: "5.16", text: "Hydronic circulating pumps ≥1 HP, 3-phase: NEMA Premium motors; ≥5 HP also with variable frequency drives", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_6_1", pointNumber: "6.1", text: "Ductwork installed without kinks, sharp bends, compressions, or excessive coiled flexible ductwork", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_6_3", pointNumber: "6.3", text: "All supply and return ducts in unconditioned space insulated to ≥ R-6, including connections to trunk ducts", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_6_6", pointNumber: "6.6", text: "Common Space: supply, return, and exhaust ductwork and plenums sealed at all transverse joints, seams, and penetrations", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_7_4", pointNumber: "7.4", text: "Ventilation override control installed and labeled if function not obvious; townhouses: readily accessible to occupant", category: "Indoor Air Quality", mandatory: true },
  { id: "es_7_5_1", pointNumber: "7.5.1", text: "Outdoor air inlet on ducted return: motorized damper automatically restricts airflow during vent off-cycle and occupant override", category: "Indoor Air Quality", mandatory: false },
  { id: "es_7_6", pointNumber: "7.6", text: "System fan in dwelling unit rated ≤ 3 sones if intermittent, ≤ 2 sones if continuous, or exempted", category: "Indoor Air Quality", mandatory: true },
  { id: "es_7_7", pointNumber: "7.7", text: "If Vent System controller operates HVAC fan: fan operation is intermittent and ECM/ICM type OR controls reduce runtime for HVAC hours", category: "Indoor Air Quality", mandatory: true },
  { id: "es_7_8", pointNumber: "7.8", text: "In-unit bathroom fans or in-line fans are ENERGY STAR certified if used as part of dwelling-unit mechanical ventilation system", category: "Indoor Air Quality", mandatory: true },
  { id: "es_7_10_1", pointNumber: "7.10.1", text: "Ventilation air inlets pull air directly from outdoors, not from attic, crawlspace, garage, or adjacent dwelling unit", category: "Indoor Air Quality", mandatory: false },
  { id: "es_7_10_2", pointNumber: "7.10.2", text: "Inlets ≥2 ft above grade or roof deck; ≥10 ft from known contamination sources; ≥3 ft from dryer exhausts and roof sources", category: "Indoor Air Quality", mandatory: false },
  { id: "es_7_10_3", pointNumber: "7.10.3", text: "Inlets provided with rodent/insect screen with ≤ 0.5 inch mesh", category: "Indoor Air Quality", mandatory: false },
  { id: "es_8_2", pointNumber: "8.2", text: "Bathroom mechanical exhaust: continuous ≥ 20 CFM / ≤ 2 sones; intermittent ≥ 50 CFM; vented directly to outdoors", category: "Indoor Air Quality", mandatory: true },
  { id: "es_8_4", pointNumber: "8.4", text: "Shared garage exhaust system equipped with controls that sense CO and NO2", category: "Indoor Air Quality", mandatory: true },
  { id: "es_9_1", pointNumber: "9.1", text: "MERV 6+ filter(s) in each ducted mechanical system serving a dwelling unit; all return and outdoor air passes through filter", category: "Indoor Air Quality", mandatory: true },
  { id: "es_9_1_1", pointNumber: "9.1.1", text: "Filter access panel includes gasket and fits snugly against exposed edge of filter when closed to prevent bypass", category: "Indoor Air Quality", mandatory: true },
  { id: "es_10_1", pointNumber: "10.1", text: "Furnaces, boilers, and water heaters within pressure boundary are mechanically drafted or direct-vented", category: "Indoor Air Quality", mandatory: true },
  { id: "es_10_2", pointNumber: "10.2", text: "Fireplaces within pressure boundary are direct-vented", category: "Indoor Air Quality", mandatory: true },
  { id: "es_10_3", pointNumber: "10.3", text: "No unvented combustion appliances other than cooking ranges or ovens inside pressure boundary", category: "Indoor Air Quality", mandatory: true },
  { id: "es_11_3", pointNumber: "11.3", text: "In-unit storage water heaters: heat trap confirmed by visual inspection or on AHRI certificate", category: "Water Efficiency", mandatory: true },
  { id: "es_12_1_1", pointNumber: "12.1.1", text: "ERI Path: All common spaces (except lobby, mechanical rooms, and safety exceptions) have occupancy sensors or automatic bi-level lighting controls installed and operation verified", category: "Energy Efficient Systems", mandatory: true },

  { id: "es_12_2", pointNumber: "12.2", text: "Exterior lighting: automatic switching on timers or photocell controls, except 24-hour, security, or individual unit meter fixtures", category: "Energy Efficient Systems", mandatory: true },
  { id: "es_12_3", pointNumber: "12.3", text: "Common spaces and garages: 90% of installed lighting fixtures are integrated LED or contain LED lamps", category: "Energy Efficient Systems", mandatory: true },


  { id: "es_13_2", pointNumber: "13.2", text: "ERI Path: Common space refrigerators and dishwashers are ENERGY STAR certified; showerheads are WaterSense labeled", category: "Energy Efficient Systems", mandatory: true },
];

// ─── CHECKLIST REGISTRY ───────────────────────────────────────────────────────
// Maps programId + version + revision -> items array

// Rev.04 (v1/1.1/1.2, 02/15/2024) — same structure as Rev.03, functionally identical
// for ERI path. 12.1.1 adds "programmed timers" as lighting control option.
// We share the Rev.03 data and note the minor difference inline.
const ENERGY_STAR_MFNC_V1_REV04 = ENERGY_STAR_MFNC_V1_REV03;

// Rev.05 (v1.1/1.2/1.3, 01/15/2025) — significant restructure
const ENERGY_STAR_MFNC_V1_1_REV05 = [
  { id: "r5_1_5", pointNumber: "1.5", text: "Mass or metal-framed above-grade walls (including floor perimeter edges): continuous rigid insulation or insulated siding ≥ R-3 (CZ 1-4), ≥ R-5 (CZ 5-8)", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_1_6", pointNumber: "1.6", text: "Concrete slab floors in CZ 4-8 above ambient/garages/unconditioned spaces: floor insulation meets U-factor per 2009 IECC Table 502.1.2", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_1_7_2", pointNumber: "1.7.2", text: "Heated plenums: insulation at top meets Item 1.6 or exceeds mass floor R-value per 2009 IECC Table 502.2(1)", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_1_7_3", pointNumber: "1.7.3", text: "Heated plenums: bottom has at least R-13 insulation", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_1_8_1", pointNumber: "1.8.1", text: "Garages with space heating: walls insulated ≥ R-5ci (CZ 5-6), ≥ R-7.5ci (CZ 7), ≥ R-9.5ci (CZ 8)", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_1_8_2", pointNumber: "1.8.2", text: "Garages with space heating: ceiling insulation meets Item 1.6 or exceeds mass floor R-value per 2009 IECC", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_1_2", pointNumber: "1.2", text: "All insulation achieves Grade I installation per ANSI/RESNET/ICC 301", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_2_1", pointNumber: "2.1", text: "Air barrier fully aligned: dropped ceilings/soffits below unconditioned attics, chase/dead space, and all other ceilings", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_2_2", pointNumber: "2.2", text: "Air barrier fully aligned: walls behind showers, tubs, staircases, and fireplaces", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_2_3", pointNumber: "2.3", text: "Air barrier fully aligned: architectural bump-outs, dead space, and all other exterior walls", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_2_4", pointNumber: "2.4", text: "Air barrier fully aligned: floors above garages, floors above unconditioned spaces, and cantilevered floors", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_2_5", pointNumber: "2.5", text: "Air barrier fully aligned: all other floors adjoining unconditioned space (rim/band joists at exterior wall or porch roof)", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_3_4", pointNumber: "3.4", text: "Wood-framed above-grade walls assessed for advanced framing details (assessment only)", category: "High Performance Building Envelope", mandatory: false },
  { id: "r5_3_5", pointNumber: "3.5", text: "Above-grade and at-grade concrete floor edges (podiums, balconies, projected slabs) assessed for complete thermal break; in CZ 4-8, total building UA documented in Multifamily Workbook (assessment only)", category: "High Performance Building Envelope", mandatory: false },
  { id: "r5_3_6", pointNumber: "3.6", text: "Slabs on grade assessed for insulation where walls separate conditioned from unconditioned space (assessment only)", category: "High Performance Building Envelope", mandatory: false },
  { id: "r5_4_1", pointNumber: "4.1", text: "Ducts, flues, shafts, plumbing, piping, wiring, exhaust fans & other penetrations are sealed, with blocking/flashing as needed", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_4_2", pointNumber: "4.2", text: "Attic access panels, roof hatches and drop-down stairs are gasketed (not caulked) or equipped with gasketed covers", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_4_3", pointNumber: "4.3", text: "Recessed lighting fixtures are ICAT labeled and gasketed", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_4_4", pointNumber: "4.4", text: "Drywall is sealed to top plate during installation, or from attic side at all unconditioned attic/wall interfaces", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_4_5", pointNumber: "4.5", text: "Rough opening around windows & exterior doors is sealed", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_4_6", pointNumber: "4.6", text: "Assemblies separating attached garages from occupiable space are sealed; air barrier installed, sealed, and aligned", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_4_7", pointNumber: "4.7", text: "Doors adjacent to unconditioned space or ambient conditions made substantially air-tight with door seal and weatherstripping", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_4_8", pointNumber: "4.8", text: "Unit entrance doors from corridor/stairwell made substantially air-tight with door seal and weatherstripping", category: "High Performance Building Envelope", mandatory: true },
  { id: "r5_5_9", pointNumber: "5.9", text: "All heating and cooling systems serving a dwelling unit have thermostatic controls within the unit", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_5_10_2", pointNumber: "5.10.2", text: "All indoor/terminal units: system turns on and provides heat on call; turns off when heating setpoint met", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_5_10_3", pointNumber: "5.10.3", text: "All indoor/terminal units: system turns on and provides cooling on call; turns off when cooling setpoint met", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_5_11", pointNumber: "5.11", text: "Where present in CZ 4-8: stair and elevator shaft vents equipped with motorized dampers; verified closed at inspection", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_5_12", pointNumber: "5.12", text: "Garage heating, plenum heating, and freeze protection systems: automatic controls shut off above 40°F space or pipe wall temperatures", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_5_12_1", pointNumber: "5.12.1", text: "Heat tracing for freeze protection: controls based on pipe wall temperature and minimum R-3 pipe insulation", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_5_13", pointNumber: "5.13", text: "Snow/ice-melting systems: automatic controls to shut off when pavement above 50°F with no precipitation", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_5_14", pointNumber: "5.14", text: "Hydronic systems: terminal heating/cooling equipment separated from riser by control valve or terminal pump", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_5_15", pointNumber: "5.15", text: "Hydronic systems: terminal units equipped with pressure independent balancing valves or control valves", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_5_16", pointNumber: "5.16", text: "Hydronic systems: piping insulated per Item 4.42 of National HVAC Design Report including at planks and penetrations", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_5_17", pointNumber: "5.17", text: "Hydronic circulating pumps ≥1 HP, 3-phase: NEMA Premium motors; ≥5 HP also with variable frequency drives", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_6_1", pointNumber: "6.1", text: "Ductwork installed without kinks, sharp bends, compressions, or excessive coiled flexible ductwork", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_6_3", pointNumber: "6.3", text: "All supply and return ducts in unconditioned space insulated to ≥ R-6, including connections to trunk ducts", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_6_5", pointNumber: "6.5", text: "Common Space: supply, return, and exhaust ductwork and plenums sealed at all transverse joints, seams, and penetrations", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_7_4", pointNumber: "7.4", text: "Ventilation override control installed and labeled if function not obvious", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_7_5_1", pointNumber: "7.5.1", text: "Outdoor air inlet on ducted return: motorized damper automatically restricts airflow during vent off-cycle and occupant override", category: "Indoor Air Quality", mandatory: false },
  { id: "r5_7_6", pointNumber: "7.6", text: "Where OA inlets are connected to dwelling unit HVAC system: motorized damper installed that closes when no call for ventilation or fan is off", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_7_7", pointNumber: "7.7", text: "System fan in dwelling unit rated ≤ 3 sones if intermittent, ≤ 2 sones if continuous, or exempted", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_7_8", pointNumber: "7.8", text: "If Vent System controller operates HVAC fan: fan operation is intermittent and ECM/ICM type OR controls reduce runtime for HVAC hours", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_7_9", pointNumber: "7.9", text: "In-unit bathroom fans or in-line fans are ENERGY STAR certified if used as part of dwelling-unit mechanical ventilation system", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_7_11_1", pointNumber: "7.11.1", text: "Ventilation air inlets pull air directly from outdoors, not from attic, crawlspace, garage, or adjacent dwelling unit", category: "Indoor Air Quality", mandatory: false },
  { id: "r5_7_11_2", pointNumber: "7.11.2", text: "Inlets ≥2 ft above grade or roof deck; ≥10 ft from known contamination sources; ≥3 ft from dryer exhausts and roof sources", category: "Indoor Air Quality", mandatory: false },
  { id: "r5_7_11_3", pointNumber: "7.11.3", text: "Inlets provided with rodent/insect screen with ≤ 0.5 in. mesh", category: "Indoor Air Quality", mandatory: false },
  { id: "r5_8_2", pointNumber: "8.2", text: "Bathroom mechanical exhaust: continuous ≥ 20 CFM / ≤ 2 sones; intermittent ≥ 50 CFM; vented directly to outdoors", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_8_4", pointNumber: "8.4", text: "Shared garage exhaust system equipped with controls that sense CO and NO2", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_9_1", pointNumber: "9.1", text: "MERV 6+ filter(s) in each ducted mechanical system serving a dwelling unit; all return and outdoor air passes through filter", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_9_1_1", pointNumber: "9.1.1", text: "Filter access panel includes gasket and fits snugly against exposed edge of filter when closed to prevent bypass", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_10_1", pointNumber: "10.1", text: "Furnaces, boilers, and water heaters within pressure boundary are mechanically drafted or direct-vented", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_10_2", pointNumber: "10.2", text: "Fireplaces within pressure boundary are direct-vented", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_10_3", pointNumber: "10.3", text: "No unvented combustion appliances other than cooking ranges or ovens inside pressure boundary", category: "Indoor Air Quality", mandatory: true },
  { id: "r5_12_1_1", pointNumber: "12.1.1", text: "ERI Path: All common spaces (except lobby, mechanical rooms, safety exceptions) have occupancy/vacancy sensors, programmed timers, or automatic bi-level lighting controls; operation verified", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_12_2", pointNumber: "12.2", text: "Exterior lighting: automatic switching on timers or photocell controls, except 24-hour, security, or individual unit meter fixtures", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_12_3", pointNumber: "12.3", text: "Common spaces, exterior, and garages: 90% of installed lighting fixtures are integrated LED or contain LED lamps", category: "Energy Efficient Systems", mandatory: true },
  { id: "r5_13_2", pointNumber: "13.2", text: "ERI Path: Common space refrigerators and dishwashers are ENERGY STAR certified; showerheads are WaterSense labeled", category: "Energy Efficient Systems", mandatory: true },
];

// ─── EARTHCRAFT MFNC V6 ───────────────────────────────────────────────────────
// Tier field: "ALL"// ─── EARTHCRAFT MFNC V6 — NEW CONSTRUCTION ──────────────────────────────────────
// Source: ECMF-Workbook-V6.xlsx (New Construction)
// ALL = required at Certified + Gold; GOLD = required at Gold only
// ec_nc_du2_5 is the NC version of DU 2.5 (vs renovation ec_du2_4)

const EARTHCRAFT_CERTIFIED_V6 = [
  // ── SITE PLANNING ──────────────────────────────────────────────────────────
    { id: "ec_sp2_7", pointNumber: "SP 2.7", tier: "ALL", text: "Outdoor community gathering space provided on site", category: "Site Planning" },
  // ── RESOURCE EFFICIENCY ─────────────────────────────────────────────────────
  { id: "ec_v7_re1_0",   pointNumber: "RE 1.0",  tier: "ALL", text: "Limit framing at all windows and doors", category: "Resource Efficiency" },
  { id: "ec_v7_re1_1",   pointNumber: "RE 1.1",  tier: "ALL", text: "Engineered roof framing (90%)", category: "Resource Efficiency" },
  // ── DURABILITY & MOISTURE ───────────────────────────────────────────────────
  { id: "ec_v7_du1_6",   pointNumber: "DU 1.6",  tier: "ALL", text: "Continuous foundation termite flashing (required if slab edge is insulated)", category: "Durability & Moisture Management" },
  { id: "ec_du1_5",      pointNumber: "DU 1.7",  tier: "ALL", text: "Maintain 2\" clearance between wall siding and roof surface", category: "Durability & Moisture Management" },
  { id: "ec_du1_6",      pointNumber: "DU 1.8",  tier: "ALL", text: "Install level air conditioner condensing unit pad", category: "Durability & Moisture Management" },
  { id: "ec_du1_7",      pointNumber: "DU 1.9",  tier: "ALL", text: "Roof drip edge with ≥ 1/4\" overhang", category: "Durability & Moisture Management" },
  { id: "ec_v7_du1_10",  pointNumber: "DU 1.10", tier: "ALL", text: "Drain pan installed for all water heaters and washing machines", category: "Durability & Moisture Management" },
  { id: "ec_nc_du2_5",   pointNumber: "DU 2.5",  tier: "ALL", text: "Do not install wet or water-damaged building materials", category: "Durability & Moisture Management" },
  { id: "ec_v7_du2_6",   pointNumber: "DU 2.6",  tier: "ALL", text: "Capillary break between foundation and framing at all exterior walls", category: "Durability & Moisture Management" },
  { id: "ec_v7_du2_7",   pointNumber: "DU 2.7",  tier: "ALL", text: "Drainage board and damp proofing installed for all below-grade walls", category: "Durability & Moisture Management" },
  { id: "ec_du2_8",      pointNumber: "DU 2.8",  tier: "ALL", text: "Design for additional dehumidification: rough-in electrical and plumbing for dehumidifier", category: "Durability & Moisture Management" },
  // ── INDOOR AIR QUALITY ──────────────────────────────────────────────────────
  { id: "ec_iaq1",       pointNumber: "IAQ 1.0", tier: "ALL", text: "No unvented combustion fireplaces, appliances, or space heaters; all combustion appliances mechanically drafted or direct-vented", category: "Indoor Air Quality", mergedWith: ["es_10_1","es_10_2","es_10_3"] },
  { id: "ec_v7_iaq1_1",  pointNumber: "IAQ 1.1", tier: "ALL", text: "All fireplaces have outdoor combustion air supply; masonry-built fireplaces have gasketed doors", category: "Indoor Air Quality" },
  { id: "ec_iaq1_2",     pointNumber: "IAQ 1.2", tier: "ALL", text: "No atmospherically vented water heaters or furnaces", category: "Indoor Air Quality" },
  { id: "ec_iaq1_3",     pointNumber: "IAQ 1.3", tier: "ALL", text: "Sealed-combustion or electric water heater installed within conditioned space", category: "Indoor Air Quality" },
  { id: "ec_iaq1_4",     pointNumber: "IAQ 1.4", tier: "ALL", text: "Carbon monoxide detector installed if combustion appliances exist (one per unit)", category: "Indoor Air Quality" },
  { id: "ec_iaq2",       pointNumber: "IAQ 2.0", tier: "ALL", text: "Protect all ducts and indoor coils until floor/wall finishing is complete", category: "Indoor Air Quality" },
  { id: "ec_iaq2_1",     pointNumber: "IAQ 2.1", tier: "ALL", text: "Filter(s) easily accessible for property maintenance; MERV 6+ minimum; all return and outdoor air passes through filter", category: "Indoor Air Quality", mergedWith: ["es_9_1"] },
  { id: "ec_v7_iaq2_2",  pointNumber: "IAQ 2.2", tier: "ALL", text: "Rodent and corrosion-proof screens with mesh ≤0.5\" on all openings not fully sealed or caulked", category: "Indoor Air Quality" },
  { id: "ec_v7_iaq2_3",  pointNumber: "IAQ 2.3", tier: "ALL", text: "All outdoor supply air crosses a filter prior to distribution", category: "Indoor Air Quality" },
  { id: "ec_iaq2_5",     pointNumber: "IAQ 2.5", tier: "ALL", text: "No carpet in below-grade units", category: "Indoor Air Quality" },
  // ── BUILDING ENVELOPE ───────────────────────────────────────────────────────
  { id: "ec_v7_be3_9",   pointNumber: "BE 3.9",  tier: "ALL", text: "Slab edge insulation ≥ R-10", category: "High Performance Building Envelope" },
  // ── ENERGY SYSTEMS: WATER HEATING ───────────────────────────────────────────
  { id: "ec_es5_1",      pointNumber: "ES 5.1",  tier: "ALL", text: "Heat trap on all storage water heaters; confirm by visual inspection or AHRI certificate", category: "Water Efficiency", mergedWith: ["es_11_3"] },
  { id: "ec_es5_3",      pointNumber: "ES 5.3",  tier: "ALL", text: "Pipe insulation on first 2' of hot and cold water pipes at water heater", category: "Water Efficiency" },
  // ── WATER EFFICIENCY ────────────────────────────────────────────────────────
  { id: "ec_we1_0",      pointNumber: "WE 1.0",  tier: "ALL", text: "Meet National Energy Policy Act low-flow standards for all fixtures", category: "Water Efficiency" },
  { id: "ec_we1_1",      pointNumber: "WE 1.1",  tier: "ALL", text: "Detect and repair all leaks at water-using fixtures, appliances, and equipment", category: "Water Efficiency" },
  { id: "ec_we1_2",      pointNumber: "WE 1.2",  tier: "ALL", text: "Low-flow fixtures: WaterSense toilet ≤1.28 gpf; urinal ≤0.5 gpf; lavatory faucet ≤1.5 gpm; showerhead ≤2.0 gpm", category: "Water Efficiency", mergedWith: ["es_13_2"] },
  { id: "ec_du2_6",      pointNumber: "WE 2.3",  tier: "ALL", text: "Newly installed and existing plants maintain distance ≥2' from building at maturity", category: "Durability & Moisture Management" },
];

const EARTHCRAFT_GOLD_V6 = [
  ...EARTHCRAFT_CERTIFIED_V6,
  // ── RESOURCE EFFICIENCY: GOLD ───────────────────────────────────────────────
  { id: "ec_v7_re1_2",   pointNumber: "RE 1.2",  tier: "GOLD", text: "Advanced framing: 2-stud corners where structurally feasible; ladder T-walls; headers sized for actual loads", category: "Resource Efficiency" },
  // ── DURABILITY & MOISTURE: GOLD ─────────────────────────────────────────────
  { id: "ec_v7_du2_9",   pointNumber: "DU 2.9",  tier: "GOLD", text: "Additional dehumidification system installed: basement or sealed crawlspace system", category: "Durability & Moisture Management" },
  { id: "ec_du2_7",      pointNumber: "DU 2.10", tier: "GOLD", text: "Foundation drain at outside perimeter edge of footing surrounded with 6\" clean gravel and filter fabric", category: "Durability & Moisture Management" },
  // ── INDOOR AIR QUALITY: GOLD ────────────────────────────────────────────────
  { id: "ec_v7_iaq1_5",  pointNumber: "IAQ 1.5", tier: "GOLD", text: "If installed, all fireplaces meet indoor air quality guidelines and have gasketed doors", category: "Indoor Air Quality" },
  { id: "ec_iaq2_6",     pointNumber: "IAQ 2.6", tier: "GOLD", text: "Filters are ≥ MERV 8", category: "Indoor Air Quality" },
  // ── BUILDING ENVELOPE: GOLD ─────────────────────────────────────────────────
  { id: "ec_v7_be3_10",  pointNumber: "BE 3.10", tier: "GOLD", text: "Insulation quality: Grade I throughout OR Grade II with continuous insulated sheathing ≥ R-3 (100% coverage)", category: "High Performance Building Envelope" },
  { id: "ec_v7_be3_11",  pointNumber: "BE 3.11", tier: "GOLD", text: "Corners insulated to ≥ R-6", category: "High Performance Building Envelope" },
  { id: "ec_v7_be3_12",  pointNumber: "BE 3.12", tier: "GOLD", text: "Headers insulated to ≥ R-3", category: "High Performance Building Envelope" },
  { id: "ec_v7_be3_13",  pointNumber: "BE 3.13", tier: "GOLD", text: "Fiberglass batts are unfaced and friction-fit throughout", category: "High Performance Building Envelope" },
];

const EARTHCRAFT_SF2024_CERTIFIED = [
  // ── RESOURCE EFFICIENCY ──────────────────────────────────────────────────────
  { id: "ec_v7_re1_0",    pointNumber: "RE 1.0",  tier: "ALL", text: "Limit framing at all windows and doors", category: "Resource Efficiency" },
  // ── DURABILITY & MOISTURE MANAGEMENT ────────────────────────────────────────
  { id: "ec_sf_du2_9",    pointNumber: "DU 2.9",  tier: "ALL", text: "Vapor barriers installed only under slab(s) and/or in crawlspace(s); not on vertical walls", category: "Durability & Moisture Management" },
  // ── HIGH PERFORMANCE BUILDING ENVELOPE: AIR SEALING ────────────────────────
  { id: "ec_sf_be2_0",    pointNumber: "BE 2.0",  tier: "ALL", text: "Seal bottom plates to subfloor or foundation for entire unit envelope", category: "High Performance Building Envelope" },
  { id: "ec_sf_be2_1",    pointNumber: "BE 2.1",  tier: "ALL", text: "Block and seal joist cavities at: above attached garage walls; above supporting walls at cantilevered floors; under attic knee walls; between units and corridors", category: "High Performance Building Envelope" },
  { id: "ec_sf_be2_2",    pointNumber: "BE 2.2",  tier: "ALL", text: "Block stud cavities at all changes in ceiling height", category: "High Performance Building Envelope" },
  { id: "ec_sf_be2_3",    pointNumber: "BE 2.3",  tier: "ALL", text: "Install blocking and baffles at all insulated and vented vaulted ceilings", category: "High Performance Building Envelope" },
  { id: "ec_sf_be2_6",    pointNumber: "BE 2.6",  tier: "ALL", text: "Install rigid air barriers: behind tubs and showers on insulated walls; at attic knee wall on attic-side including skylight shafts; at chases in contact with the building envelope; along staircases on insulated walls; along porch roofs", category: "High Performance Building Envelope" },
  { id: "ec_sf_be2_7",    pointNumber: "BE 2.7",  tier: "ALL", text: "Install weatherstripping at: all exterior doors; attic knee wall doors, scuttle holes, and pull-down stairs located within conditioned space", category: "High Performance Building Envelope" },
  { id: "es_4_2",         pointNumber: "BE 2.8",  tier: "ALL", text: "All recessed can lights are air tight and gasketed; IC-rated in insulated ceilings", category: "High Performance Building Envelope" },
  { id: "ec_sf_be2_9",    pointNumber: "BE 2.9",  tier: "ALL", text: "Gypcrete installed on all framed floors separating unit envelopes", category: "High Performance Building Envelope" },
  // ── HIGH PERFORMANCE BUILDING ENVELOPE: INSULATION ──────────────────────────
  { id: "ec_sf_be4_3",    pointNumber: "BE 4.3",  tier: "ALL", text: "Exterior wall insulation: walls and band joists ≥ R-13; fireplace chases on exterior walls ≥ R-13; foundation walls (CZ 1/2/3 ≥ R-5 continuous or ≥ R-13 cavity; CZ 4/5 ≥ R-10 continuous or ≥ R-13 cavity)", category: "High Performance Building Envelope" },
  { id: "ec_sf_be4_4",    pointNumber: "BE 4.4",  tier: "ALL", text: "Ceiling insulation: unconditioned attics (CZ 1/2/3 ≥ R-38; CZ 4/5 ≥ R-49); wind baffles at eaves in every vented bay; attic platforms allow full-depth insulation below; depth rulers installed for loose-fill attic insulation", category: "High Performance Building Envelope" },
  { id: "ec_sf_be4_10",   pointNumber: "BE 4.10", tier: "ALL", text: "Slab edge insulation: Climate Zone 2/3 ≥ R-5; Climate Zone 4/5 ≥ R-10", category: "High Performance Building Envelope" },
  { id: "ec_sf_be5_3",    pointNumber: "BE 5.3",  tier: "ALL", text: "NFRC certification label present on all installed doors, windows, and skylights", category: "High Performance Building Envelope" },
  // ── ENERGY EFFICIENT SYSTEMS: HVAC EQUIPMENT ────────────────────────────────
  { id: "ec_sf_es1_1",    pointNumber: "ES 1.1",  tier: "ALL", text: "If programmable thermostat installed for heat pump, verify it includes adaptive recovery technology", category: "Energy Efficient Systems" },
  // ── ENERGY EFFICIENT SYSTEMS: HVAC DUCT SYSTEM ──────────────────────────────
  { id: "ec_iaq2",        pointNumber: "ES 2.7",  tier: "ALL", text: "Indoor coil protected until indoor finishes are complete (drywall, paint)", category: "Indoor Air Quality" },
  // ── ENERGY EFFICIENT SYSTEMS: FILTERS ───────────────────────────────────────
  { id: "ec_sf_es1_8",    pointNumber: "ES 1.8",  tier: "ALL", text: "Filters are ≥ MERV 8", category: "Indoor Air Quality" },
  // ── ENERGY EFFICIENT SYSTEMS: VENTILATION ───────────────────────────────────
  { id: "ec_sf_es4_2",    pointNumber: "ES 4.2",  tier: "ALL", text: "Seal seams of all intake and exhaust ducts with mastic", category: "Indoor Air Quality" },
  { id: "ec_sf_es4_3",    pointNumber: "ES 4.3",  tier: "ALL", text: "ASHRAE compliant exhaust fans rated ≥ 50 cfm installed in all bathrooms and ducted to outside", category: "Indoor Air Quality" },
  { id: "ec_sf_es4_4",    pointNumber: "ES 4.4",  tier: "ALL", text: "Gas kitchen range and/or cooktop vented to exterior with ASHRAE compliant ≥ 100 cfm fan", category: "Indoor Air Quality" },
  { id: "ec_sf_es4_5",    pointNumber: "ES 4.5",  tier: "ALL", text: "Back-draft dampers installed for kitchen and bathroom exhaust", category: "Indoor Air Quality" },
  { id: "ec_sf_es4_6",    pointNumber: "ES 4.6",  tier: "ALL", text: "ENERGY STAR certified bathroom exhaust fans", category: "Indoor Air Quality" },
  { id: "ec_es4_4",       pointNumber: "ES 4.7",  tier: "ALL", text: "Duct clothes dryers to outside", category: "Indoor Air Quality" },
  // ── ENERGY EFFICIENT SYSTEMS: APPLIANCES ────────────────────────────────────
  { id: "ec_sf_es6_3",    pointNumber: "ES 6.3",  tier: "ALL", text: "If installed, ENERGY STAR clothes washer and dryer kit (in residential units and/or communal laundry facility)", category: "Energy Efficient Systems" },
  // ── DURABILITY & MOISTURE MANAGEMENT ────────────────────────────────────────
  { id: "ec_du2_6",       pointNumber: "WE 2.3",  tier: "ALL", text: "Plants installed to maintain distance ≥ 2' from building at maturity", category: "Durability & Moisture Management" },


  // ── SITE PLANNING ─────────────────────────────────────────────────────────────
  { id: "ec_sf_sp2_7",   pointNumber: "SP 2.7",  tier: "ALL", text: "Pervious paving used for hardscapes and surface parking areas", category: "Site Planning" },
  // ── RESOURCE EFFICIENCY (ADDITIONAL) ───────────────────────────────────────
  { id: "ec_sf_re1_4",   pointNumber: "RE 1.4",  tier: "ALL", text: "Floor joists are 24\" on center (≥80%)", category: "Resource Efficiency" },
  { id: "ec_sf_re1_5",   pointNumber: "RE 1.5",  tier: "ALL", text: "Non-load bearing wall studs are 24\" on center", category: "Resource Efficiency" },
  { id: "ec_sf_re2_5",   pointNumber: "RE 2.5",  tier: "ALL", text: "Structural headers are steel or engineered wood (≥90%)", category: "Resource Efficiency" },
  // ── DURABILITY & MOISTURE MANAGEMENT (ADDITIONAL) ──────────────────────────
  { id: "ec_sf_du1_17",  pointNumber: "DU 1.17", tier: "ALL", text: "Termite mesh system installed", category: "Durability & Moisture Management" },
  { id: "ec_sf_du1_20",  pointNumber: "DU 1.20", tier: "ALL", text: "All entrance doors have overhang with ≥3' depth", category: "Durability & Moisture Management" },
  { id: "ec_sf_du2_14",  pointNumber: "DU 2.14", tier: "ALL", text: "Humidistat or thermidistat installed with whole-house variable speed cooling system", category: "Durability & Moisture Management" },
  // ── HIGH PERFORMANCE BUILDING ENVELOPE (ADDITIONAL) ────────────────────────
  { id: "ec_sf_be4_19",  pointNumber: "BE 4.19", tier: "ALL", text: "Basement wall insulated", category: "High Performance Building Envelope" },
  { id: "ec_sf_be4_20",  pointNumber: "BE 4.20", tier: "ALL", text: "Attic knee wall insulated to ≥ R-22 with continuous insulated air barrier on attic side", category: "High Performance Building Envelope" },
  // ── ENERGY EFFICIENT SYSTEMS (ADDITIONAL) ──────────────────────────────────
  { id: "ec_sf_es1_13",  pointNumber: "ES 1.13", tier: "ALL", text: "Condenser units are spaced at least 2 feet apart", category: "Energy Efficient Systems" },
  { id: "ec_sf_es1_16",  pointNumber: "ES 1.16", tier: "ALL", text: "Zone control system installed", category: "Energy Efficient Systems" },
  // ── INDOOR AIR QUALITY (ADDITIONAL) ────────────────────────────────────────
  { id: "ec_sf_iaq2_10", pointNumber: "IAQ 2.10",tier: "ALL", text: "No carpet installed in any unit (all floors)", category: "Indoor Air Quality" },
  { id: "ec_sf_iaq2_12", pointNumber: "IAQ 2.12",tier: "ALL", text: "Permanent walk-off mats installed at each building entry", category: "Indoor Air Quality" },
  { id: "ec_sf_es4_19",  pointNumber: "ES 4.19", tier: "ALL", text: "Bathroom exhaust fans rated ≤1 sone", category: "Indoor Air Quality" },
  { id: "ec_sf_es4_20",  pointNumber: "ES 4.20", tier: "ALL", text: "Bathroom exhaust fans have automatic controls with humidistat or timer", category: "Indoor Air Quality" },
  { id: "ec_sf_es4_21",  pointNumber: "ES 4.21", tier: "ALL", text: "Energy recovery ventilator (ERV) installed for whole-unit ventilation strategy", category: "Indoor Air Quality" },
  { id: "ec_sf_es4_23",  pointNumber: "ES 4.23", tier: "ALL", text: "Storage rooms vented to outside", category: "Indoor Air Quality" },
  // ── WATER EFFICIENCY (ADDITIONAL) ──────────────────────────────────────────
  { id: "ec_sf_we1_14",  pointNumber: "WE 1.14", tier: "ALL", text: "Leak detection sensors installed at kitchens, bathrooms, and laundry in all residential units", category: "Water Efficiency" },
];

const EARTHCRAFT_SF2024_GOLD = [
  ...EARTHCRAFT_SF2024_CERTIFIED,
  // ── RESOURCE EFFICIENCY: GOLD ────────────────────────────────────────────────
  { id: "ec_v7_re1_2",    pointNumber: "RE 1.2",  tier: "GOLD", text: "Advanced framing: 2-stud corners where structurally feasible; ladder T-walls; headers sized for actual loads", category: "Resource Efficiency" },
  // ── DURABILITY & MOISTURE MANAGEMENT: GOLD ──────────────────────────────────
  { id: "ec_du2_7",       pointNumber: "DU 2.10", tier: "GOLD", text: "Foundation drain at outside perimeter edge of footing surrounded with 6\" clean gravel and fabric filter", category: "Durability & Moisture Management" },
  { id: "ec_sf_du2_11",   pointNumber: "DU 2.11", tier: "GOLD", text: "Dedicated dehumidification system in basement and/or closed crawlspace areas", category: "Durability & Moisture Management" },
  { id: "ec_sf_du2_12",   pointNumber: "DU 2.12", tier: "GOLD", text: "Design for or install additional dehumidification: rough-in electrical and plumbing for whole-unit dehumidifier OR install whole-unit ENERGY STAR dehumidifier with pump and drain to outdoors", category: "Durability & Moisture Management" },
  // ── HIGH PERFORMANCE BUILDING ENVELOPE: GOLD ────────────────────────────────
  { id: "ec_v7_be3_11",   pointNumber: "BE 4.11", tier: "GOLD", text: "Corners insulated to ≥ R-6", category: "High Performance Building Envelope" },
  { id: "ec_sf_be4_12",   pointNumber: "BE 4.12", tier: "GOLD", text: "Headers insulated to ≥ R-5", category: "High Performance Building Envelope" },
  { id: "ec_sf_be4_13",   pointNumber: "BE 4.13", tier: "GOLD", text: "Unconditioned attic: energy heel trusses or raised top plate installed", category: "High Performance Building Envelope" },
  // ── ENERGY EFFICIENT SYSTEMS: GOLD ──────────────────────────────────────────
  { id: "ec_sf_es2_11",   pointNumber: "ES 2.11", tier: "GOLD", text: "Fully duct all returns", category: "Energy Efficient Systems" },
  { id: "ec_sf_es2_12",   pointNumber: "ES 2.12", tier: "GOLD", text: "Install rigid ductwork or pull all flex ducts with no pinches, supported at intervals ≤ 4'", category: "Energy Efficient Systems" },
  { id: "ec_sf_es4_9",    pointNumber: "ES 4.9",  tier: "GOLD", text: "ENERGY STAR qualified ceiling fans installed", category: "Energy Efficient Systems" },
];


const MRF_ITEMS = [
  // ── HVAC & MECHANICAL ────────────────────────────────────────────────────────
  { id: "mrf_1_0", pointNumber: "HVAC Equipment", tier: "ALL", category: "Minimum Rated Features",
    text: "Nameplate on indoor and outdoor units. Capture make, model numbers, SEER2, and HSPF2. Document all unique configurations by unit type." },
  { id: "mrf_1_1", pointNumber: "Thermostat", tier: "ALL", category: "Minimum Rated Features",
    text: "Identify type: Basic / Programmable / Smart. For heat pump systems, confirm adaptive recovery feature is present." },
  { id: "mrf_1_2", pointNumber: "Mechanical Ventilation", tier: "ALL", category: "Minimum Rated Features",
    text: "Unit label showing system type, cfm rate, hours/day of operation, and fan watts. Document each unique configuration by unit type." },
  { id: "mrf_1_3", pointNumber: "Dehumidifier", tier: "ALL", category: "Minimum Rated Features",
    text: "If installed, photograph nameplate. Capture model number and note location (in-unit, basement, or crawlspace)." },

  // ── BUILDING ENVELOPE ────────────────────────────────────────────────────────
  { id: "mrf_2_0", pointNumber: "Wall Insulation", tier: "ALL", category: "Minimum Rated Features",
    text: "Insulation in framing bays. Capture cavity R-value, continuous R-value, grade (I/II), insulation types, framing material, spacing, and depth. Document all unique wall assemblies." },
  { id: "mrf_2_1", pointNumber: "Ceiling Insulation", tier: "ALL", category: "Minimum Rated Features",
    text: "Attic insulation with depth ruler visible. Capture cavity R-value, continuous R-value, grade, insulation types, framing spacing, and depth. Document all unique roof assemblies." },
  { id: "mrf_2_2", pointNumber: "Foundation Insulation", tier: "ALL", category: "Minimum Rated Features",
    text: "Foundation insulation showing R-value, grade, and type. Confirm alignment with thermal boundary." },
  { id: "mrf_2_3", pointNumber: "Rim & Band Insulation", tier: "ALL", category: "Minimum Rated Features",
    text: "Rim and band joist insulation. Capture R-value, grade, and insulation type." },
  { id: "mrf_2_4", pointNumber: "Duct Insulation", tier: "ALL", category: "Minimum Rated Features",
    text: "Ductwork insulation label. Capture R-value and note whether located in conditioned or unconditioned space." },
  { id: "mrf_2_5", pointNumber: "Windows", tier: "ALL", category: "Minimum Rated Features",
    text: "NFRC label on each unique window type. Capture U-factor and SHGC for all window configurations." },
  { id: "mrf_2_6", pointNumber: "Doors", tier: "ALL", category: "Minimum Rated Features",
    text: "NFRC label on each door type (main entry, patio, additional). Capture U-factor and SHGC." },
  { id: "mrf_2_7", pointNumber: "Roof Properties", tier: "ALL", category: "Minimum Rated Features",
    text: "Roof surface showing exterior color. Confirm presence or absence of radiant barrier." },

  // ── WATER HEATING & PLUMBING ─────────────────────────────────────────────────
  { id: "mrf_3_0", pointNumber: "Water Heater", tier: "ALL", category: "Minimum Rated Features",
    text: "Nameplate showing brand, model number, and location. If recirc system present, note pump wattage and control type." },
  { id: "mrf_3_1", pointNumber: "Hot Water Pipe Insulation", tier: "ALL", category: "Minimum Rated Features",
    text: "Supply pipes at water heater and throughout distribution. Confirm R-3 or better insulation on all hot water pipes." },
  { id: "mrf_3_2", pointNumber: "Water Fixtures", tier: "ALL", category: "Minimum Rated Features",
    text: "Flow rate markings on shower, bathroom faucet, and toilet. Capture shower (gpm), faucet (gpm), and toilet (gpf)." },

  // ── VENTILATION & EXHAUST ────────────────────────────────────────────────────
  { id: "mrf_4_0", pointNumber: "Bath Fans", tier: "ALL", category: "Minimum Rated Features",
    text: "Model label on each unique bath fan. Capture model number and sone rating." },
  { id: "mrf_4_1", pointNumber: "Kitchen Exhaust", tier: "ALL", category: "Minimum Rated Features",
    text: "Model label on each unique range hood or exhaust fan. Capture model number and sone rating." },

  // ── APPLIANCES ───────────────────────────────────────────────────────────────
  { id: "mrf_5_0", pointNumber: "Refrigerator", tier: "ALL", category: "Minimum Rated Features",
    text: "Model label on all unit types (standard, ADA, common area). Confirm ENERGY STAR certification." },
  { id: "mrf_5_1", pointNumber: "Dishwasher", tier: "ALL", category: "Minimum Rated Features",
    text: "Model label on all unit types. Confirm ENERGY STAR certification." },
  { id: "mrf_5_2", pointNumber: "Stove / Cooktop", tier: "ALL", category: "Minimum Rated Features",
    text: "Unit label or data plate. Note fuel source (electric / gas / propane) and whether induction or convection." },
  { id: "mrf_5_3", pointNumber: "Clothes Dryer", tier: "ALL", category: "Minimum Rated Features",
    text: "Model label. Capture model number, fuel source, and location. Note quantity if central installation." },
  { id: "mrf_5_4", pointNumber: "Clothes Washer", tier: "ALL", category: "Minimum Rated Features",
    text: "Model label. Capture model number and location. Note quantity if central installation." },
  { id: "mrf_5_5", pointNumber: "Ceiling Fan", tier: "ALL", category: "Minimum Rated Features",
    text: "Model label showing model number and cfm/watt rating." },

  // ── LIGHTING ─────────────────────────────────────────────────────────────────
  { id: "mrf_6_0", pointNumber: "Interior Lighting", tier: "ALL", category: "Minimum Rated Features",
    text: "Representative fixtures in units and common areas. Record percentage of LED, CFL, pin-based, and incandescent." },
  { id: "mrf_6_1", pointNumber: "Exterior Lighting", tier: "ALL", category: "Minimum Rated Features",
    text: "Exterior fixtures at building perimeter. Record percentage of LED at all exterior locations." },
];

// Repeatable structured entries for MRF envelope items — several assemblies can exist per item
// (e.g. an Interior wall entry and an Exterior wall entry on the same "Wall Insulation" item).
const MULTI_ENTRY_CONFIG = {
  mrf_2_0: { label: "Wall assemblies", entryLabel: "wall assembly", repeatable: true, fields: [
    { key: "wallType", type: "select", label: "Wall type", options: [["interior","Interior"],["exterior","Exterior"],["breezeway","Breezeway"]] },
    { key: "grade", type: "select", label: "Grade", options: [["GI","GI"],["GII","GII"],["GIII","GIII"]] },
    { key: "rValue", type: "text", label: "R-value" },
  ]},
  mrf_2_1: { label: "Ceiling assemblies", entryLabel: "ceiling assembly", fields: [
    { key: "location", type: "select", label: "Location", options: [["unconditioned_vented_attic","Unconditioned Vented Attic"],["sealed_attic","Sealed Attic"],["vaulted_roof","Vaulted Roof (Exposed Exterior)"]] },
    { key: "rValue", type: "text", label: "R-value" },
  ]},
  mrf_2_2: { label: "Foundation assemblies", entryLabel: "foundation assembly", fields: [
    { key: "rValue", type: "text", label: "R-value" },
    { key: "perimeterDepth", type: "text", label: "Perimeter insulation depth (ft)" },
    { key: "underslabDepth", type: "text", label: "Underslab insulation depth (ft)" },
  ]},
  mrf_2_3: { label: "Rim & band entries", entryLabel: "entry", fields: [
    { key: "rValue", type: "text", label: "R-value" },
  ]},
  mrf_2_4: { label: "Duct insulation entries", entryLabel: "entry", fields: [
    { key: "rValue", type: "text", label: "R-value" },
  ]},
  mrf_2_5: { label: "Window types", entryLabel: "window type", repeatable: true, fields: [
    { key: "uValue", type: "decimal", label: "U-Value" },
    { key: "shgc", type: "decimal", label: "SHGC" },
  ]},
  mrf_4_0: { label: "Bath fans", entryLabel: "bath fan", repeatable: true, fields: [
    { key: "modelNumber", type: "text", label: "Model number" },
    { key: "soneRating", type: "decimal", label: "Sone rating" },
  ]},
};

const CHECKLIST_REGISTRY = {
  "energy_star_mfnc||1 / 1.1 / 1.2||Rev. 03": ENERGY_STAR_MFNC_V1_REV03,
  "energy_star_mfnc||1 / 1.1 / 1.2||Rev. 04": ENERGY_STAR_MFNC_V1_REV04,
  "energy_star_mfnc||1.1 / 1.2 / 1.3||Rev. 05": ENERGY_STAR_MFNC_V1_1_REV05,
  "earthcraft_certified||V6||New Construction": EARTHCRAFT_CERTIFIED_V6,
  "earthcraft_gold||V6||New Construction": EARTHCRAFT_GOLD_V6,
  "earthcraft_certified||V7||New Construction": EARTHCRAFT_CERTIFIED_V7,
  "earthcraft_gold||V7||New Construction": EARTHCRAFT_GOLD_V7,
  "earthcraft_sf2024_certified||v2024||Southface": EARTHCRAFT_SF2024_CERTIFIED,
  "earthcraft_sf2024_gold||v2024||Southface":      EARTHCRAFT_SF2024_GOLD,
};

function getItemsForSelection(programSelections, categoryId) {
  const seen = new Set();
  const result = [];
  // MRF items are program-agnostic — always show when viewing MRF category
  if (categoryId === "Minimum Rated Features") {
    return MRF_ITEMS.map(i => ({ ...i, _cat: "Minimum Rated Features" }));
  }
  for (const sel of programSelections) {
    const key = `${sel.programId}||${sel.version}||${sel.revision}`;
    const items = (CHECKLIST_REGISTRY[key] || []).filter(i => i.category === categoryId);
    for (const item of items) {
      if (!seen.has(item.id)) { seen.add(item.id); result.push({ ...item, sourceKey: key }); }
    }
  }
  return result;
}

// ─── STORAGE ──────────────────────────────────────────────────────────────────
const STORAGE_KEY = "greencert_v2";
function loadData() {
  try { const r = localStorage.getItem(STORAGE_KEY); return r ? JSON.parse(r) : { projects: [], records: {} }; }
  catch { return { projects: [], records: {} }; }
}
function saveData(d) { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(d)); } catch {} }

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function calcCatProgress(items, records, projectId, categoryId) {
  if (!items.length) return { pct: 0, pass: 0, fail: 0, na: 0, total: 0 };
  let pass = 0, fail = 0, na = 0;
  items.forEach(item => {
    const r = records[`${projectId}__${categoryId}__${item.id}`];
    if (r?.status === "pass") pass++;
    else if (r?.status === "fail") fail++;
    else if (r?.status === "na") na++;
  });
  return { pct: Math.round(((pass + na) / items.length) * 100), pass, fail, na, total: items.length };
}

function calcProjectProgress(project, records) {
  let total = 0, verified = 0, fail = 0;
  CATEGORIES.forEach(cat => {
    const items = getItemsForSelection(project.programs || [], cat.id);
    items.forEach(item => {
      total++;
      const r = records[`${project.id}__${cat.id}__${item.id}`];
      if (r?.status === "pass" || r?.status === "na") verified++;
      if (r?.status === "fail") fail++;
    });
  });
  return { pct: total ? Math.round((verified / total) * 100) : 0, fail, total, verified };
}

function programLabel(sel) {
  const p = PROGRAM_CATALOG.find(x => x.id === sel.programId);
  return p ? `${p.label} v${sel.version} ${sel.revision}` : sel.programId;
}

function programColor(programId) {
  return PROGRAM_CATALOG.find(x => x.id === programId)?.color || "#6B7280";
}

// ─── UI ATOMS ─────────────────────────────────────────────────────────────────
function fmtDate(iso) {
  if (!iso) return null;
  const d = new Date(iso);
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }) +
    " " + d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
}

function ProgressRing({ pct, size = 56, stroke = 5, fail = 0 }) {
  const r = (size - stroke) / 2;
  const circ = 2 * Math.PI * r;
  const offset = circ - (pct / 100) * circ;
  const color = fail > 0 ? "#EF4444" : pct === 100 ? "#10B981" : "#3B82F6";
  return (
    <svg width={size} height={size} style={{ transform: "rotate(-90deg)", flexShrink: 0 }}>
      <circle cx={size/2} cy={size/2} r={r} fill="none" stroke="#E5E7EB" strokeWidth={stroke}/>
      <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={color} strokeWidth={stroke}
        strokeDasharray={circ} strokeDashoffset={offset} strokeLinecap="round"
        style={{ transition: "stroke-dashoffset 0.5s ease" }}/>
    </svg>
  );
}

function ProgressBar({ pct, fail }) {
  const color = fail > 0 ? "#EF4444" : pct === 100 ? "#10B981" : "#3B82F6";
  return (
    <div style={{ background: "#F3F4F6", borderRadius: 4, height: 6 }}>
      <div style={{ width: `${pct}%`, height: "100%", background: color, borderRadius: 4, transition: "width 0.4s ease" }}/>
    </div>
  );
}

function StatusBadge({ status }) {
  const m = { pass: ["#D1FAE5","#065F46","Pass"], fail: ["#FEE2E2","#991B1B","Fail"], na: ["#F3F4F6","#4B5563","N/A"] };
  const s = m[status]; if (!s) return null;
  return <span style={{ background: s[0], color: s[1], fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 20 }}>{s[2]}</span>;
}

// ─── SCREEN: PROJECT LIST ─────────────────────────────────────────────────────
function ProjectList({ projects, records, onSelect, onCreate, onDelete, auth, onLogout }) {
  const [confirmId, setConfirmId] = useState(null);
  const confirmProj = projects.find(p => p.id === confirmId);

  return (
    <div style={{ paddingBottom: 80 }}>

      {/* SharePoint connection bar */}
      <div style={{ padding: "12px 20px", background: auth ? "#F0FDF4" : "#F9FAFB", borderBottom: "1px solid #E5E7EB", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
        {auth ? (
          <>
            <div style={{ minWidth: 0 }}>
              <p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: "#059669" }}>☁ Connected to SharePoint</p>
              <p style={{ margin: "2px 0 0", fontSize: 11, color: "#6B7280", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{auth.user?.name || auth.user?.email}</p>
            </div>
            <button onClick={onLogout}
              style={{ fontSize: 11, color: "#9CA3AF", background: "none", border: "1px solid #E5E7EB", borderRadius: 6, padding: "4px 10px", cursor: "pointer", flexShrink: 0, fontFamily: "DM Sans, sans-serif" }}>
              Disconnect
            </button>
          </>
        ) : (
          <>
            <p style={{ margin: 0, fontSize: 12, color: "#9CA3AF" }}>SharePoint not connected</p>
            <button onClick={startLogin}
              style={{ fontSize: 12, fontWeight: 600, color: "#FFF", background: "#0078D4", border: "none", borderRadius: 6, padding: "6px 14px", cursor: "pointer", flexShrink: 0, fontFamily: "DM Sans, sans-serif" }}>
              Connect
            </button>
          </>
        )}
      </div>
      <div style={{ padding: "20px 20px 12px", borderBottom: "1px solid #F3F4F6" }}>
        <p style={{ margin: 0, fontSize: 11, fontWeight: 700, color: "#9CA3AF", letterSpacing: "0.08em", textTransform: "uppercase" }}>Active projects</p>
      </div>
      {!projects.length && (
        <div style={{ padding: "48px 20px", textAlign: "center" }}>
          <div style={{ fontSize: 36, marginBottom: 12 }}>☑️</div>
          <p style={{ margin: 0, fontSize: 15, color: "#6B7280" }}>No projects yet</p>
          <p style={{ margin: "4px 0 0", fontSize: 13, color: "#9CA3AF" }}>Tap + to get started</p>
        </div>
      )}
      {projects.map(proj => {
        const pg = calcProjectProgress(proj, records);
        return (
          <div key={proj.id}
            style={{ padding: "14px 20px", borderBottom: "1px solid #F9FAFB", display: "flex", alignItems: "center", gap: 14, background: "#FFF" }}>
            <div onClick={() => onSelect(proj)} style={{ position: "relative", flexShrink: 0, cursor: "pointer" }}>
              <ProgressRing pct={pg.pct} fail={pg.fail}/>
              <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700, color: pg.fail>0?"#EF4444":pg.pct===100?"#10B981":"#3B82F6" }}>{pg.pct}%</div>
            </div>
            <div onClick={() => onSelect(proj)} style={{ flex: 1, minWidth: 0, cursor: "pointer" }}>
              <p style={{ margin: 0, fontSize: 15, fontWeight: 600, color: "#111827", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{proj.name}</p>
              <p style={{ margin: "3px 0 0", fontSize: 12, color: "#9CA3AF" }}>
                {(proj.programs||[]).length} program{proj.programs?.length!==1?"s":""} · {pg.verified}/{pg.total} items
                {pg.fail>0 && <span style={{ color: "#EF4444", fontWeight: 600 }}> · {pg.fail} fail{pg.fail>1?"s":""}</span>}
              </p>
            </div>
            <button onClick={() => setConfirmId(proj.id)}
              style={{ background: "none", border: "none", cursor: "pointer", padding: "4px 6px", color: "#D1D5DB", fontSize: 18, flexShrink: 0, lineHeight: 1 }}
              title="Delete project">🗑</button>
          </div>
        );
      })}

      {/* Delete confirmation sheet */}
      {confirmId && confirmProj && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 100, display: "flex", alignItems: "flex-end" }}
          onClick={() => setConfirmId(null)}>
          <div onClick={e => e.stopPropagation()}
            style={{ width: "100%", maxWidth: 430, margin: "0 auto", background: "#FFF", borderRadius: "16px 16px 0 0", padding: "24px 20px 36px" }}>
            <p style={{ margin: "0 0 6px", fontSize: 16, fontWeight: 700, color: "#111827" }}>Delete "{confirmProj.name}"?</p>
            <p style={{ margin: "0 0 24px", fontSize: 13, color: "#6B7280" }}>This will permanently delete the project and all its inspection records. This can't be undone.</p>
            <button onClick={() => { onDelete(confirmId); setConfirmId(null); }}
              style={{ width: "100%", padding: 14, background: "#EF4444", border: "none", borderRadius: 12, color: "#FFF", fontSize: 15, fontWeight: 600, cursor: "pointer", fontFamily: "DM Sans, sans-serif", marginBottom: 10 }}>
              Delete project
            </button>
            <button onClick={() => setConfirmId(null)}
              style={{ width: "100%", padding: 14, background: "none", border: "1.5px solid #E5E7EB", borderRadius: 12, color: "#374151", fontSize: 15, fontWeight: 500, cursor: "pointer", fontFamily: "DM Sans, sans-serif" }}>
              Cancel
            </button>
          </div>
        </div>
      )}

      <button onClick={onCreate}
        style={{ position: "fixed", bottom: 28, right: 24, width: 56, height: 56, borderRadius: "50%", background: "#1B4332", border: "none", color: "#FFF", fontSize: 28, cursor: "pointer", boxShadow: "0 4px 16px rgba(27,67,50,.35)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}>+</button>
    </div>
  );
}

// ─── SCREEN: CREATE PROJECT ───────────────────────────────────────────────────
function CreateProject({ onSave, onBack }) {
  const [name, setName] = useState("");
  const [advisor, setAdvisor] = useState("");
  const [step, setStep] = useState("name"); // name | programs | version
  const [selections, setSelections] = useState([]); // [{programId, version, revision}]
  const [pickingProgram, setPickingProgram] = useState(null); // programId being configured
  const [pickVersion, setPickVersion] = useState(null);

  const startAddProgram = () => setPickingProgram("choose");

  const confirmVersionRevision = (programId, version, revision) => {
    setSelections(s => [...s, { programId, version, revision }]);
    setPickingProgram(null); setPickVersion(null);
  };

  const removeSelection = (idx) => setSelections(s => s.filter((_, i) => i !== idx));

  if (step === "name") {
    return (
      <div style={{ padding: "24px 20px" }}>
        <h2 style={{ margin: "0 0 24px", fontSize: 20, fontWeight: 700, color: "#111827" }}>New project</h2>
        <label style={{ fontSize: 12, fontWeight: 700, color: "#6B7280", letterSpacing: "0.06em", textTransform: "uppercase" }}>Project name</label>
        <input value={name} onChange={e => setName(e.target.value)}
          placeholder="e.g. Green Park"
          style={{ display: "block", width: "100%", marginTop: 8, padding: "12px 14px", fontSize: 16, border: "1.5px solid #E5E7EB", borderRadius: 10, outline: "none", boxSizing: "border-box", fontFamily: "DM Sans, sans-serif" }}/>
        <label style={{ display: "block", marginTop: 20, fontSize: 12, fontWeight: 700, color: "#6B7280", letterSpacing: "0.06em", textTransform: "uppercase" }}>Technical Advisor</label>
        <input value={advisor} onChange={e => setAdvisor(e.target.value)}
          placeholder="Full name"
          style={{ display: "block", width: "100%", marginTop: 8, padding: "12px 14px", fontSize: 16, border: "1.5px solid #E5E7EB", borderRadius: 10, outline: "none", boxSizing: "border-box", fontFamily: "DM Sans, sans-serif" }}/>
        <button onClick={() => name.trim() && setStep("programs")} disabled={!name.trim()}
          style={{ marginTop: 24, width: "100%", padding: 14, background: !name.trim()?"#E5E7EB":"#1B4332", color: !name.trim()?"#9CA3AF":"#FFF", border: "none", borderRadius: 12, fontSize: 15, fontWeight: 600, cursor: !name.trim()?"not-allowed":"pointer", fontFamily: "DM Sans, sans-serif" }}>
          Next
        </button>
      </div>
    );
  }

  // Program picker modal
  if (pickingProgram === "choose") {
    const already = new Set(selections.map(s => s.programId));
    return (
      <div style={{ padding: "24px 20px" }}>
        <h3 style={{ margin: "0 0 20px", fontSize: 17, fontWeight: 700, color: "#111827" }}>Select program</h3>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {PROGRAM_CATALOG.filter(p => !already.has(p.id)).map(p => (
            <div key={p.id} onClick={() => setPickingProgram(p.id)}
              style={{ padding: "14px 16px", border: "1.5px solid #E5E7EB", borderRadius: 10, cursor: "pointer", display: "flex", alignItems: "center", gap: 12 }}>
              <div style={{ width: 10, height: 10, borderRadius: "50%", background: p.color, flexShrink: 0 }}/>
              <span style={{ fontSize: 14, fontWeight: 500, color: "#111827" }}>{p.label}</span>
            </div>
          ))}
        </div>
        <button onClick={() => setPickingProgram(null)} style={{ marginTop: 20, width: "100%", padding: 12, background: "none", border: "1.5px solid #E5E7EB", borderRadius: 10, fontSize: 14, cursor: "pointer", fontFamily: "DM Sans, sans-serif", color: "#6B7280" }}>Cancel</button>
      </div>
    );
  }

  if (pickingProgram && pickingProgram !== "choose") {
    const prog = PROGRAM_CATALOG.find(p => p.id === pickingProgram);
    return (
      <div style={{ padding: "24px 20px" }}>
        <h3 style={{ margin: "0 0 6px", fontSize: 17, fontWeight: 700, color: "#111827" }}>{prog.label}</h3>
        <p style={{ margin: "0 0 20px", fontSize: 13, color: "#9CA3AF" }}>Select version and revision</p>
        {prog.versions.map(v => (
          <div key={v.version} style={{ marginBottom: 16 }}>
            <p style={{ margin: "0 0 8px", fontSize: 13, fontWeight: 600, color: "#374151" }}>Version {v.version}</p>
            {v.revisions.map(rev => (
              <div key={rev} onClick={() => confirmVersionRevision(prog.id, v.version, rev)}
                style={{ padding: "12px 16px", border: "1.5px solid #E5E7EB", borderRadius: 10, cursor: "pointer", marginBottom: 6, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 14, color: "#111827" }}>{rev}</span>
                <span style={{ fontSize: 12, color: prog.color, fontWeight: 600 }}>Select →</span>
              </div>
            ))}
          </div>
        ))}
        <button onClick={() => setPickingProgram(null)} style={{ marginTop: 8, width: "100%", padding: 12, background: "none", border: "1.5px solid #E5E7EB", borderRadius: 10, fontSize: 14, cursor: "pointer", fontFamily: "DM Sans, sans-serif", color: "#6B7280" }}>Back</button>
      </div>
    );
  }

  // Programs step
  return (
    <div style={{ padding: "24px 20px" }}>
      <h2 style={{ margin: "0 0 4px", fontSize: 20, fontWeight: 700, color: "#111827" }}>{name}</h2>
      <p style={{ margin: "0 0 20px", fontSize: 13, color: "#9CA3AF" }}>Add the programs being pursued</p>

      {selections.length > 0 && (
        <div style={{ marginBottom: 16 }}>
          {selections.map((sel, i) => {
            const p = PROGRAM_CATALOG.find(x => x.id === sel.programId);
            return (
              <div key={i} style={{ padding: "10px 14px", background: p.color+"12", border: `1.5px solid ${p.color}`, borderRadius: 10, marginBottom: 8, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <div>
                  <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: p.color }}>{p.label}</p>
                  <p style={{ margin: "2px 0 0", fontSize: 11, color: p.color+"BB" }}>{sel.version} · {sel.revision}</p>
                </div>
                <button onClick={() => removeSelection(i)}
                  style={{ background: "none", border: "none", color: p.color, fontSize: 18, cursor: "pointer", padding: "0 4px" }}>×</button>
              </div>
            );
          })}
        </div>
      )}

      <button onClick={startAddProgram}
        style={{ width: "100%", padding: "12px", border: "2px dashed #D1D5DB", borderRadius: 10, background: "#F9FAFB", color: "#6B7280", fontSize: 14, cursor: "pointer", fontFamily: "DM Sans, sans-serif" }}>
        + Add program
      </button>

      <button onClick={() => selections.length && onSave({ id: Date.now().toString(), name: name.trim(), advisor: advisor.trim(), programs: selections, createdAt: new Date().toISOString() })}
        disabled={!selections.length}
        style={{ marginTop: 20, width: "100%", padding: 14, background: !selections.length?"#E5E7EB":"#1B4332", color: !selections.length?"#9CA3AF":"#FFF", border: "none", borderRadius: 12, fontSize: 15, fontWeight: 600, cursor: !selections.length?"not-allowed":"pointer", fontFamily: "DM Sans, sans-serif" }}>
        Create project
      </button>
    </div>
  );
}


// ─── SHARED: ITEM ROW ────────────────────────────────────────────────────────
function ItemRow({ project, item, records, onSelectItem, showCategory }) {
  const itemCat = item._cat;
  const recKey = `${project.id}__${itemCat}__${item.id}`;
  const rec = records[recKey]||{};
  const itemPrograms = (project.programs||[]).filter(s => {
    const k = `${s.programId}||${s.version}||${s.revision}`;
    return (CHECKLIST_REGISTRY[k]||[]).some(i => i.id === item.id);
  }).map(s => PROGRAM_CATALOG.find(x => x.id === s.programId)).filter(Boolean);
  const tierBadge = null;
  return (
    <div style={{ padding: "14px 20px", borderBottom: "1px solid #F9FAFB", background: rec.status==="fail"?"#FFF5F5":"#FFF" }}>
      {showCategory && <p style={{ margin: "0 0 4px", fontSize: 10, fontWeight: 700, color: "#9CA3AF", textTransform: "uppercase", letterSpacing: "0.05em" }}>{itemCat}</p>}
      <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
        <div style={{ flex: 1 }}>
          <div style={{ display: "flex", alignItems: "baseline", gap: 7, marginBottom: 3 }}>
            {item.pointNumber && (
              <span style={{ fontSize: 11, fontWeight: 700, color: "#1565C0", background: "#EFF6FF", padding: "1px 6px", borderRadius: 5, letterSpacing: "0.02em", flexShrink: 0 }}>
                {item.pointNumber}
              </span>
            )}
            <p style={{ margin: 0, fontSize: 13, color: "#111827", lineHeight: 1.55 }}>{item.text}</p>
          </div>
          <div style={{ marginTop: 4, display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
            {tierBadge}
            {itemPrograms.map(prog => {
              const isEC = prog.id === "earthcraft_certified" || prog.id === "earthcraft_gold" || prog.id === "earthcraft_sf2024_certified" || prog.id === "earthcraft_sf2024_gold";
              const isGoldItem = isEC && item.tier === "GOLD";
              const label = isEC ? (isGoldItem ? "EarthCraft Gold" : "EarthCraft Certified") : prog.label;
              const bg = isGoldItem ? "#FEF9C3" : prog.color+"18";
              const color = isGoldItem ? "#A16207" : prog.color;
              return <span key={prog.id} style={{ fontSize: 10, padding: "1px 7px", borderRadius: 20, background: bg, color, fontWeight: 600 }}>{label}</span>;
            })}
            {rec.photos?.length>0 && <span style={{ fontSize: 11, color: "#10B981", fontWeight: 600 }}>📷</span>}
            {!rec.photos?.length && item._cat === "Minimum Rated Features" && rec.status && rec.status !== "na" && (
              <span style={{ fontSize: 10, fontWeight: 600, color: "#EF4444", background: "#FEF2F2", padding: "1px 6px", borderRadius: 20 }}>📷 missing</span>
            )}
            {rec.note && <span style={{ fontSize: 11, color: "#6B7280" }}>📝</span>}
          </div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6, flexShrink: 0 }}>
          {rec.status && <StatusBadge status={rec.status}/>}
          <button onClick={() => onSelectItem(item)}
            style={{ fontSize: 11, padding: "5px 11px", border: "1.5px solid #E5E7EB", borderRadius: 8, background: "#FFF", color: "#374151", cursor: "pointer", fontFamily: "DM Sans, sans-serif", fontWeight: 500 }}>
            {rec.status ? "Update" : "Document"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── SHARED: SEARCH BAR ───────────────────────────────────────────────────────
function SearchBar({ query, onChange, placeholder }) {
  return (
    <div style={{ position: "relative", display: "flex", alignItems: "center" }}>
      <svg style={{ position: "absolute", left: 10, zIndex: 1, pointerEvents: "none" }} width="14" height="14" viewBox="0 0 20 20" fill="none">
        <circle cx="8.5" cy="8.5" r="5.5" stroke="#9CA3AF" strokeWidth="2"/>
        <path d="M13 13L17 17" stroke="#9CA3AF" strokeWidth="2" strokeLinecap="round"/>
      </svg>
      <input
        value={query}
        onChange={e => onChange(e.target.value)}
        placeholder={placeholder}
        style={{ width: "100%", padding: "9px 32px 9px 32px", fontSize: 13, border: "1.5px solid #E5E7EB", borderRadius: 8, outline: "none", background: "#FFF", fontFamily: "DM Sans, sans-serif", color: "#111827", boxSizing: "border-box" }}
      />
      {query && (
        <button onClick={() => onChange("")}
          style={{ position: "absolute", right: 8, background: "none", border: "none", color: "#9CA3AF", fontSize: 18, cursor: "pointer", padding: 0, lineHeight: 1 }}>×</button>
      )}
    </div>
  );
}

// ─── SCREEN: PROJECT DASHBOARD ────────────────────────────────────────────────
function ProjectDashboard({ project, records, onSelectCategory, onSelectItem, auth, setAuth, updateRecord }) {
  const pg = calcProjectProgress(project, records);
  const isMRF = (cat) => cat.id === "Minimum Rated Features";
  const [query, setQuery] = useState("");
  const q = query.trim().toLowerCase();
  const [syncState, setSyncState] = useState({ running: false, done: 0, total: 0, errors: [] });

  // All items across every category, tagged with their source category
  const allProjectItems = CATEGORIES.flatMap(cat =>
    getItemsForSelection(project.programs||[], cat.id).map(i => ({ ...i, _cat: cat.id }))
  );

  // Every photo, across every item in the project, that has never been uploaded to SharePoint
  const pendingPhotoJobs = () => {
    const jobs = [];
    for (const item of allProjectItems) {
      const key = `${project.id}__${item._cat}__${item.id}`;
      const rec = records[key];
      const photosMeta = (rec?.photos || []).map(p => typeof p === "string" ? { id: p, syncedAt: null, spFileName: null } : p);
      for (const p of photosMeta) {
        if (!p.syncedAt) jobs.push({ item, key, rec, photoId: p.id });
      }
    }
    return jobs;
  };
  const pendingCount = pendingPhotoJobs().length;

  const handleUploadToSharePoint = async () => {
    if (!auth) { startLogin(); return; }
    const token = await getValidToken(auth, setAuth);
    if (!token) { setSyncState({ running: false, done: 0, total: 0, errors: ["Could not connect to SharePoint — please reconnect and try again."] }); return; }
    const jobs = pendingPhotoJobs();
    if (!jobs.length) return;
    setSyncState({ running: true, done: 0, total: jobs.length, errors: [] });
    const workingRecords = {}; // key -> latest record as we mutate it within this batch
    let done = 0; const errors = [];
    for (const job of jobs) {
      const base = workingRecords[job.key] || job.rec;
      try {
        const dataUrl = await idbGetPhoto(`${job.key}__${job.photoId}`);
        if (!dataUrl) throw new Error("Photo not found locally");
        const nextNum = base.nextPhotoNum || 1;
        const label = sanitizeSpName(job.item.pointNumber || job.item.text || job.item.id);
        const fileName = `${label} - ${nextNum}.${extFromDataUrl(dataUrl)}`;
        await uploadPhotoToSharePoint(token, fileName, dataUrl);
        const updatedPhotos = (base.photos||[]).map(p => {
          const pid = typeof p === "string" ? p : p.id;
          if (pid !== job.photoId) return typeof p === "string" ? { id: p, syncedAt: null, spFileName: null } : p;
          return { id: pid, syncedAt: new Date().toISOString(), spFileName: fileName };
        });
        const updatedRec = { ...base, photos: updatedPhotos, nextPhotoNum: nextNum + 1 };
        workingRecords[job.key] = updatedRec;
        updateRecord(project.id, job.item._cat, job.item.id, updatedRec);
      } catch (e) {
        errors.push(`${job.item.pointNumber || job.item.id}: ${e.message}`);
      }
      done++;
      setSyncState({ running: true, done, total: jobs.length, errors });
    }
    setSyncState({ running: false, done, total: jobs.length, errors });
  };

  const searchResults = q
    ? allProjectItems.filter(i =>
        i.text.toLowerCase().includes(q) ||
        (i.pointNumber||"").toLowerCase().includes(q)
      )
    : null;

  const CatRow = ({ cat }) => {
    const items = getItemsForSelection(project.programs||[], cat.id);
    const mrf = cat.id === "Minimum Rated Features";
    const p = calcCatProgress(items, records, project.id, cat.id);
    if (!items.length && !mrf) return null;
    const accentColor = mrf ? "#059669" : (p.fail>0?"#EF4444":p.pct===100?"#10B981":"#3B82F6");
    return (
      <div onClick={() => onSelectCategory(cat)}
        style={{ padding: "14px 20px", borderBottom: "1px solid #F9FAFB", cursor: "pointer" }}
        onTouchStart={e => e.currentTarget.style.background="#F9FAFB"}
        onTouchEnd={e => e.currentTarget.style.background="#FFF"}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flex: 1, minWidth: 0 }}>
            <span style={{ fontSize: 10, fontWeight: 700, color: mrf?"#059669":"#6B7280", background: mrf?"#D1FAE5":"#F3F4F6", padding: "2px 6px", borderRadius: 4, flexShrink: 0 }}>{cat.code}</span>
            <span style={{ fontSize: 14, fontWeight: 600, color: mrf?"#059669":"#111827", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{cat.id}</span>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
            {p.fail > 0 && <span style={{ fontSize: 11, fontWeight: 700, color: "#EF4444", background: "#FEE2E2", padding: "2px 7px", borderRadius: 20 }}>{p.fail} fail</span>}
            {!mrf && <span style={{ fontSize: 13, fontWeight: 700, color: accentColor }}>{p.pct}%</span>}
            {mrf && items.length === 0 && <span style={{ fontSize: 11, color: "#9CA3AF", fontStyle: "italic" }}>Coming soon</span>}
            <span style={{ color: "#D1D5DB" }}>›</span>
          </div>
        </div>
        {!mrf && <ProgressBar pct={p.pct} fail={p.fail}/>}
        <p style={{ margin: "4px 0 0", fontSize: 11, color: mrf?"#059669":"#9CA3AF" }}>
          {mrf ? "Energy modeling documentation" : `${p.pass+p.na}/${p.total} verified${p.fail>0?` · ${p.fail} failing`:""}`}
        </p>
      </div>
    );
  };

  return (
    <div style={{ paddingBottom: 40 }}>
      {/* Hero */}
      <div style={{ background: "linear-gradient(135deg,#1B4332,#2D6A4F)", padding: "24px 20px 20px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
          <div style={{ position: "relative" }}>
            <ProgressRing pct={pg.pct} size={72} stroke={6} fail={pg.fail}/>
            <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: 700, color: pg.fail>0?"#EF4444":pg.pct===100?"#10B981":"#60A5FA" }}>{pg.pct}%</div>
          </div>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: "#FFF" }}>{project.name}</h2>
            <p style={{ margin: "3px 0 0", fontSize: 12, color: "#A7F3D0" }}>{pg.verified}/{pg.total} items verified</p>
            {project.advisor && <p style={{ margin: "2px 0 0", fontSize: 12, color: "#6EE7B7" }}>TA: {project.advisor}</p>}
            {pg.fail>0 && <p style={{ margin: "2px 0 0", fontSize: 12, color: "#FCA5A5", fontWeight: 600 }}>⚠ {pg.fail} item{pg.fail>1?"s":""} failing</p>}
          </div>
        </div>
        <div style={{ marginTop: 14, display: "flex", gap: 6, flexWrap: "wrap" }}>
          {(project.programs||[]).map((sel, i) => {
            const p = PROGRAM_CATALOG.find(x => x.id === sel.programId);
            return (
              <span key={i} style={{ fontSize: 10, padding: "3px 9px", borderRadius: 20, background: "rgba(255,255,255,0.15)", color: "#FFF", fontWeight: 500 }}>
                {p?.label} {sel.version} {sel.revision}
              </span>
            );
          })}
        </div>
      </div>

      {/* SharePoint photo sync */}
      <div style={{ padding: "12px 20px", background: "#F9FAFB", borderBottom: "1px solid #E5E7EB", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
        <div style={{ minWidth: 0 }}>
          <p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: "#374151" }}>☁ SharePoint photo sync</p>
          <p style={{ margin: "2px 0 0", fontSize: 11, color: "#9CA3AF", wordBreak: "break-word" }}>
            📁 {SP_FOLDER}
          </p>
          <p style={{ margin: "2px 0 0", fontSize: 11, color: "#9CA3AF" }}>
            {syncState.running ? `Uploading ${syncState.done}/${syncState.total}…`
              : !auth ? "Not connected"
              : pendingCount === 0 ? "All photos synced"
              : `${pendingCount} photo${pendingCount>1?"s":""} pending`}
          </p>
          {!syncState.running && syncState.errors.length>0 && (
            <p style={{ margin: "2px 0 0", fontSize: 11, color: "#EF4444" }}>{syncState.errors.length} failed — tap to retry</p>
          )}
        </div>
        <button onClick={handleUploadToSharePoint} disabled={syncState.running || (auth && pendingCount===0)}
          style={{ fontSize: 12, fontWeight: 600, color: "#FFF", background: !auth ? "#0078D4" : (pendingCount===0 ? "#D1D5DB" : "#059669"), border: "none", borderRadius: 6, padding: "6px 14px", cursor: syncState.running ? "wait" : "pointer", flexShrink: 0, fontFamily: "DM Sans, sans-serif" }}>
          {!auth ? "Connect" : syncState.running ? "Uploading…" : pendingCount===0 ? "Synced" : "Upload to SharePoint"}
        </button>
      </div>

      {/* Global search bar */}
      <div style={{ padding: "12px 20px", borderBottom: "1px solid #F3F4F6", background: "#F9FAFB" }}>
        <SearchBar query={query} onChange={setQuery} placeholder="Search all items across every category…"/>
        {q && <p style={{ margin: "6px 0 0", fontSize: 11, color: "#9CA3AF" }}>{searchResults.length} result{searchResults.length !== 1 ? "s" : ""} across all categories</p>}
      </div>

      {/* Search results OR category list */}
      {searchResults ? (
        <>
          {searchResults.length === 0 && (
            <div style={{ padding: "32px 20px", textAlign: "center", color: "#9CA3AF" }}>
              <p style={{ margin: 0, fontSize: 14 }}>No items match "{query}"</p>
            </div>
          )}
          {searchResults.map(item => (
            <ItemRow key={item.id + item._cat} project={project} item={item} records={records} onSelectItem={onSelectItem} showCategory={true}/>
          ))}
        </>
      ) : (
        <>
          {CATEGORIES.map(cat => <CatRow key={cat.id} cat={cat}/>)}
        </>
      )}
    </div>
  );
}

// ─── SCREEN: CHECKLIST ────────────────────────────────────────────────────────
// Search is scoped to this category only.
function ChecklistView({ project, category, records, onSelectItem }) {
  const allItems = getItemsForSelection(project.programs||[], category.id).map(i => ({ ...i, _cat: category.id }));
  const [query, setQuery] = useState("");
  const p = calcCatProgress(allItems, records, project.id, category.id);
  const q = query.trim().toLowerCase();

  const displayItems = q
    ? allItems.filter(i =>
        i.text.toLowerCase().includes(q) ||
        (i.pointNumber||"").toLowerCase().includes(q)
      )
    : allItems;

  return (
    <div style={{ paddingBottom: 40 }}>
      <div style={{ padding: "14px 20px 12px", background: "#F9FAFB", borderBottom: "1px solid #F3F4F6" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: "#111827" }}>{category.id}</h3>
            <p style={{ margin: "3px 0 0", fontSize: 12, color: "#9CA3AF" }}>
              {p.pass+p.na}/{p.total} verified
              {p.fail > 0 && <span style={{ color: "#EF4444" }}> · {p.fail} failing</span>}
            </p>
          </div>
          <span style={{ fontSize: 22, fontWeight: 700, color: p.fail>0?"#EF4444":p.pct===100?"#10B981":"#3B82F6" }}>{p.pct}%</span>
        </div>
        <div style={{ marginTop: 10 }}><ProgressBar pct={p.pct} fail={p.fail}/></div>
        <div style={{ marginTop: 12 }}>
          <SearchBar query={query} onChange={setQuery} placeholder={`Search in ${category.id}…`}/>
        </div>
        {q && <p style={{ margin: "6px 0 0", fontSize: 11, color: "#9CA3AF" }}>{displayItems.length} of {allItems.length} items</p>}
      </div>
      {displayItems.length === 0 && q && (
        <div style={{ padding: "32px 20px", textAlign: "center", color: "#9CA3AF" }}>
          <p style={{ margin: 0, fontSize: 14 }}>No items match "{query}"</p>
        </div>
      )}
      {displayItems.map(item => (
        <ItemRow key={item.id} project={project} item={item} records={records} onSelectItem={onSelectItem} showCategory={false}/>
      ))}
    </div>
  );
}


// Strips a decimal text input down to at most one "." and 2 digits after it (e.g. U-Value, SHGC)
function sanitizeDecimal2(raw) {
  let v = raw.replace(/[^0-9.]/g, "");
  const firstDot = v.indexOf(".");
  if (firstDot !== -1) v = v.slice(0, firstDot + 1) + v.slice(firstDot + 1).replace(/\./g, "");
  const [intPart, decPart] = v.split(".");
  return decPart !== undefined ? `${intPart}.${decPart.slice(0, 2)}` : v;
}

// Renders just the field inputs for one entry — shared by the repeatable list and single-entry views
function EntryFieldInputs({ fields, entry, onFieldChange }) {
  const inputStyle = { width: "100%", padding: "10px 12px", border: "1.5px solid #E5E7EB", borderRadius: 8, fontSize: 14, fontFamily: "DM Sans, sans-serif", color: "#111827", boxSizing: "border-box" };
  return fields.map(f => (
    <div key={f.key}>
      <label style={{ display: "block", marginBottom: 4, fontSize: 11, color: "#9CA3AF" }}>{f.label}</label>
      {f.type === "select" ? (
        <select value={entry[f.key]||""} onChange={e => onFieldChange(f.key, e.target.value)}
          style={{ ...inputStyle, background: "#FFF" }}>
          <option value="">Select...</option>
          {f.options.map(([val, label]) => <option key={val} value={val}>{label}</option>)}
        </select>
      ) : f.type === "decimal" ? (
        <input type="text" inputMode="decimal" value={entry[f.key]||""} onChange={e => onFieldChange(f.key, sanitizeDecimal2(e.target.value))}
          placeholder="0.00" style={inputStyle}/>
      ) : (
        <input type="text" value={entry[f.key]||""} onChange={e => onFieldChange(f.key, e.target.value)}
          placeholder={f.label} style={inputStyle}/>
      )}
    </div>
  ));
}

// Repeatable structured entries (currently: wall assemblies only) — see MULTI_ENTRY_CONFIG
function MultiEntryList({ config, entries, onAdd, onRemove, onFieldChange }) {
  return (
    <div style={{ marginBottom: 24 }}>
      <p style={{ margin: "0 0 10px", fontSize: 11, fontWeight: 700, color: "#6B7280", textTransform: "uppercase", letterSpacing: "0.06em" }}>
        {config.label}
      </p>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {entries.map((entry, idx) => (
          <div key={idx} style={{ position: "relative", padding: 14, border: "1.5px solid #E5E7EB", borderRadius: 10, display: "flex", flexDirection: "column", gap: 10 }}>
            <button onClick={() => onRemove(idx)} title={`Remove this ${config.entryLabel}`}
              style={{ position: "absolute", top: 8, right: 8, width: 22, height: 22, borderRadius: "50%", background: "#F3F4F6", border: "none", color: "#6B7280", fontSize: 13, cursor: "pointer", lineHeight: "22px" }}>×</button>
            <EntryFieldInputs fields={config.fields} entry={entry} onFieldChange={(key, val) => onFieldChange(idx, key, val)}/>
          </div>
        ))}
      </div>
      <button onClick={onAdd}
        style={{ marginTop: 10, width: "100%", padding: "10px", border: "1.5px dashed #D1D5DB", borderRadius: 10, background: "#F9FAFB", color: "#6B7280", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "DM Sans, sans-serif" }}>
        + Add {config.entryLabel}
      </button>
    </div>
  );
}

// Single fixed set of structured fields — no add/remove (ceiling, foundation, rim & band, duct insulation)
function SingleEntryFields({ config, entry, onFieldChange }) {
  return (
    <div style={{ marginBottom: 24 }}>
      <p style={{ margin: "0 0 10px", fontSize: 11, fontWeight: 700, color: "#6B7280", textTransform: "uppercase", letterSpacing: "0.06em" }}>
        {config.label}
      </p>
      <div style={{ padding: 14, border: "1.5px solid #E5E7EB", borderRadius: 10, display: "flex", flexDirection: "column", gap: 10 }}>
        <EntryFieldInputs fields={config.fields} entry={entry} onFieldChange={onFieldChange}/>
      </div>
    </div>
  );
}

// ─── SCREEN: ITEM DETAIL ──────────────────────────────────────────────────────
// Autosaves on status tap and on photo add/remove. Note saves on blur.
function ItemDetail({ project, category, item, record, onSave }) {
  const [status, setStatus] = useState(record?.status||"");
  const [note, setNote] = useState(record?.note||"");
  const [photos, setPhotos] = useState([]);   // [{id, dataUrl}] — dataUrls live in IndexedDB
  const [photosLoading, setPhotosLoading] = useState(!!record?.photos?.length);
  const [saved, setSaved] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const entryConfig = MULTI_ENTRY_CONFIG[item.id];
  const [entries, setEntries] = useState(() => {
    if (record?.entries?.length) return record.entries;
    if (entryConfig && !entryConfig.repeatable) return [Object.fromEntries(entryConfig.fields.map(f => [f.key, ""]))];
    return [];
  });
  const fileRef = useRef();
  const noteTimer = useRef();
  // Snapshot of note+timestamp as they stood when the note field was last focused —
  // used to log ONE history entry per edit session instead of one per autosave.
  const noteSnapshot = useRef({ note: record?.note||"", updatedAt: record?.updatedAt||null });

  // Mirrors of the state above, always kept current. The debounced note autosave (below) can
  // fire up to 800ms after it's scheduled — if a photo/entry/status changed in that window, a
  // save() call still holding the OLD state closure would silently overwrite the newer data.
  // Refs sidestep that: whichever save() closure runs, it reads the freshest values here.
  const statusRef = useRef(status);
  const noteRef = useRef(note);
  const photosRef = useRef(photos);
  const entriesRef = useRef(entries);

  // Derive stable key for IndexedDB lookup — each photo gets its own suffixed slot
  const photoKey = `${project.id}__${category.id}__${item.id}`;
  const MAX_PHOTOS = 5;

  const isMRF = category.id === "Minimum Rated Features";
  const photoRequired = (val) => isMRF && val !== "na" && photos.length === 0;

  // Load photos from IndexedDB on mount. Entries may be plain id strings (pre-sync-tracking
  // records) or {id, syncedAt, spFileName} objects — normalize either way.
  useEffect(() => {
    const meta = (record?.photos || []).map(p => typeof p === "string" ? { id: p, syncedAt: null, spFileName: null } : p);
    if (!meta.length) { setPhotosLoading(false); return; }
    Promise.all(meta.map(m => idbGetPhoto(`${photoKey}__${m.id}`).then(dataUrl => ({ ...m, dataUrl }))))
      .then(results => { setPhotos(results.filter(r => r.dataUrl)); setPhotosLoading(false); })
      .catch(() => setPhotosLoading(false));
  }, [photoKey]);

  const save = (overrides = {}) => {
    // photos field in record holds sync metadata only — the image data lives in IndexedDB.
    // Base values come from the refs (always current), not the state closures (can be stale).
    const { archive, ...visibleOverrides } = overrides;
    const rec = {
      status: statusRef.current,
      note: noteRef.current,
      photos: photosRef.current.map(({ id, syncedAt, spFileName }) => ({ id, syncedAt: syncedAt||null, spFileName: spFileName||null })),
      entries: entriesRef.current,
      updatedAt: new Date().toISOString(),
      ...visibleOverrides,
    };
    // Notes/entries/photos may be documented before a status is picked (e.g. before a photo is
    // uploaded) — only skip saving if there's truly nothing to save yet.
    const hasEntryContent = rec.entries?.some(e => Object.values(e).some(v => v));
    if (!rec.status && !rec.note && !rec.photos?.length && !hasEntryContent) return;
    if (archive) {
      rec.history = [...(record?.history||[]), archive];
    } else if (record?.history) {
      rec.history = record.history;
    }
    onSave(rec);
    setSaved(true);
    setTimeout(() => setSaved(false), 1200);
  };

  const handleStatus = (val) => {
    if (photoRequired(val)) return;
    setStatus(val);
    statusRef.current = val;
    // A status change is a discrete, deliberate action — archive it every time, unlike note autosaves
    const archive = (record?.status && val !== record.status)
      ? { status: record.status, note: record.note||"", updatedAt: record.updatedAt }
      : undefined;
    save({ status: val, archive });
  };

  const addEntry = () => {
    const blank = Object.fromEntries(entryConfig.fields.map(f => [f.key, ""]));
    const next = [...entriesRef.current, blank];
    setEntries(next);
    entriesRef.current = next;
    save({ entries: next });
  };

  const removeEntry = (idx) => {
    const next = entriesRef.current.filter((_, i) => i !== idx);
    setEntries(next);
    entriesRef.current = next;
    save({ entries: next });
  };

  const updateEntry = (idx, key, val) => {
    const next = entriesRef.current.map((e, i) => i === idx ? { ...e, [key]: val } : e);
    setEntries(next);
    entriesRef.current = next;
    save({ entries: next });
  };

  const handleAddPhoto = e => {
    const file = e.target.files[0]; e.target.value = ""; if (!file || photosRef.current.length >= MAX_PHOTOS) return;
    const reader = new FileReader();
    reader.onload = async ev => {
      const dataUrl = ev.target.result;
      const id = `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
      await idbSavePhoto(`${photoKey}__${id}`, dataUrl);
      const next = [...photosRef.current, { id, dataUrl, syncedAt: null, spFileName: null }];
      setPhotos(next);
      photosRef.current = next;
      save({ photos: next.map(({ id, syncedAt, spFileName }) => ({ id, syncedAt, spFileName })) });
    };
    reader.readAsDataURL(file);
  };

  const handleRemovePhoto = async (id) => {
    await idbDeletePhoto(`${photoKey}__${id}`);
    const next = photosRef.current.filter(p => p.id !== id);
    setPhotos(next);
    photosRef.current = next;
    save({ photos: next.map(({ id, syncedAt, spFileName }) => ({ id, syncedAt, spFileName })) });
  };

  const handleNoteFocus = () => {
    // Baseline for this edit session — used on blur to decide whether to log one history entry
    noteSnapshot.current = { note: record?.note||"", updatedAt: record?.updatedAt||null };
  };

  const handleNoteChange = (val) => {
    setNote(val);
    noteRef.current = val;
    // Debounce note saves — only write after 800ms of no typing. Never archives history itself,
    // so pausing mid-sentence doesn't spam the log; only the final blur below does that.
    // Not gated on status — notes can be documented before a status/photo exists.
    clearTimeout(noteTimer.current);
    noteTimer.current = setTimeout(() => save({ note: val }), 800);
  };

  const handleNoteBlur = () => {
    clearTimeout(noteTimer.current);
    const changed = record?.status && note !== noteSnapshot.current.note;
    const archive = changed ? { status: statusRef.current, note: noteSnapshot.current.note, updatedAt: noteSnapshot.current.updatedAt || record?.updatedAt } : undefined;
    save({ note, archive });
  };

  const itemPrograms = (project.programs||[]).filter(s => {
    const k = `${s.programId}||${s.version}||${s.revision}`;
    return (CHECKLIST_REGISTRY[k]||[]).some(i => i.id === item.id);
  }).map(s => PROGRAM_CATALOG.find(x => x.id === s.programId)).filter(Boolean);

  return (
    <div style={{ padding: "20px 20px 40px" }}>
      {/* Item card */}
      <div style={{ background: "#F9FAFB", border: "1px solid #F3F4F6", borderRadius: 12, padding: "14px 16px", marginBottom: 24 }}>
        {item.pointNumber && (
          <span style={{ display: "inline-block", marginBottom: 8, fontSize: 12, fontWeight: 700, color: "#1565C0", background: "#EFF6FF", padding: "2px 8px", borderRadius: 6, letterSpacing: "0.02em" }}>
            {item.pointNumber}
          </span>
        )}
        <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: "#111827", lineHeight: 1.55 }}>{item.text}</p>
        <p style={{ margin: "6px 0 0", fontSize: 11, color: "#9CA3AF" }}>{category.id} · {project.name}</p>
        <div style={{ marginTop: 6, display: "flex", gap: 6, flexWrap: "wrap" }}>
          {item.mergedWith && <span style={{ fontSize: 10, padding: "1px 7px", borderRadius: 20, background: "#F0FDF4", color: "#166534", fontWeight: 600 }}>Multi-program</span>}
          {itemPrograms.map(prog => {
              const isEC = prog.id === "earthcraft_certified" || prog.id === "earthcraft_gold" || prog.id === "earthcraft_sf2024_certified" || prog.id === "earthcraft_sf2024_gold";
              const isGoldItem = isEC && item.tier === "GOLD";
              const label = isEC ? (isGoldItem ? "EarthCraft Gold" : "EarthCraft Certified") : prog.label;
              const bg = isGoldItem ? "#FEF9C3" : prog.color+"18";
              const color = isGoldItem ? "#A16207" : prog.color;
              return <span key={prog.id} style={{ fontSize: 10, padding: "1px 7px", borderRadius: 20, background: bg, color, fontWeight: 600 }}>{label}</span>;
            })}
        </div>
      </div>

      {/* Structured entries (wall/ceiling/foundation assemblies, etc.) */}
      {entryConfig && (entryConfig.repeatable ? (
        <MultiEntryList config={entryConfig} entries={entries} onAdd={addEntry} onRemove={removeEntry} onFieldChange={updateEntry}/>
      ) : (
        <SingleEntryFields config={entryConfig} entry={entries[0]||{}} onFieldChange={(key, val) => updateEntry(0, key, val)}/>
      ))}

      {/* Autosave indicator */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <p style={{ margin: 0, fontSize: 11, fontWeight: 700, color: "#6B7280", textTransform: "uppercase", letterSpacing: "0.06em" }}>Status</p>
        {saved && <span style={{ fontSize: 11, color: "#10B981", fontWeight: 600 }}>✓ Saved</span>}
      </div>

      {/* Photos — up to 5 per item, shown FIRST for MRF items so the requirement is front and center */}
      <div style={{ marginBottom: 24 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <p style={{ margin: 0, fontSize: 11, fontWeight: 700, color: "#6B7280", textTransform: "uppercase", letterSpacing: "0.06em" }}>
            Photos{isMRF && <span style={{ color: "#EF4444" }}> *</span>}
          </p>
          {isMRF && photos.length===0 && (
            <span style={{ fontSize: 11, fontWeight: 600, color: "#EF4444", background: "#FEF2F2", padding: "2px 8px", borderRadius: 20 }}>Required to confirm</span>
          )}
          {isMRF && photos.length>0 && (
            <span style={{ fontSize: 11, fontWeight: 600, color: "#10B981", background: "#F0FDF4", padding: "2px 8px", borderRadius: 20 }}>✓ {photos.length} photo{photos.length>1?"s":""} uploaded</span>
          )}
        </div>
        {photosLoading ? (
          <div style={{ width: "100%", height: 80, borderRadius: 12, background: "#F3F4F6", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <span style={{ fontSize: 12, color: "#9CA3AF" }}>Loading photos…</span>
          </div>
        ) : (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {photos.map(p => (
              <div key={p.id} style={{ position: "relative", width: 84, height: 84 }}>
                <img src={p.dataUrl} alt="" style={{ width: 84, height: 84, borderRadius: 10, display: "block", objectFit: "cover" }}/>
                <button onClick={() => handleRemovePhoto(p.id)}
                  style={{ position: "absolute", top: 4, right: 4, width: 22, height: 22, borderRadius: "50%", background: "rgba(0,0,0,.6)", border: "none", color: "#FFF", fontSize: 13, cursor: "pointer" }}>×</button>
                {p.syncedAt && (
                  <span title={`Uploaded to SharePoint as ${p.spFileName}`}
                    style={{ position: "absolute", bottom: 4, left: 4, fontSize: 11, background: "rgba(16,185,129,.9)", color: "#FFF", borderRadius: "50%", width: 18, height: 18, display: "flex", alignItems: "center", justifyContent: "center" }}>☁</span>
                )}
              </div>
            ))}
            {photos.length < MAX_PHOTOS && (
              <button onClick={() => fileRef.current.click()} title={isMRF && photos.length===0 ? "Upload a photo to enable confirmation" : "Add a photo"}
                style={{ width: 84, height: 84, border: `2px dashed ${isMRF && photos.length===0 ? "#FCA5A5" : "#D1D5DB"}`, borderRadius: 10, background: isMRF && photos.length===0 ? "#FFF5F5" : "#F9FAFB", color: isMRF && photos.length===0 ? "#EF4444" : "#6B7280", fontSize: 24, cursor: "pointer", fontFamily: "DM Sans, sans-serif" }}>
                +
              </button>
            )}
          </div>
        )}
        {photos.length===0 && (
          <p style={{ margin: "8px 0 0", fontSize: 12, color: isMRF ? "#EF4444" : "#9CA3AF" }}>
            {isMRF ? "Upload a photo to enable confirmation" : "Take or upload a photo"}
          </p>
        )}
        {photos.length>0 && <p style={{ margin: "8px 0 0", fontSize: 11, color: "#9CA3AF" }}>{photos.length}/{MAX_PHOTOS} photos</p>}
        <input ref={fileRef} type="file" accept="image/*" capture="environment" onChange={handleAddPhoto} style={{ display: "none" }}/>
      </div>

      {/* Status */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <p style={{ margin: 0, fontSize: 11, fontWeight: 700, color: "#6B7280", textTransform: "uppercase", letterSpacing: "0.06em" }}>Status</p>
        {saved && <span style={{ fontSize: 11, color: "#10B981", fontWeight: 600 }}>✓ Saved</span>}
      </div>

      {/* Status buttons — Pass and Fail blocked on MRF without photo */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 24 }}>
        {[["pass","#D1FAE5","#065F46","#10B981","Pass"],["fail","#FEE2E2","#991B1B","#EF4444","Fail"],["na","#F3F4F6","#4B5563","#9CA3AF","N/A"]].map(([id,bg,col,brd,label]) => {
          const blocked = photoRequired(id);
          return (
            <button key={id} onClick={() => handleStatus(id)} disabled={blocked}
              title={blocked ? "Upload a photo first" : ""}
              style={{ padding: "12px 8px", border: `2px solid ${status===id ? brd : blocked ? "#F3F4F6" : "#E5E7EB"}`, borderRadius: 10, background: status===id ? bg : blocked ? "#F9FAFB" : "#FFF", color: status===id ? col : blocked ? "#D1D5DB" : "#6B7280", fontSize: 14, fontWeight: 700, cursor: blocked ? "not-allowed" : "pointer", fontFamily: "DM Sans, sans-serif", position: "relative" }}>
              {label}
              {blocked && <span style={{ display: "block", fontSize: 9, fontWeight: 400, marginTop: 2, color: "#FCA5A5" }}>photo first</span>}
            </button>
          );
        })}
      </div>

      {/* Timestamp */}
      {record?.updatedAt && (
        <p style={{ margin: "-16px 0 20px", fontSize: 11, color: "#9CA3AF" }}>
          Last updated {fmtDate(record.updatedAt)}
        </p>
      )}

      {/* Note */}
      <p style={{ margin: "0 0 10px", fontSize: 11, fontWeight: 700, color: "#6B7280", textTransform: "uppercase", letterSpacing: "0.06em" }}>
        Note <span style={{ fontWeight: 400, color: "#9CA3AF", textTransform: "none" }}>(optional)</span>
      </p>
      <textarea value={note} onChange={e => handleNoteChange(e.target.value)}
        onFocus={handleNoteFocus} onBlur={handleNoteBlur}
        placeholder="Add a note..." rows={3}
        style={{ width: "100%", padding: "12px 14px", border: "1.5px solid #E5E7EB", borderRadius: 10, fontSize: 14, fontFamily: "DM Sans, sans-serif", color: "#111827", resize: "none", outline: "none", boxSizing: "border-box" }}/>

      {/* History — prior status changes, most recent first */}
      {record?.history?.length > 0 && (
        <div style={{ marginTop: 24 }}>
          <button onClick={() => setHistoryOpen(o => !o)}
            style={{ display: "flex", alignItems: "center", gap: 6, width: "100%", background: "none", border: "none", padding: 0, cursor: "pointer", fontFamily: "DM Sans, sans-serif" }}>
            <span style={{ fontSize: 11, fontWeight: 700, color: "#6B7280", textTransform: "uppercase", letterSpacing: "0.06em" }}>
              History ({record.history.length} prior {record.history.length === 1 ? "entry" : "entries"})
            </span>
            <span style={{ fontSize: 10, color: "#9CA3AF", transform: historyOpen ? "rotate(180deg)" : "none" }}>▾</span>
          </button>
          {historyOpen && (
            <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 8 }}>
              {[...record.history].reverse().map((h, i) => (
                <div key={i} style={{ padding: "10px 12px", background: "#F9FAFB", border: "1px solid #F3F4F6", borderRadius: 10 }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: h.note ? 6 : 0 }}>
                    <StatusBadge status={h.status}/>
                    <span style={{ fontSize: 11, color: "#9CA3AF" }}>{fmtDate(h.updatedAt)}</span>
                  </div>
                  {h.note && <p style={{ margin: 0, fontSize: 12, color: "#4B5563", lineHeight: 1.5 }}>{h.note}</p>}
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}


// ─── APP ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [data, setData] = useState(() => loadData());
  const [auth, setAuth] = useState(() => loadAuth());

  // Handle OAuth redirect callback — fires when Microsoft redirects back with ?code=
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const code = params.get('code');
    const error = params.get('error');
    if (!code && !error) return;
    // Clean the URL immediately so a refresh doesn't re-trigger
    window.history.replaceState({}, document.title, window.location.pathname);
    if (error) { console.error('Auth error:', error, params.get('error_description')); return; }
    exchangeCode(code).then(async result => {
      if (!result.access_token) { console.error('Token exchange failed:', result); return; }
      const user = await fetchUserInfo(result.access_token);
      const authData = {
        accessToken: result.access_token,
        refreshToken: result.refresh_token,
        expiresAt: Date.now() + result.expires_in * 1000,
        user: { name: user.displayName, email: user.userPrincipalName },
      };
      saveAuth(authData);
      setAuth(authData);
    }).catch(e => console.error('Auth callback error:', e));
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const [screen, setScreen] = useState("projects");
  const [activeProject, setActiveProject] = useState(null);
  const [activeCategory, setActiveCategory] = useState(null);
  const [activeItem, setActiveItem] = useState(null);

  useEffect(() => { saveData(data); }, [data]);

  const updateRecord = (projectId, categoryId, itemId, value) => {
    const key = `${projectId}__${categoryId}__${itemId}`;
    setData(d => ({ ...d, records: { ...d.records, [key]: value } }));
  };

  const deleteProject = (projectId) => {
    setData(d => {
      const newRecords = { ...d.records };
      Object.keys(newRecords).forEach(k => { if (k.startsWith(projectId + "__")) delete newRecords[k]; });
      return { ...d, projects: d.projects.filter(p => p.id !== projectId), records: newRecords };
    });
  };

  const navBack = () => {
    if (screen === "item") { setScreen("checklist"); setActiveItem(null); }
    else if (screen === "checklist") { setScreen("dashboard"); setActiveCategory(null); }
    else if (screen === "dashboard") { setScreen("projects"); setActiveProject(null); }
    else if (screen === "create") setScreen("projects");
  };

  const titles = { projects: "Field Documentation Tracker", create: "New project", dashboard: activeProject?.name||"", checklist: activeCategory?.id||"", item: "Document item" };

  return (
    <div style={{ maxWidth: 430, margin: "0 auto", minHeight: "100vh", background: "#FFF", fontFamily: "DM Sans, sans-serif" }}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap" rel="stylesheet"/>
      <div style={{ position: "sticky", top: 0, zIndex: 40, background: "#FFF", borderBottom: "1px solid #F3F4F6", padding: "14px 20px", display: "flex", alignItems: "center", gap: 12 }}>
        {screen !== "projects" && (
          <button onClick={navBack} style={{ width: 32, height: 32, border: "none", background: "none", cursor: "pointer", fontSize: 22, color: "#374151", padding: 0, flexShrink: 0 }}>‹</button>
        )}
        <div style={{ display: "flex", alignItems: "center", gap: 8, flex: 1, minWidth: 0 }}>
          {screen === "projects" && <span style={{ fontSize: 20 }}>☑️</span>}
          <h1 style={{ margin: 0, fontSize: screen==="projects"?20:17, fontWeight: 700, color: "#111827", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{titles[screen]}</h1>
        </div>
      </div>

      {screen === "projects" && <ProjectList projects={data.projects} records={data.records} onSelect={p=>{setActiveProject(p);setScreen("dashboard");}} onCreate={()=>setScreen("create")} onDelete={deleteProject} auth={auth} onLogout={()=>{clearAuth();setAuth(null);}}/>}
      {screen === "create" && <CreateProject onSave={proj=>{setData(d=>({...d,projects:[...d.projects,proj]}));setScreen("projects");}} onBack={navBack}/>}
      {screen === "dashboard" && activeProject && (
        <ProjectDashboard
          project={activeProject}
          records={data.records}
          onSelectCategory={cat=>{setActiveCategory(cat);setScreen("checklist");}}
          onSelectItem={item=>{setActiveItem(item);setScreen("item");}}
          auth={auth}
          setAuth={setAuth}
          updateRecord={updateRecord}
        />
      )}
      {screen === "checklist" && activeProject && activeCategory && (
        <ChecklistView
          project={activeProject}
          category={activeCategory}
          records={data.records}
          onSelectItem={item=>{setActiveItem(item);setScreen("item");}}
        />
      )}
      {screen === "item" && activeProject && activeItem && (
        <ItemDetail
          project={activeProject}
          category={{ id: activeItem._cat || activeCategory?.id }}
          item={activeItem}
          record={data.records[`${activeProject.id}__${activeItem._cat||activeCategory?.id}__${activeItem.id}`]}
          onSave={val=>{updateRecord(activeProject.id, activeItem._cat||activeCategory?.id, activeItem.id, val);}}
        />
      )}
    </div>
  );
}