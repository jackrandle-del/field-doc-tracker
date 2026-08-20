import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

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

// Resolves a folder path (relative to the site's default drive root, e.g. a project's
// "Site Visits" folder) to its Graph item id. Throws if the folder doesn't exist.
const getSharePointFolderId = async (siteId, token, path) => {
  const cleanPath = path.replace(/^\/+|\/+$/g, "");
  const encodedPath = cleanPath.split("/").map(encodeURIComponent).join("/");
  const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodedPath}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    throw new Error(res.status === 404
      ? "SharePoint folder not found — check the path and try again."
      : `Could not access SharePoint folder (${res.status})`);
  }
  return (await res.json()).id;
};

// Creates a subfolder under a known parent folder, or reuses it if it already exists
// (e.g. a second upload on the same inspection date) — never creates a duplicate.
const createOrGetSubfolder = async (siteId, token, parentId, name) => {
  const createRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${parentId}/children`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
  });
  if (createRes.ok) return (await createRes.json()).id;
  if (createRes.status === 409) {
    const getRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${parentId}:/${encodeURIComponent(name)}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (getRes.ok) return (await getRes.json()).id;
  }
  const err = await createRes.json().catch(() => ({}));
  throw new Error(err.error?.message || `Could not create date folder (${createRes.status})`);
};

// MM.DD.YYYY, matching the existing inspection-date subfolder naming already used in SharePoint
const formatDateFolderName = (d) =>
  `${String(d.getMonth()+1).padStart(2,"0")}.${String(d.getDate()).padStart(2,"0")}.${d.getFullYear()}`;

const uploadPhotoToFolder = async (siteId, token, folderId, fileName, dataUrl) => {
  const blob = await (await fetch(dataUrl)).blob();
  const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${folderId}:/${encodeURIComponent(fileName)}:/content`, {
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

// ── EKOTROPE ENERGY MODEL (.xml) PARSING ──────────────────────────────────────
// Parses the fields we've confirmed the schema for. Sections we haven't seen a
// real populated example of (e.g. foundation walls for basement/crawlspace units)
// are intentionally left unparsed rather than guessed — better no reference than
// a wrong one, given the whole point is catching discrepancies against the model.
function parseEkotropeXml(xmlText) {
  const doc = new DOMParser().parseFromString(xmlText, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("This file doesn't look like valid XML.");
  const building = doc.querySelector("buildingfile > building");
  if (!building) throw new Error("This doesn't look like an Ekotrope energy model export.");

  const txt = (el, sel) => { if (!el) return null; const n = el.querySelector(sel); const v = n?.textContent; return v ? v.trim() : null; };
  const num = (el, sel) => { const v = txt(el, sel); return v !== null && v !== "" && !isNaN(v) ? parseFloat(v) : null; };
  const bool = (el, sel) => txt(el, sel) === "true";

  // <notes> is modeler free text — often HTML-escaped (e.g. "&lt;div&gt;"), which the XML
  // parser decodes back into literal tag-looking text rather than real markup. Strip it.
  const rawNotes = txt(building, "notes");
  const notes = rawNotes ? rawNotes.replace(/<[^>]+>/g, "\n").replace(/\n{2,}/g, "\n").split("\n").map(s => s.trim()).filter(Boolean).join("\n") : null;

  const walls = Array.from(building.querySelectorAll("aboveGradeWalls > aboveGradeWall")).map(w => {
    const t = w.querySelector("aboveGradeWallType");
    return { name: txt(w, "name"), rCavity: num(t, "frameCavityInsRval"), rContinuous: num(t, "continousInsRval"), grade: txt(t, "cavityInsGrade") };
  });

  const ceilings = Array.from(building.querySelectorAll("roofs > roof")).map(r => {
    const t = r.querySelector("ceilingType");
    return { name: txt(r, "name") || txt(t, "name"), type: txt(t, "type"), rCavity: num(t, "cavityInsRval"), rContinuous: num(t, "continousInsRval"), grade: txt(t, "cavityInsGrade"), radiantBarrier: txt(r, "radiantBarrier"), exteriorColor: txt(r, "exteriorColor") };
  });

  const foundation = Array.from(building.querySelectorAll("slabs > slab")).map(s => {
    const t = s.querySelector("slabType");
    return { name: txt(s, "name") || txt(t, "name"), perimeterR: num(t, "perimeterInsRVal"), perimeterDepth: num(t, "perimeterInsDepth"), underslabR: num(t, "unSlabInsRVal"), underslabWidth: num(t, "unSlabInsWidth"), grade: txt(t, "insulGrade") };
  });

  const windowsSeen = new Map();
  building.querySelectorAll("windows > window").forEach(w => {
    const t = w.querySelector("windowType");
    const name = txt(t, "name") || txt(w, "name");
    if (name && !windowsSeen.has(name)) windowsSeen.set(name, { name, uValue: num(t, "uValue"), shgc: num(t, "sHGC") });
  });
  const windows = Array.from(windowsSeen.values());

  const doors = Array.from(building.querySelectorAll("doors > door")).map(d => {
    const t = d.querySelector("doorType");
    return { name: txt(d, "name") || txt(t, "name"), rValueOpaque: num(t, "rvalOpaque") };
  });

  const ductsSeen = new Map();
  building.querySelectorAll("ductSystems > ductSystem > duct").forEach(d => {
    const percentArea = num(d, "percentArea");
    if (!percentArea) return;
    const key = txt(d, "type") + "|" + txt(d, "location") + "|" + num(d, "RVal");
    if (!ductsSeen.has(key)) ductsSeen.set(key, { type: txt(d, "type"), location: txt(d, "location"), rVal: num(d, "RVal"), percentArea });
  });
  const ducts = Array.from(ductsSeen.values());

  const infil = building.querySelector("infiltration");
  const ventilation = infil ? { type: txt(infil, "mechVentType"), rateCfm: num(infil, "mechVentRate"), hoursPerDay: num(infil, "mechVentHoursPerDay"), fanWatts: num(infil, "mechVentFanWatts") } : null;

  // Equipment instances vary a lot by type (heat pump, furnace, boiler, water heater…) so
  // rather than guess every possible "*Type" child tag name, find whichever one is present.
  const hvac = [], waterHeaters = [];
  building.querySelectorAll("equipInfo > equipmentInstances > equipmentInstance").forEach(inst => {
    const libraryType = txt(inst, "libraryType");
    const typeChild = Array.from(inst.children).find(c => /Type$/.test(c.tagName) && c.tagName !== "libraryType");
    if (!typeChild) return;
    const entry = {
      libraryType, name: txt(typeChild, "name"), fuelType: txt(typeChild, "fuelType"),
      heatingEfficiency: num(typeChild, "heatingEfficiency"), coolingEfficiency: num(typeChild, "coolingEfficiency"),
      seasonalEqEff: num(typeChild, "seasonalEqEff"), effUnitType: txt(typeChild, "effUnitType"),
      ratedOutCapacity: num(typeChild, "ratedOutCapacity") || num(typeChild, "heatingRatedOutCapacity47"),
      energyFactor: num(typeChild, "energyFactor"), tankVolumeGallons: num(typeChild, "tankVolumeGallons"),
    };
    if (/water heat/i.test(libraryType || "")) waterHeaters.push(entry); else hvac.push(entry);
  });

  const dhwEl = building.querySelector("equipInfo > dhwDistribution");
  const dhw = dhwEl ? { allFixturesLowFlow: bool(dhwEl, "allFixturesLowFlow"), allDhwPipesInsulatedR3: bool(dhwEl, "allDhwPipesInsulatedR3") } : null;

  const la = building.querySelector("lightApps");
  const lighting = la ? { intLED: num(la, "percentIntLED"), extLED: num(la, "percentExtLED"), cfl: num(la, "percentCFL"), fluorescent: num(la, "percentFluorescent"), incandescent: num(la, "percentIntLight") } : null;
  const refrigerator = la && num(la, "refrigeratorKWH") ? { kwhYear: num(la, "refrigeratorKWH"), location: txt(la, "refrigeratorLocation") } : null;
  const dishwasher = la && num(la, "dishwasherkWhYear") ? { capacity: txt(la, "dishwasherCapacity"), kwhYear: num(la, "dishwasherkWhYear") } : null;
  const range = la ? { fuel: txt(la, "ovenFuel"), induction: bool(la, "inductionRange"), convection: bool(la, "convectionOven") } : null;
  const dryer = la ? { fuel: txt(la, "dryerFuel"), location: txt(la, "dryerLocation"), cef: num(la, "dryerCEF") } : null;
  const washer = la ? { location: txt(la, "washerLocation"), ler: num(la, "washerLER"), capacity: num(la, "washerCapacity") } : null;
  const ceilingFan = la && num(la, "ceilingFanCFMWatt") ? { cfmWatt: num(la, "ceilingFanCFMWatt") } : null;

  return { notes, walls, ceilings, foundation, windows, doors, ducts, ventilation, hvac, waterHeaters, dhw, lighting, refrigerator, dishwasher, range, dryer, washer, ceilingFan };
}

const fmtNum = (n, d) => n === null || n === undefined ? null : Number(n).toFixed(d === undefined ? 1 : d).replace(/\.0+$/, "");

// Maps an MRF item id to reference lines drawn from the parsed energy model.
// Items with no entry here have nothing comparable in an Ekotrope export.
const MRF_MODEL_FIELDS = {
  mrf_1_0: (m) => m.hvac.length ? m.hvac.map(h => `${h.libraryType}: ${h.name || ""}${h.heatingEfficiency ? ` — Htg ${fmtNum(h.heatingEfficiency)}` : ""}${h.coolingEfficiency ? ` / Clg ${fmtNum(h.coolingEfficiency)}` : ""}${h.seasonalEqEff ? ` ${fmtNum(h.seasonalEqEff)} ${h.effUnitType || ""}` : ""}`) : null,
  mrf_1_2: (m) => m.ventilation ? [`${m.ventilation.type || "Mechanical ventilation"} — ${fmtNum(m.ventilation.rateCfm, 0)} CFM, ${fmtNum(m.ventilation.hoursPerDay)} hrs/day, ${fmtNum(m.ventilation.fanWatts)}W`] : null,
  mrf_2_0: (m) => m.walls.length ? m.walls.map(w => `${w.name}: R-${fmtNum(w.rCavity)} cavity${w.rContinuous ? ` + R-${fmtNum(w.rContinuous)} continuous` : ""}, Grade ${w.grade || "?"}`) : null,
  mrf_2_1: (m) => m.ceilings.length ? m.ceilings.map(c => `${c.name}${c.type ? ` (${c.type})` : ""}: R-${fmtNum(c.rCavity)}${c.rContinuous ? ` + R-${fmtNum(c.rContinuous)} continuous` : ""}, Grade ${c.grade || "?"}`) : null,
  mrf_2_2: (m) => m.foundation.length ? m.foundation.map(f => `${f.name}: Perimeter R-${fmtNum(f.perimeterR)} (${fmtNum(f.perimeterDepth)} ft), Underslab R-${fmtNum(f.underslabR)} (${fmtNum(f.underslabWidth)} ft), Grade ${f.grade || "?"}`) : null,
  mrf_2_4: (m) => m.ducts.length ? m.ducts.map(d => `${d.type} duct, ${d.location}: R-${fmtNum(d.rVal)} (${fmtNum(d.percentArea, 0)}% of area)`) : null,
  mrf_2_5: (m) => m.windows.length ? m.windows.map(w => `${w.name}: U-${fmtNum(w.uValue, 2)}, SHGC ${fmtNum(w.shgc, 2)}`) : null,
  mrf_2_6: (m) => m.doors.length ? m.doors.map(d => `${d.name}: R-${fmtNum(d.rValueOpaque, 2)} opaque`) : null,
  mrf_2_7: (m) => m.ceilings.length ? m.ceilings.map(c => `${c.exteriorColor || "?"} color, radiant barrier: ${c.radiantBarrier || "?"}`) : null,
  mrf_3_0: (m) => m.waterHeaters.length ? m.waterHeaters.map(w => `${w.name || ""}: ${w.fuelType || ""}, ${fmtNum(w.energyFactor, 2)} EF/UEF, ${fmtNum(w.tankVolumeGallons, 0)} gal`) : null,
  mrf_3_1: (m) => m.dhw ? [`Model assumes hot water pipes insulated ≥R-3: ${m.dhw.allDhwPipesInsulatedR3 ? "Yes" : "No"}`] : null,
  mrf_3_2: (m) => m.dhw ? [`Model assumes all fixtures low-flow: ${m.dhw.allFixturesLowFlow ? "Yes" : "No"}`] : null,
  mrf_5_0: (m) => m.refrigerator ? [`${fmtNum(m.refrigerator.kwhYear, 0)} kWh/yr, ${m.refrigerator.location}`] : null,
  mrf_5_1: (m) => m.dishwasher ? [`${m.dishwasher.capacity} capacity, ${fmtNum(m.dishwasher.kwhYear, 0)} kWh/yr`] : null,
  mrf_5_2: (m) => m.range ? [`${m.range.fuel}${m.range.induction ? ", induction" : ""}${m.range.convection ? ", convection" : ""}`] : null,
  mrf_5_3: (m) => m.dryer ? [`${m.dryer.fuel}, ${m.dryer.location}, CEF ${fmtNum(m.dryer.cef, 2)}`] : null,
  mrf_5_4: (m) => m.washer ? [`${m.washer.location}, LER ${fmtNum(m.washer.ler, 0)}, ${fmtNum(m.washer.capacity, 1)} cu ft`] : null,
  mrf_5_5: (m) => m.ceilingFan ? [`${fmtNum(m.ceilingFan.cfmWatt, 1)} CFM/Watt`] : null,
  mrf_6_0: (m) => m.lighting ? [`LED ${fmtNum(m.lighting.intLED, 0)}%, CFL ${fmtNum(m.lighting.cfl, 0)}%, Fluorescent ${fmtNum(m.lighting.fluorescent, 0)}%, Incandescent ${fmtNum(m.lighting.incandescent, 0)}%`] : null,
  mrf_6_1: (m) => m.lighting ? [`LED ${fmtNum(m.lighting.extLED, 0)}%`] : null,
};

const EARTHCRAFT_CERTIFIED_V7 = [
  // SP 2.7 removed 2026-08-20: mandatory-mislabeled -- genuinely optional in the real V7
  // workbook, already correctly tracked as ec_opt_sp_sp_2_7 in EARTHCRAFT_OPTIONAL_LIBRARY.
  // ── SITE PLANNING ──────────────────────────────────────────────────────────
  { id: "ec_du1_5",  pointNumber: "DU 1.7",  tier: "ALL",  text: "Maintain 2\" clearance between wall siding and roof surface", category: "Durability & Moisture Management" },
  { id: "ec_du1_6",  pointNumber: "DU 1.8",  tier: "ALL",  text: "Install level air conditioner condensing unit pad", category: "Durability & Moisture Management" },
  { id: "ec_nc_du2_5", pointNumber: "DU 2.5",  tier: "ALL",  text: "Do not install wet or water-damaged building materials", category: "Durability & Moisture Management" },
  { id: "ec_du2_8",  pointNumber: "DU 2.8",  tier: "ALL",  text: "Design for or install additional dehumidification: rough-in electrical/plumbing for dehumidifier OR install whole-unit ENERGY STAR dehumidifier", category: "Durability & Moisture Management" },
  { id: "ec_iaq1",   pointNumber: "IAQ 1.0", tier: "ALL",  text: "No unvented combustion fireplaces, appliances, or space heaters; all combustion appliances mechanically drafted or direct-vented (EarthCraft IAQ 1 / Energy Star 10.1–10.3)", category: "Indoor Air Quality", mergedWith: ["es_10_1","es_10_2","es_10_3"] },
  { id: "ec_iaq1_2", pointNumber: "IAQ 1.2 / 1.3", tier: "ALL",  text: "Sealed-combustion or electric water heater installed within thermal envelope; no unit-level atmospherically vented water heaters or furnaces (EarthCraft IAQ 1.1–1.2 / Energy Star ES 5.0)", category: "Indoor Air Quality", mergedWith: ["es_10_1","ec_es5_0_ref"] },
  { id: "ec_iaq1_3", pointNumber: "IAQ 1.4", tier: "ALL",  text: "Carbon monoxide detector installed if combustion appliances exist (one per unit)", category: "Indoor Air Quality" },
  { id: "ec_iaq2",   pointNumber: "IAQ 2",   tier: "ALL",  text: "Protect all ducts and indoor coils until floor/wall finishing is complete", category: "Indoor Air Quality" },
  { id: "ec_iaq2_1", pointNumber: "IAQ 2.1", tier: "ALL",  text: "Filter is easily accessible for property maintenance; MERV 6+ minimum installed in each ducted system; all return and outdoor air passes through filter prior to distribution (EarthCraft IAQ 2.1–2.3 / Energy Star 9.1)", category: "Indoor Air Quality", mergedWith: ["es_9_1"] },
  { id: "ec_iaq2_5", pointNumber: "IAQ 2.5", tier: "ALL",  text: "No carpet in below-grade units", category: "Indoor Air Quality" },
  { id: "ec_es5_1",  pointNumber: "ES 5.1",  tier: "ALL",  text: "Heat trap on all storage water heaters; confirm presence by visual inspection or AHRI certificate (EarthCraft ES 5.1 / Energy Star 11.3)", category: "Energy Efficient Systems", mergedWith: ["es_11_3"] },
  { id: "ec_es5_3",  pointNumber: "ES 5.3",  tier: "ALL",  text: "Pipe insulation on first 2' of hot and cold water pipes at water heater", category: "Energy Efficient Systems" },
  { id: "ec_we1_0",  pointNumber: "WE 1.0",  tier: "ALL",  text: "Meet National Energy Policy Act low-flow standards for all fixtures", category: "Water Efficiency" },
  { id: "ec_we1_1",  pointNumber: "WE 1.1",  tier: "ALL",  text: "Detect no leaks at any water-using fixture, appliance or equipment", category: "Water Efficiency" },
  { id: "ec_we1_2",  pointNumber: "WE 1.2",  tier: "ALL",  text: "Low-flow fixtures throughout: WaterSense toilet ≤1.28 gpf; WaterSense urinal ≤0.5 gpf; WaterSense lavatory faucet ≤1.5 gpm; WaterSense showerhead ≤2.0 gpm (EarthCraft WE 1.2 / Energy Star 13.2)", category: "Water Efficiency", mergedWith: ["es_13_2"] },
  { id: "ec_du2_6",  pointNumber: "WE 2.3",  tier: "ALL",  text: "Newly installed and existing plants maintain distance ≥2' from building at maturity", category: "Water Efficiency" },
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
  // ── Added 2026-08-20 from the full V7 mandatory-checklist audit (see CLAUDE.md) ──────────────
  { id: "ec_v7_sp_sp_3_8", pointNumber: "SP 3.8", text: "Label all storm drains or storm inlets to discourage dumping of pollutants", category: "Site Planning", tier: "ALL" },
  { id: "ec_v7_sp_sp_3_9", pointNumber: "SP 3.9", text: "Road/vehicle cleaning protocols posted and enforced", category: "Site Planning", tier: "ALL" },
  { id: "ec_v7_re_re_1_0", pointNumber: "RE 1.0", text: "Limit framing at all windows and doors", category: "Resource Efficiency", tier: "ALL" },
  { id: "ec_v7_du_du_1", pointNumber: "DU 1", text: "All roof valleys direct water away from walls, dormers, chimneys, etc.", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_1", pointNumber: "DU 1.1", text: "Install drainage plane per manufacturer's specifications", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_2_1", pointNumber: "DU 1.2 > 1", text: "Integrate drainage plane with: > Window and door pan flashing at sills and side flashing", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_2_2", pointNumber: "DU 1.2 > 2", text: "Integrate drainage plane with: > Window and door head/top flashing", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_3", pointNumber: "DU 1.3", text: "Double layer of building paper or house wrap behind cementitious stucco, stone veneer or synthetic stone veneer on framed walls", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_4", pointNumber: "DU 1.4", text: "Roof gutters discharge water ≥5' from foundation", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_5_1", pointNumber: "DU 1.5 > 1", text: "Flashing: > Self-sealing bituminous membrane or equivalent at valleys and roof deck penetrations", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_5_2", pointNumber: "DU 1.5 > 2", text: "Flashing: > Step and kick-out flashing at wall/roof and wall/porch intersections, flashing ≥4” on wall surface and integrated with wall and roof/deck/porch drainage planes", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_2", pointNumber: "DU 2", text: "Gravel bed (57's, no fines) beneath sub-grade slabs, on grade slabs, or raised slabs", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_2_1", pointNumber: "DU 2.1", text: "100% coverage of ≥6mil vapor barrier beneath all slabs, in all crawlspaces", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_2_2", pointNumber: "DU 2.2", text: "Foundation drain on top of sub-grade footing", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_be_be_1_1", pointNumber: "BE 1.1", text: "Seal bottom plates to subfloor or foundation for entire unit envelope", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_5_1", pointNumber: "BE 1.5 > 1", text: "Seal penetrations through: > Foundations and exterior wall assemblies", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_5_2", pointNumber: "BE 1.5 > 2", text: "Seal penetrations through: > Top and bottom plates", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_5_5", pointNumber: "BE 1.5 > 5", text: "Seal penetrations through: > Sheathing", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_5_7", pointNumber: "BE 1.5 > 7", text: "Seal penetrations through: > All ceilings", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_1", pointNumber: "BE 1.6 > 1", text: "Seal penetrations around: > Shower, sinks, toilets and tub drains", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_2", pointNumber: "BE 1.6 > 2", text: "Seal penetrations around: > HVAC supply and return boots sealed to subfloor or drywall (floor, walls, or ceilings)", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_3", pointNumber: "BE 1.6 > 3", text: "Seal penetrations around: > Window and door rough openings", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_4", pointNumber: "BE 1.6 > 4", text: "Seal penetrations around: > All drywall penetrations (common walls between attached units included)", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_5", pointNumber: "BE 1.6 > 5", text: "Seal penetrations around: > Exhaust fans to drywall", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_6", pointNumber: "BE 1.6 > 6", text: "Seal penetrations around: > Attic pull-down stairs, scuttle holes and kneewall doors", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_7", pointNumber: "BE 1.6 > 7", text: "Seal penetrations around: > Chases", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_7_1", pointNumber: "BE 1.7 > 1", text: "Seal seams and gaps in: > Band joist sheathing", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_7_2", pointNumber: "BE 1.7 > 2", text: "Seal seams and gaps in: > Exterior wall sheathing", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_8_1", pointNumber: "BE 1.8 > 1", text: "Install rigid air barriers: > Behind tubs and showers on insulated walls", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_9_1", pointNumber: "BE 1.9 > 1", text: "Install weather-stripping at: > All exterior doors (if not included in door assembly)", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_10", pointNumber: "BE 1.10", text: "All recessed can lights must be air tight, gasketed at all floors and also IC-rated in insulated ceilings;  in Climate Zone 4, insulate exterior surface of fixture to ≥R-10", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_12", pointNumber: "BE 1.12", text: "Units adjacent to CMU walls: framing and sub-floor at unit envelope, including interstitial space, must be sealed to CMU", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_13", pointNumber: "BE 1.13", text: "Seal top plate to drywall at the attic level", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_1", pointNumber: "BE 3 > 1", text: "Floors: > Framed ≥ R-19", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_2", pointNumber: "BE 3 > 2", text: "Floors: > Cantilevered ≥ R-30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_3", pointNumber: "BE 3 > 3", text: "Floors: > Podium/Elevated Slab ≥ R-19", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_1_1", pointNumber: "BE 3.1 > 1", text: "Walls: > Exterior walls and band joists ≥ R-15", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_1_2", pointNumber: "BE 3.1 > 2", text: "Walls: > Elevator walls adjacent to dwelling units ≥ R-13", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_1_3", pointNumber: "BE 3.1 > 3", text: "Walls: > Foundation walls ≥ R-10 continuous or ≥ R-13 cavity; Climate Zone 2/3 ≥ R-5 continuous or ≥ R-13 cavity; Climate Zone 4 ≥ R-10 continuous or ≥ R-13 cavity", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_2_1", pointNumber: "BE 3.2 > 1", text: "Ceilings/Roof: > Vented: Climate Zone 4 ≥ R-49", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_2_2", pointNumber: "BE 3.2 > 2", text: "Ceilings/Roof: > Continuous Roof Deck: Climate Zone 4 ≥ R-30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_2_3", pointNumber: "BE 3.2 > 3", text: "Ceilings/Roof: > Cathedral: Climate Zone 4 ≥ R-38", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_3_1", pointNumber: "BE 3.3 > 1", text: "Attic/Roof: > Install wind baffles at eaves in every vented bay, or equivalent air barrier at edge of ceiling", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_3_2", pointNumber: "BE 3.3 > 2", text: "Attic/Roof: > Energy heel trusses or raised top plate", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_3_3", pointNumber: "BE 3.3 > 3", text: "Attic/Roof: > Attic platforms allow for full-depth insulation below", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_5", pointNumber: "BE 3.5", text: "Attic pull-down/scuttle hole ≥ R-49", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_7", pointNumber: "BE 3.7", text: "Steel framed buildings require thermal break ≥ R-7.5", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_1", pointNumber: "BE 4 > 1", text: "Door U-factors and SHGC: > U-factor ≤0.35", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_2", pointNumber: "BE 4 > 2", text: "Door U-factors and SHGC: > SHGC ≤ 0.30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_1_1", pointNumber: "BE 4.1 > 1", text: "Window U-factor and SHGC: > U-factor ≤0.30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_1_2", pointNumber: "BE 4.1 > 2", text: "Window U-factor and SHGC: > SHGC ≤ 0.30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_2_1", pointNumber: "BE 4.2 > 1", text: "Skylight U-factor and SHGC: > U-factor ≤0.55", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_2_2", pointNumber: "BE 4.2 > 2", text: "Skylight U-factor and SHGC: > SHGC ≤ 0.30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_es_es_1_1", pointNumber: "ES 1.1", text: "If programmable thermostat installed for heat pump, include adaptive recovery technology", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2", pointNumber: "ES 2", text: "Seal air handlers and duct systems with mastic", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_1", pointNumber: "ES 2.1", text: "Install ducts per ACCA Manual D duct design", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_2", pointNumber: "ES 2.2", text: "Fully duct all supply and return ducts", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_3_1", pointNumber: "ES 2.3 > 1", text: "Duct insulation: > ≥ R-6: Ducts in conditioned and interstitial spaces (between floors)", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_3_2", pointNumber: "ES 2.3 > 2", text: "Duct insulation: > ≥ R-8: Ducts in unconditioned space", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_4", pointNumber: "ES 2.4", text: "No ducts in exterior walls or vaulted ceilings and no plenum within 2' of roofline.", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_5", pointNumber: "ES 2.5", text: "Locate all air handlers within conditioned space", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_6", pointNumber: "ES 2.6", text: "Indoor coil protected until finished floor installed", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_8", pointNumber: "ES 2.8", text: "No duct take-offs within 6\" of supply plenum or supply trunk cap", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_9", pointNumber: "ES 2.9", text: "Design and construct mechanical closets accessible for service and maintenance requirements", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4", pointNumber: "ES 4", text: "Install exhaust fans in all bathrooms and duct to outside", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_1", pointNumber: "ES 4.1", text: "Gas kitchen range vented to exterior  ≥100 cfm fan", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_3_2", pointNumber: "ES 4.3 > 2", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > ≥ 2' above grade", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_3_4", pointNumber: "ES 4.3 > 4", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > Fresh air duct may not be run to the roof", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_3_5", pointNumber: "ES 4.3 > 5", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > Fresh air shutoff may not be controlled by humidistat", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_3_6", pointNumber: "ES 4.3 > 6", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > Install rigid duct with insulation", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_3_7", pointNumber: "ES 4.3 > 7", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > All intakes must be ducted to exterior of building", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_4", pointNumber: "ES 4.4", text: "Seal seams of all intake and exhaust ducts with mastic", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_5", pointNumber: "ES 4.5", text: "Duct clothes dryers to outside", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_7", pointNumber: "ES 4.7", text: "Back-draft dampers for kitchen and bathroom exhaust", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_5_0", pointNumber: "ES 5.0", text: "Water Heater must be installed in conditioned space. If gas, direct vent", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_5_1", pointNumber: "ES 5.1", text: "Heat trap on all storage water heaters", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_5_3", pointNumber: "ES 5.3", text: "Pipe insulation on first 2'", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_6", pointNumber: "ES 6", text: "High-efficacy lighting in 100% of all permanent fixtures", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_6_1", pointNumber: "ES 6.1", text: "If installed, ENERGY STAR dishwasher", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_6_2", pointNumber: "ES 6.2", text: "If installed, ENERGY STAR refrigerator", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_7", pointNumber: "ES 7", text: "100% LED bulbs in all corridor/breezeway and all common spaces", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_we_we_2", pointNumber: "WE 2", text: "Cover all exposed soil with 2\"-3\" mulch layer", category: "Water Efficiency", tier: "ALL" },
  { id: "ec_v7_we_we_2_1_1", pointNumber: "WE 2.1 > 1", text: "Irrigation system: > Must have rain sensor shutoff switch", category: "Water Efficiency", tier: "ALL" }
];

const EARTHCRAFT_GOLD_V7 = [
  { id: "ec_du1_5",  pointNumber: "DU 1.7",  tier: "ALL",  text: "Maintain 2\" clearance between wall siding and roof surface", category: "Durability & Moisture Management" },
  { id: "ec_du1_6",  pointNumber: "DU 1.8",  tier: "ALL",  text: "Install level air conditioner condensing unit pad", category: "Durability & Moisture Management" },
  { id: "ec_nc_du2_5", pointNumber: "DU 2.5",  tier: "ALL",  text: "Do not install wet or water-damaged building materials", category: "Durability & Moisture Management" },
  { id: "ec_du2_8",  pointNumber: "DU 2.8",  tier: "ALL",  text: "Design for or install additional dehumidification: rough-in electrical/plumbing for dehumidifier OR install whole-unit ENERGY STAR dehumidifier", category: "Durability & Moisture Management" },
  { id: "ec_iaq1",   pointNumber: "IAQ 1.0", tier: "ALL",  text: "No unvented combustion fireplaces, appliances, or space heaters; all combustion appliances mechanically drafted or direct-vented (EarthCraft IAQ 1 / Energy Star 10.1–10.3)", category: "Indoor Air Quality", mergedWith: ["es_10_1","es_10_2","es_10_3"] },
  { id: "ec_iaq1_2", pointNumber: "IAQ 1.2 / 1.3", tier: "ALL",  text: "Sealed-combustion or electric water heater installed within thermal envelope; no unit-level atmospherically vented water heaters or furnaces (EarthCraft IAQ 1.1–1.2 / Energy Star ES 5.0)", category: "Indoor Air Quality", mergedWith: ["es_10_1","ec_es5_0_ref"] },
  { id: "ec_iaq1_3", pointNumber: "IAQ 1.4", tier: "ALL",  text: "Carbon monoxide detector installed if combustion appliances exist (one per unit)", category: "Indoor Air Quality" },
  { id: "ec_iaq2",   pointNumber: "IAQ 2",   tier: "ALL",  text: "Protect all ducts and indoor coils until floor/wall finishing is complete", category: "Indoor Air Quality" },
  { id: "ec_iaq2_1", pointNumber: "IAQ 2.1", tier: "ALL",  text: "Filter is easily accessible for property maintenance; MERV 6+ minimum installed in each ducted system; all return and outdoor air passes through filter prior to distribution (EarthCraft IAQ 2.1–2.3 / Energy Star 9.1)", category: "Indoor Air Quality", mergedWith: ["es_9_1"] },
  { id: "ec_iaq2_5", pointNumber: "IAQ 2.5", tier: "ALL",  text: "No carpet in below-grade units", category: "Indoor Air Quality" },
  { id: "ec_es5_1",  pointNumber: "ES 5.1",  tier: "ALL",  text: "Heat trap on all storage water heaters; confirm presence by visual inspection or AHRI certificate (EarthCraft ES 5.1 / Energy Star 11.3)", category: "Energy Efficient Systems", mergedWith: ["es_11_3"] },
  { id: "ec_es5_3",  pointNumber: "ES 5.3",  tier: "ALL",  text: "Pipe insulation on first 2' of hot and cold water pipes at water heater", category: "Energy Efficient Systems" },
  { id: "ec_we1_0",  pointNumber: "WE 1.0",  tier: "ALL",  text: "Meet National Energy Policy Act low-flow standards for all fixtures", category: "Water Efficiency" },
  { id: "ec_we1_1",  pointNumber: "WE 1.1",  tier: "ALL",  text: "Detect no leaks at any water-using fixture, appliance or equipment", category: "Water Efficiency" },
  { id: "ec_we1_2",  pointNumber: "WE 1.2",  tier: "ALL",  text: "Low-flow fixtures throughout: WaterSense toilet ≤1.28 gpf; WaterSense urinal ≤0.5 gpf; WaterSense lavatory faucet ≤1.5 gpm; WaterSense showerhead ≤2.0 gpm (EarthCraft WE 1.2 / Energy Star 13.2)", category: "Water Efficiency", mergedWith: ["es_13_2"] },
  { id: "ec_du2_6",  pointNumber: "WE 2.3",  tier: "ALL",  text: "Newly installed and existing plants maintain distance ≥2' from building at maturity", category: "Water Efficiency" },
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
  { id: "ec_du2_7",  pointNumber: "DU 2.10", tier: "GOLD", text: "If installed, drain at outside perimeter edge of footing surrounded with 6\" clearstone and filter fabric", category: "Durability & Moisture Management" },
  { id: "ec_iaq2_6", pointNumber: "IAQ 2.6", tier: "GOLD", text: "Filters are ≥ MERV 8", category: "Indoor Air Quality" },
  { id: "ec_v7_re1_2_1", pointNumber: "RE 1.2 > 1", text: "2-stud corners where structurally feasible", category: "Resource Efficiency", tier: "GOLD", points: 3 },
  { id: "ec_v7_re1_2_2", pointNumber: "RE 1.2 > 2", text: "Ladder T-walls where structurally feasible", category: "Resource Efficiency", tier: "GOLD", points: 2 },
  { id: "ec_v7_re1_2_3", pointNumber: "RE 1.2 > 3", text: "Size headers for loads (non-structural headers in non-load bearing walls)", category: "Resource Efficiency", tier: "GOLD", points: 1 },
  { id: "ec_v7_du2_9", pointNumber: "DU 2.9", text: "Additional dehumidification system installed: basement or sealed crawlspace system", category: "Durability & Moisture Management", tier: "GOLD" },
  { id: "ec_v7_iaq1_5", pointNumber: "IAQ 1.5", text: "If installed, all fireplaces meet indoor air quality guidelines and have gasketed doors", category: "Indoor Air Quality", tier: "GOLD" },
  { id: "ec_v7_be1_14", pointNumber: "BE 1.14", text: "Top plate sealed to drywall at all levels", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be3_11", pointNumber: "BE 3.11", text: "Corners insulated to ≥ R-6", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be3_12", pointNumber: "BE 3.12", text: "Headers insulated to ≥ R-3", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be3_13", pointNumber: "BE 3.13", text: "Fiberglass batts are unfaced and friction-fit throughout", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be5_0", pointNumber: "BE 5.0", text: "Ducts in unconditioned attic: buried in R-49 insulation OR ducts with R-8 insulation encapsulated in 1.5\" closed-cell foam and buried under ≥2\" blown insulation", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_es1_11", pointNumber: "ES 1.11", text: "HVAC equipment is ENERGY STAR qualified; for split systems, the pairing must be qualified", category: "Energy Efficient Systems", tier: "GOLD" },
  // ── Added 2026-08-20 from the full V7 mandatory-checklist audit (see CLAUDE.md) ──────────────
  { id: "ec_v7_sp_sp_3_8", pointNumber: "SP 3.8", text: "Label all storm drains or storm inlets to discourage dumping of pollutants", category: "Site Planning", tier: "ALL" },
  { id: "ec_v7_sp_sp_3_9", pointNumber: "SP 3.9", text: "Road/vehicle cleaning protocols posted and enforced", category: "Site Planning", tier: "ALL" },
  { id: "ec_v7_re_re_1_0", pointNumber: "RE 1.0", text: "Limit framing at all windows and doors", category: "Resource Efficiency", tier: "ALL" },
  { id: "ec_v7_du_du_1", pointNumber: "DU 1", text: "All roof valleys direct water away from walls, dormers, chimneys, etc.", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_1", pointNumber: "DU 1.1", text: "Install drainage plane per manufacturer's specifications", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_2_1", pointNumber: "DU 1.2 > 1", text: "Integrate drainage plane with: > Window and door pan flashing at sills and side flashing", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_2_2", pointNumber: "DU 1.2 > 2", text: "Integrate drainage plane with: > Window and door head/top flashing", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_3", pointNumber: "DU 1.3", text: "Double layer of building paper or house wrap behind cementitious stucco, stone veneer or synthetic stone veneer on framed walls", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_4", pointNumber: "DU 1.4", text: "Roof gutters discharge water ≥5' from foundation", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_5_1", pointNumber: "DU 1.5 > 1", text: "Flashing: > Self-sealing bituminous membrane or equivalent at valleys and roof deck penetrations", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_1_5_2", pointNumber: "DU 1.5 > 2", text: "Flashing: > Step and kick-out flashing at wall/roof and wall/porch intersections, flashing ≥4” on wall surface and integrated with wall and roof/deck/porch drainage planes", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_2", pointNumber: "DU 2", text: "Gravel bed (57's, no fines) beneath sub-grade slabs, on grade slabs, or raised slabs", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_2_1", pointNumber: "DU 2.1", text: "100% coverage of ≥6mil vapor barrier beneath all slabs, in all crawlspaces", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_du_du_2_2", pointNumber: "DU 2.2", text: "Foundation drain on top of sub-grade footing", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v7_be_be_1_1", pointNumber: "BE 1.1", text: "Seal bottom plates to subfloor or foundation for entire unit envelope", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_5_1", pointNumber: "BE 1.5 > 1", text: "Seal penetrations through: > Foundations and exterior wall assemblies", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_5_2", pointNumber: "BE 1.5 > 2", text: "Seal penetrations through: > Top and bottom plates", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_5_5", pointNumber: "BE 1.5 > 5", text: "Seal penetrations through: > Sheathing", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_5_7", pointNumber: "BE 1.5 > 7", text: "Seal penetrations through: > All ceilings", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_1", pointNumber: "BE 1.6 > 1", text: "Seal penetrations around: > Shower, sinks, toilets and tub drains", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_2", pointNumber: "BE 1.6 > 2", text: "Seal penetrations around: > HVAC supply and return boots sealed to subfloor or drywall (floor, walls, or ceilings)", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_3", pointNumber: "BE 1.6 > 3", text: "Seal penetrations around: > Window and door rough openings", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_4", pointNumber: "BE 1.6 > 4", text: "Seal penetrations around: > All drywall penetrations (common walls between attached units included)", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_5", pointNumber: "BE 1.6 > 5", text: "Seal penetrations around: > Exhaust fans to drywall", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_6", pointNumber: "BE 1.6 > 6", text: "Seal penetrations around: > Attic pull-down stairs, scuttle holes and kneewall doors", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_6_7", pointNumber: "BE 1.6 > 7", text: "Seal penetrations around: > Chases", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_7_1", pointNumber: "BE 1.7 > 1", text: "Seal seams and gaps in: > Band joist sheathing", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_7_2", pointNumber: "BE 1.7 > 2", text: "Seal seams and gaps in: > Exterior wall sheathing", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_8_1", pointNumber: "BE 1.8 > 1", text: "Install rigid air barriers: > Behind tubs and showers on insulated walls", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_9_1", pointNumber: "BE 1.9 > 1", text: "Install weather-stripping at: > All exterior doors (if not included in door assembly)", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_10", pointNumber: "BE 1.10", text: "All recessed can lights must be air tight, gasketed at all floors and also IC-rated in insulated ceilings;  in Climate Zone 4, insulate exterior surface of fixture to ≥R-10", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_12", pointNumber: "BE 1.12", text: "Units adjacent to CMU walls: framing and sub-floor at unit envelope, including interstitial space, must be sealed to CMU", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_1_13", pointNumber: "BE 1.13", text: "Seal top plate to drywall at the attic level", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_1", pointNumber: "BE 3 > 1", text: "Floors: > Framed ≥ R-19", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_2", pointNumber: "BE 3 > 2", text: "Floors: > Cantilevered ≥ R-30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_3", pointNumber: "BE 3 > 3", text: "Floors: > Podium/Elevated Slab ≥ R-19", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_1_1", pointNumber: "BE 3.1 > 1", text: "Walls: > Exterior walls and band joists ≥ R-15", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_1_2", pointNumber: "BE 3.1 > 2", text: "Walls: > Elevator walls adjacent to dwelling units ≥ R-13", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_1_3", pointNumber: "BE 3.1 > 3", text: "Walls: > Foundation walls ≥ R-10 continuous or ≥ R-13 cavity; Climate Zone 2/3 ≥ R-5 continuous or ≥ R-13 cavity; Climate Zone 4 ≥ R-10 continuous or ≥ R-13 cavity", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_2_1", pointNumber: "BE 3.2 > 1", text: "Ceilings/Roof: > Vented: Climate Zone 4 ≥ R-49", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_2_2", pointNumber: "BE 3.2 > 2", text: "Ceilings/Roof: > Continuous Roof Deck: Climate Zone 4 ≥ R-30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_2_3", pointNumber: "BE 3.2 > 3", text: "Ceilings/Roof: > Cathedral: Climate Zone 4 ≥ R-38", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_3_1", pointNumber: "BE 3.3 > 1", text: "Attic/Roof: > Install wind baffles at eaves in every vented bay, or equivalent air barrier at edge of ceiling", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_3_2", pointNumber: "BE 3.3 > 2", text: "Attic/Roof: > Energy heel trusses or raised top plate", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_3_3", pointNumber: "BE 3.3 > 3", text: "Attic/Roof: > Attic platforms allow for full-depth insulation below", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_5", pointNumber: "BE 3.5", text: "Attic pull-down/scuttle hole ≥ R-49", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_3_7", pointNumber: "BE 3.7", text: "Steel framed buildings require thermal break ≥ R-7.5", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_1", pointNumber: "BE 4 > 1", text: "Door U-factors and SHGC: > U-factor ≤0.35", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_2", pointNumber: "BE 4 > 2", text: "Door U-factors and SHGC: > SHGC ≤ 0.30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_1_1", pointNumber: "BE 4.1 > 1", text: "Window U-factor and SHGC: > U-factor ≤0.30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_1_2", pointNumber: "BE 4.1 > 2", text: "Window U-factor and SHGC: > SHGC ≤ 0.30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_2_1", pointNumber: "BE 4.2 > 1", text: "Skylight U-factor and SHGC: > U-factor ≤0.55", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_be_be_4_2_2", pointNumber: "BE 4.2 > 2", text: "Skylight U-factor and SHGC: > SHGC ≤ 0.30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v7_es_es_1_1", pointNumber: "ES 1.1", text: "If programmable thermostat installed for heat pump, include adaptive recovery technology", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2", pointNumber: "ES 2", text: "Seal air handlers and duct systems with mastic", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_1", pointNumber: "ES 2.1", text: "Install ducts per ACCA Manual D duct design", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_2", pointNumber: "ES 2.2", text: "Fully duct all supply and return ducts", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_3_1", pointNumber: "ES 2.3 > 1", text: "Duct insulation: > ≥ R-6: Ducts in conditioned and interstitial spaces (between floors)", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_3_2", pointNumber: "ES 2.3 > 2", text: "Duct insulation: > ≥ R-8: Ducts in unconditioned space", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_4", pointNumber: "ES 2.4", text: "No ducts in exterior walls or vaulted ceilings and no plenum within 2' of roofline.", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_5", pointNumber: "ES 2.5", text: "Locate all air handlers within conditioned space", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_6", pointNumber: "ES 2.6", text: "Indoor coil protected until finished floor installed", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_8", pointNumber: "ES 2.8", text: "No duct take-offs within 6\" of supply plenum or supply trunk cap", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_2_9", pointNumber: "ES 2.9", text: "Design and construct mechanical closets accessible for service and maintenance requirements", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4", pointNumber: "ES 4", text: "Install exhaust fans in all bathrooms and duct to outside", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_1", pointNumber: "ES 4.1", text: "Gas kitchen range vented to exterior  ≥100 cfm fan", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_3_2", pointNumber: "ES 4.3 > 2", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > ≥ 2' above grade", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_3_4", pointNumber: "ES 4.3 > 4", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > Fresh air duct may not be run to the roof", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_3_5", pointNumber: "ES 4.3 > 5", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > Fresh air shutoff may not be controlled by humidistat", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_3_6", pointNumber: "ES 4.3 > 6", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > Install rigid duct with insulation", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_3_7", pointNumber: "ES 4.3 > 7", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > All intakes must be ducted to exterior of building", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_4", pointNumber: "ES 4.4", text: "Seal seams of all intake and exhaust ducts with mastic", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_5", pointNumber: "ES 4.5", text: "Duct clothes dryers to outside", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_4_7", pointNumber: "ES 4.7", text: "Back-draft dampers for kitchen and bathroom exhaust", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_5_0", pointNumber: "ES 5.0", text: "Water Heater must be installed in conditioned space. If gas, direct vent", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_5_1", pointNumber: "ES 5.1", text: "Heat trap on all storage water heaters", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_5_3", pointNumber: "ES 5.3", text: "Pipe insulation on first 2'", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_6", pointNumber: "ES 6", text: "High-efficacy lighting in 100% of all permanent fixtures", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_6_1", pointNumber: "ES 6.1", text: "If installed, ENERGY STAR dishwasher", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_6_2", pointNumber: "ES 6.2", text: "If installed, ENERGY STAR refrigerator", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_es_es_7", pointNumber: "ES 7", text: "100% LED bulbs in all corridor/breezeway and all common spaces", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v7_we_we_2", pointNumber: "WE 2", text: "Cover all exposed soil with 2\"-3\" mulch layer", category: "Water Efficiency", tier: "ALL" },
  { id: "ec_v7_we_we_2_1_1", pointNumber: "WE 2.1 > 1", text: "Irrigation system: > Must have rain sensor shutoff switch", category: "Water Efficiency", tier: "ALL" },
  { id: "ec_v7_be_be_4_4_1", pointNumber: "BE 4.4 > 1", text: "Door U-factor: > Opaque door:  U factor≤ 0.17", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be_be_4_4_2", pointNumber: "BE 4.4 > 2", text: "Door U-factor: > Door with ≤ 50% glass:  U-factor ≤ 0.23", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be_be_4_4_3", pointNumber: "BE 4.4 > 3", text: "Door U-factor: > Door with > 50% glass:  U-factor ≤ 0.26", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be_be_4_5_1", pointNumber: "BE 4.5 > 1", text: "Window U-factor and SHGC: > U-factor ≤0.25 or ENERGY STAR labeled window", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be_be_4_5_2", pointNumber: "BE 4.5 > 2", text: "Window U-factor and SHGC: > SHGC ≤0.27", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be_be_4_6_1", pointNumber: "BE 4.6 > 1", text: "Skylight U-factor and SHGC: > U-factor ≤0.50", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_be_be_4_6_2", pointNumber: "BE 4.6 > 2", text: "Skylight U-factor and SHGC: > SHGC ≤0.25", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v7_es_es_2_11_1", pointNumber: "ES 2.11 > 1", text: "Minimize pressure imbalance within units: > Install fully ducted jumper ducts, transfer grills, or dedicated return for each bedroom", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v7_es_es_2_12", pointNumber: "ES 2.12", text: "Install rigid duct work or pull all flex ducts with no pinches and support at intervals ≤ 5’", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v7_es_es_2_15", pointNumber: "ES 2.15", text: "HVAC system and ductwork is dry and clean", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v7_es_es_1_8", pointNumber: "ES 1.8", text: "Heating equipment efficiency: ENERGY STAR qualified furnace(s) ≥95 AFUE and within 40% of load calculation, OR ENERGY STAR qualified heat pump(s) ≥8.5 HSPF and within 25% of load calculation", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v7_es_es_1_10", pointNumber: "ES 1.10", text: "ENERGY STAR qualified cooling equipment ≥SEER 15", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v7_es_es_4_9", pointNumber: "ES 4.9", text: "If installed, ceiling fans must be ENERGY STAR qualified (1/bedroom and 1 in living room)", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v7_es_es_4_12", pointNumber: "ES 4.12", text: "Install and label accessible ventilation controls, with override controls for continuously operating ventilation fans", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v7_es_es_4_13", pointNumber: "ES 4.13", text: "Supply/exhaust fans rated at ≤3 sones (intermittent) and ≤1 sone (continuous)", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v7_es_es_6_3", pointNumber: "ES 6.3", text: "If installed, ENERGY STAR qualified clothes washer", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v7_es_es_6_4", pointNumber: "ES 6.4", text: "If installed, high efficiency clothes dryer with moisture sensor (not applicable to commercial dryers)", category: "Energy Efficient Systems", tier: "GOLD" }
];

// ─── EARTHCRAFT OPTIONAL POINTS — WORKBOOK IMPORT ─────────────────────────────
// Curated by TA review of the full EarthCraft Multifamily optional-points list, Aug 2026
// (159 items kept from a 273-item V6.5 review + 3 V7-only additions: DU 2.16, IAQ 2.14, ES 1.21).
// Matched by normalized item description (matchKey) rather than workbook code/version,
// because EarthCraft renumbers items between versions (V6.5 -> V7) while descriptions stay
// stable. Known limitation: if a version also reworded a threshold (e.g. a SHGC cutoff
// tightening from 0.27 to 0.25), the text no longer matches and the item won't be recognized
// until it's added again under its new wording — accepted tradeoff, flagged rather than
// silently guessed at.
const EARTHCRAFT_OPTIONAL_LIBRARY = [
  { id: "ec_opt_sp_2_1", pointNumber: "2.1", text: "Shade at least 50% of hardscape within 30' of building", category: "Site Planning", points: 2, matchKey: "SP|shadeatleast50ofhardscapewithin30ofbuilding" },
  { id: "ec_opt_sp_sp_2_2", pointNumber: "SP 2.2", text: "Reduce light pollution - all exterior lights full cutoff", category: "Site Planning", points: 4, matchKey: "SP|reducelightpollutionallexteriorlightsfullcutoff" },
  { id: "ec_opt_sp_sp_2_4", pointNumber: "SP 2.4", text: "Street Trees are ≤ 40' on center at minimum", category: "Site Planning", points: 1, matchKey: "SP|streettreesare40oncenteratminimum" },
  { id: "ec_opt_sp_sp_2_5_1", pointNumber: "SP 2.5 > 1", text: "Connectivity to adjacent sites: > Vehicular access (2+ connections)", category: "Site Planning", points: 1, matchKey: "SP|connectivitytoadjacentsitesvehicularaccess2connections" },
  { id: "ec_opt_sp_sp_2_5_2", pointNumber: "SP 2.5 > 2", text: "Connectivity to adjacent sites: > Dedicated pedestrian and bike access", category: "Site Planning", points: 1, matchKey: "SP|connectivitytoadjacentsitesdedicatedpedestrianandbikeaccess" },
  { id: "ec_opt_sp_sp_2_6", pointNumber: "SP 2.6", text: "Community Gardens", category: "Site Planning", points: 1, matchKey: "SP|communitygardens" },
  { id: "ec_opt_sp_sp_2_7", pointNumber: "SP 2.7", text: "Outdoor Community gathering space", category: "Site Planning", points: 2, matchKey: "SP|outdoorcommunitygatheringspace" },
  { id: "ec_opt_sp_sp_2_8", pointNumber: "SP 2.8", text: "Install plant species that serve as pollinators on site for regional wildlife and/or local endangered species for a minimum of 20% of plantings", category: "Site Planning", points: 1, matchKey: "SP|installplantspeciesthatserveaspollinatorsonsiteforregionalwildlifeandorlocalendangeredspeciesforaminimumof20ofplantings" },
  { id: "ec_opt_sp_sp_3_10", pointNumber: "SP 3.10", text: "Tree preservation and protection measures employed on site", category: "Site Planning", points: 2, matchKey: "SP|treepreservationandprotectionmeasuresemployedonsite" },
  { id: "ec_opt_sp_sp_3_11", pointNumber: "SP 3.11", text: "Leave site undisturbed and protect greenspace from future development (min 25%)", category: "Site Planning", points: 2, matchKey: "SP|leavesiteundisturbedandprotectgreenspacefromfuturedevelopmentmin25" },
  { id: "ec_opt_sp_sp_3_12", pointNumber: "SP 3.12", text: "Mill cleared logs (100%)", category: "Site Planning", points: 1, matchKey: "SP|millclearedlogs100" },
  { id: "ec_opt_sp_sp_3_13", pointNumber: "SP 3.13", text: "Grind stumps and limbs for mulch (≥80%)", category: "Site Planning", points: 1, matchKey: "SP|grindstumpsandlimbsformulch80" },
  { id: "ec_opt_sp_sp_3_14", pointNumber: "SP 3.14", text: "Tree planting (12 trees per acre; trees ≥ 2\" diameter)", category: "Site Planning", points: 2, matchKey: "SP|treeplanting12treesperacretrees2diameter" },
  { id: "ec_opt_sp_sp_4_0", pointNumber: "SP 4.0", text: "Bike racks", category: "Site Planning", points: 1, matchKey: "SP|bikeracks" },
  { id: "ec_opt_sp_sp_4_1", pointNumber: "SP 4.1", text: "Covered bike storage facility", category: "Site Planning", points: 1, matchKey: "SP|coveredbikestoragefacility" },
  { id: "ec_opt_sp_sp_4_2", pointNumber: "SP 4.2", text: "Tenant access to business center", category: "Site Planning", points: 1, matchKey: "SP|tenantaccesstobusinesscenter" },
  { id: "ec_opt_sp_sp_4_3", pointNumber: "SP 4.3", text: "Covered bus stop", category: "Site Planning", points: 2, matchKey: "SP|coveredbusstop" },
  { id: "ec_opt_sp_sp_4_4", pointNumber: "SP 4.4", text: "Electric vehicle charging facility", category: "Site Planning", points: 2, matchKey: "SP|electricvehiclechargingfacility" },
  { id: "ec_opt_cw_cw_1_2_1", pointNumber: "CW 1.2 > 1", text: "Post waste management plan and divert 75% from landfill of: > Wood", category: "Construction Waste Management", points: 2, matchKey: "CW|postwastemanagementplananddivert75fromlandfillofwood" },
  { id: "ec_opt_cw_cw_1_2_2", pointNumber: "CW 1.2 > 2", text: "Post waste management plan and divert 75% from landfill of: > Cardboard", category: "Construction Waste Management", points: 1, matchKey: "CW|postwastemanagementplananddivert75fromlandfillofcardboard" },
  { id: "ec_opt_cw_cw_1_2_3", pointNumber: "CW 1.2 > 3", text: "Post waste management plan and divert 75% from landfill of: > Metal (including beverage containers)", category: "Construction Waste Management", points: 1, matchKey: "CW|postwastemanagementplananddivert75fromlandfillofmetalincludingbeveragecontainers" },
  { id: "ec_opt_cw_cw_1_2_4", pointNumber: "CW 1.2 > 4", text: "Post waste management plan and divert 75% from landfill of: > Drywall (recycle or grind and spread on site)", category: "Construction Waste Management", points: 2, matchKey: "CW|postwastemanagementplananddivert75fromlandfillofdrywallrecycleorgrindandspreadonsite" },
  { id: "ec_opt_cw_cw_1_2_5", pointNumber: "CW 1.2 > 5", text: "Post waste management plan and divert 75% from landfill of: > Plastic (including beverage containers)", category: "Construction Waste Management", points: 1, matchKey: "CW|postwastemanagementplananddivert75fromlandfillofplasticincludingbeveragecontainers" },
  { id: "ec_opt_cw_cw_1_2_6", pointNumber: "CW 1.2 > 6", text: "Post waste management plan and divert 75% from landfill of: > Shingles", category: "Construction Waste Management", points: 2, matchKey: "CW|postwastemanagementplananddivert75fromlandfillofshingles" },
  { id: "ec_opt_cw_cw_1_3", pointNumber: "CW 1.3", text: "Central Cut Area", category: "Construction Waste Management", points: 2, matchKey: "CW|centralcutarea" },
  { id: "ec_opt_re_re_1_2_1", pointNumber: "RE 1.2 > 1", text: "Advanced Framing: > 2-stud corners where structurally feasible", category: "Resource Efficiency", points: 3, matchKey: "RE|advancedframing2studcornerswherestructurallyfeasible", goldMandatoryOverlap: true },
  { id: "ec_opt_re_re_1_2_2", pointNumber: "RE 1.2 > 2", text: "Advanced Framing: > Ladder T-walls where structurally feasible", category: "Resource Efficiency", points: 2, matchKey: "RE|advancedframingladdertwallswherestructurallyfeasible", goldMandatoryOverlap: true },
  { id: "ec_opt_re_re_1_2_3", pointNumber: "RE 1.2 > 3", text: "Advanced Framing: > Size headers for loads (non-structural headers in non-load bearing walls)", category: "Resource Efficiency", points: 1, matchKey: "RE|advancedframingsizeheadersforloadsnonstructuralheadersinnonloadbearingwalls", goldMandatoryOverlap: true },
  { id: "ec_opt_re_re_1_4", pointNumber: "RE 1.4", text: "Floor joists are 24\" on center  (≥80%)", category: "Resource Efficiency", points: 1, matchKey: "RE|floorjoistsare24oncenter80" },
  { id: "ec_opt_re_re_1_5", pointNumber: "RE 1.5", text: "Non-load bearing wall studs are 24\" on center", category: "Resource Efficiency", points: 1, matchKey: "RE|nonloadbearingwallstudsare24oncenter" },
  { id: "ec_opt_re_2", pointNumber: "2", text: "Precast insulated foundation walls (≥90%)", category: "Resource Efficiency", points: 2, matchKey: "RE|precastinsulatedfoundationwalls90" },
  { id: "ec_opt_re_2_1_1", pointNumber: "2.1 > 1", text: "Insulated concrete forms or precast autoclaved aerated concrete (≥90%): > Foundation walls", category: "Resource Efficiency", points: 2, matchKey: "RE|insulatedconcreteformsorprecastautoclavedaeratedconcrete90foundationwalls" },
  { id: "ec_opt_re_2_1_2", pointNumber: "2.1 > 2", text: "Insulated concrete forms or precast autoclaved aerated concrete (≥90%): > Exterior walls", category: "Resource Efficiency", points: 3, matchKey: "RE|insulatedconcreteformsorprecastautoclavedaeratedconcrete90exteriorwalls" },
  { id: "ec_opt_re_2_2", pointNumber: "2.2", text: "Engineered wall framing (≥90%)", category: "Resource Efficiency", points: 1, matchKey: "RE|engineeredwallframing90" },
  { id: "ec_opt_re_2_3_1", pointNumber: "2.3 > 1", text: "Deliver panelized construction or SIPs to the site pre-framed (≥90%): > Floors", category: "Resource Efficiency", points: 2, matchKey: "RE|deliverpanelizedconstructionorsipstothesitepreframed90floors" },
  { id: "ec_opt_re_2_3_2", pointNumber: "2.3 > 2", text: "Deliver panelized construction or SIPs to the site pre-framed (≥90%): > Exterior walls", category: "Resource Efficiency", points: 2, matchKey: "RE|deliverpanelizedconstructionorsipstothesitepreframed90exteriorwalls" },
  { id: "ec_opt_re_2_3_3", pointNumber: "2.3 > 3", text: "Deliver panelized construction or SIPs to the site pre-framed (≥90%): > Roof", category: "Resource Efficiency", points: 2, matchKey: "RE|deliverpanelizedconstructionorsipstothesitepreframed90roof" },
  { id: "ec_opt_re_2_3_4", pointNumber: "2.3 > 4", text: "Deliver panelized construction or SIPs to the site pre-framed (≥90%): > Modular construction", category: "Resource Efficiency", points: 2, matchKey: "RE|deliverpanelizedconstructionorsipstothesitepreframed90modularconstruction" },
  { id: "ec_opt_re_re_2_4", pointNumber: "RE 2.4", text: "Structural headers are steel or engineered wood (≥90%)", category: "Resource Efficiency", points: 2, matchKey: "RE|structuralheadersaresteelorengineeredwood90" },
  { id: "ec_opt_re_3_2", pointNumber: "3.2", text: "Lumber/Millwork/Flooring: Use No Tropical Wood", category: "Resource Efficiency", points: 2, matchKey: "RE|lumbermillworkflooringusenotropicalwood" },
  { id: "ec_opt_re_3_4_1", pointNumber: "RE 3.4 > 1", text: "Reused, recycled, MDF with no added urea-formaldehyde, local species or FSC certified wood in all: > Cabinet faces", category: "Resource Efficiency", points: 2, matchKey: "RE|reusedrecycledmdfwithnoaddedureaformaldehydelocalspeciesorfsccertifiedwoodinallcabinetfaces" },
  { id: "ec_opt_re_3_4_2", pointNumber: "RE 3.4 > 2", text: "Reused, recycled, MDF with no added urea-formaldehyde, local species or FSC certified wood in all: > Countertops", category: "Resource Efficiency", points: 2, matchKey: "RE|reusedrecycledmdfwithnoaddedureaformaldehydelocalspeciesorfsccertifiedwoodinallcountertops" },
  { id: "ec_opt_re_3_6", pointNumber: "RE 3.6", text: "Insulation (≥25% recycled content material)", category: "Resource Efficiency", points: 1, matchKey: "RE|insulation25recycledcontentmaterial" },
  { id: "ec_opt_re_3_8_1", pointNumber: "3.8 > 1", text: "Engineered trim: > Interior (≥80%)", category: "Resource Efficiency", points: 1, matchKey: "RE|engineeredtriminterior80" },
  { id: "ec_opt_re_3_8_2", pointNumber: "3.8 > 2", text: "Engineered trim: > Exterior, including soffit, fascia and trim (≥75%)", category: "Resource Efficiency", points: 1, matchKey: "RE|engineeredtrimexteriorincludingsoffitfasciaandtrim75" },
  { id: "ec_opt_re_re_4_0", pointNumber: "RE 4.0", text: "Gut Rehab (project exposing wall cavities or removing exterior cladding) or Adaptive Reuse (for adaptive reuse see addendum to worksheet)", category: "Resource Efficiency", points: 8, matchKey: "RE|gutrehabprojectexposingwallcavitiesorremovingexteriorcladdingoradaptivereuseforadaptivereuseseeaddendumtoworksheet" },
  { id: "ec_opt_du_du_1_11", pointNumber: "DU 1.11", text: "Enclosed crawlspace, if applicable to design", category: "Durability & Moisture Management", points: 2, matchKey: "DU|enclosedcrawlspaceifapplicabletodesign" },
  { id: "ec_opt_du_du_1_12", pointNumber: "DU 1.12", text: "Moisture-resistant wallboard in bathrooms", category: "Durability & Moisture Management", points: 2, matchKey: "DU|moistureresistantwallboardinbathrooms" },
  { id: "ec_opt_du_du_1_13", pointNumber: "DU 1.13", text: "Flashing at bottom of exterior walls integrated with drainage system", category: "Durability & Moisture Management", points: 2, matchKey: "DU|flashingatbottomofexteriorwallsintegratedwithdrainagesystem" },
  { id: "ec_opt_du_du_1_14", pointNumber: "DU 1.14", text: "Alternative termite treatment with no soil pretreatment", category: "Durability & Moisture Management", points: 2, matchKey: "DU|alternativetermitetreatmentwithnosoilpretreatment" },
  { id: "ec_opt_du_du_1_15_1", pointNumber: "DU 1.15 > 1", text: "Non-toxic pest treatment: > All lumber in contact with foundation (≥36\" above foundation)", category: "Durability & Moisture Management", points: 1, matchKey: "DU|nontoxicpesttreatmentalllumberincontactwithfoundation36abovefoundation" },
  { id: "ec_opt_du_du_1_15_2", pointNumber: "DU 1.15 > 2", text: "Non-toxic pest treatment: > All lumber", category: "Durability & Moisture Management", points: 2, matchKey: "DU|nontoxicpesttreatmentalllumber" },
  { id: "ec_opt_du_du_1_15_3", pointNumber: "DU 1.15 > 3", text: "Non-toxic pest treatment: > Mold inhibitor with warranty applied to all lumber", category: "Durability & Moisture Management", points: 1, matchKey: "DU|nontoxicpesttreatmentmoldinhibitorwithwarrantyappliedtoalllumber" },
  { id: "ec_opt_du_du_1_16", pointNumber: "DU 1.16", text: "Vented rain screen behind exterior cladding", category: "Durability & Moisture Management", points: 2, matchKey: "DU|ventedrainscreenbehindexteriorcladding" },
  { id: "ec_opt_du_du_1_17", pointNumber: "DU 1.17", text: "Install termite mesh system", category: "Durability & Moisture Management", points: 3, matchKey: "DU|installtermitemeshsystem" },
  { id: "ec_opt_du_du_1_20", pointNumber: "DU 1.20", text: "Insulate cold water pipes ≥R-2", category: "Durability & Moisture Management", points: 1, matchKey: "DU|insulatecoldwaterpipesr2" },
  { id: "ec_opt_du_du_1_21", pointNumber: "DU 1.21", text: "All entrance doors have overhang ≥3' depth", category: "Durability & Moisture Management", points: 1, matchKey: "DU|allentrancedoorshaveoverhang3depth" },
  { id: "ec_opt_du_du_2_9", pointNumber: "DU 2.9", text: "Additional dehumidification system: Basement or sealed crawlspace system", category: "Durability & Moisture Management", points: 2, matchKey: "DU|additionaldehumidificationsystembasementorsealedcrawlspacesystem", goldMandatoryOverlap: true },
  { id: "ec_opt_du_du_2_10", pointNumber: "DU 2.10", text: "Foundation drain at outside perimeter edge of footing surrounded with 6\" clean gravel and fabric filter", category: "Durability & Moisture Management", points: 2, matchKey: "DU|foundationdrainatoutsideperimeteredgeoffootingsurroundedwith6cleangravelandfabricfilter", goldMandatoryOverlap: true },
  { id: "ec_opt_du_du_2_11", pointNumber: "DU 2.11", text: "Install whole-house ENERGY STAR dehumidifier", category: "Durability & Moisture Management", points: 4, matchKey: "DU|installwholehouseenergystardehumidifier" },
  { id: "ec_opt_du_du_2_12", pointNumber: "DU 2.12", text: "Slab and crawlspace vapor barrier ≥10 mil or reinforced", category: "Durability & Moisture Management", points: 1, matchKey: "DU|slabandcrawlspacevaporbarrier10milorreinforced" },
  { id: "ec_opt_du_du_2_13", pointNumber: "DU 2.13", text: "Humidistat or thermidistat used with whole-house variable speed cooling system", category: "Durability & Moisture Management", points: 2, matchKey: "DU|humidistatorthermidistatusedwithwholehousevariablespeedcoolingsystem" },
  { id: "ec_opt_du_du_2_14_1", pointNumber: "DU 2.14 > 1", text: "Capillary break: > Between ground/footing or footing/foundation", category: "Durability & Moisture Management", points: 2, matchKey: "DU|capillarybreakbetweengroundfootingorfootingfoundation" },
  { id: "ec_opt_du_du_2_14_2", pointNumber: "DU 2.14 > 2", text: "Capillary break: > Between foundation and framing for all walls", category: "Durability & Moisture Management", points: 1, matchKey: "DU|capillarybreakbetweenfoundationandframingforallwalls" },
  { id: "ec_opt_iaq_iaq_1_5", pointNumber: "IAQ 1.5", text: "If installed, all fireplaces meet indoor air quality guidelines and have gasketed doors", category: "Indoor Air Quality", points: 2, matchKey: "IAQ|ifinstalledallfireplacesmeetindoorairqualityguidelinesandhavegasketeddoors", goldMandatoryOverlap: true },
  { id: "ec_opt_iaq_iaq_2_6", pointNumber: "IAQ 2.6", text: "Filters are ≥ MERV 8", category: "Indoor Air Quality", points: 1, matchKey: "IAQ|filtersaremerv8", goldMandatoryOverlap: true },
  { id: "ec_opt_iaq_iaq_2_7_1", pointNumber: "IAQ 2.7 > 1", text: "Certified low or no VOC materials: > Interior paints", category: "Indoor Air Quality", points: 1, matchKey: "IAQ|certifiedlowornovocmaterialsinteriorpaints" },
  { id: "ec_opt_iaq_iaq_2_7_2", pointNumber: "IAQ 2.7 > 2", text: "Certified low or no VOC materials: > Stains and finishes on wood floors", category: "Indoor Air Quality", points: 2, matchKey: "IAQ|certifiedlowornovocmaterialsstainsandfinishesonwoodfloors" },
  { id: "ec_opt_iaq_iaq_2_7_3", pointNumber: "IAQ 2.7 > 3", text: "Certified low or no VOC materials: > Sealants and adhesives", category: "Indoor Air Quality", points: 2, matchKey: "IAQ|certifiedlowornovocmaterialssealantsandadhesives" },
  { id: "ec_opt_iaq_iaq_2_7_4", pointNumber: "IAQ 2.7 > 4", text: "Certified low or no VOC materials: > Carpet", category: "Indoor Air Quality", points: 1, matchKey: "IAQ|certifiedlowornovocmaterialscarpet" },
  { id: "ec_opt_iaq_iaq_2_8", pointNumber: "IAQ 2.8", text: "Protect all bath fans until floor/wall finishing is complete", category: "Indoor Air Quality", points: 1, matchKey: "IAQ|protectallbathfansuntilfloorwallfinishingiscomplete" },
  { id: "ec_opt_iaq_iaq_2_9_1", pointNumber: "IAQ 2.9 > 1", text: "No added urea-formaldehyde: > Insulation", category: "Indoor Air Quality", points: 1, matchKey: "IAQ|noaddedureaformaldehydeinsulation" },
  { id: "ec_opt_iaq_iaq_2_9_2", pointNumber: "IAQ 2.9 > 2", text: "No added urea-formaldehyde: > Subfloor", category: "Indoor Air Quality", points: 1, matchKey: "IAQ|noaddedureaformaldehydesubfloor" },
  { id: "ec_opt_iaq_iaq_2_9_3", pointNumber: "IAQ 2.9 > 3", text: "No added urea-formaldehyde: > All cabinets, shelves, and countertops", category: "Indoor Air Quality", points: 2, matchKey: "IAQ|noaddedureaformaldehydeallcabinetsshelvesandcountertops" },
  { id: "ec_opt_iaq_iaq_2_10", pointNumber: "IAQ 2.10", text: "Seal all particle board surfaces with water-based sealant", category: "Indoor Air Quality", points: 1, matchKey: "IAQ|sealallparticleboardsurfaceswithwaterbasedsealant" },
  { id: "ec_opt_iaq_iaq_2_11", pointNumber: "IAQ 2.11", text: "No carpet in all units", category: "Indoor Air Quality", points: 3, matchKey: "IAQ|nocarpetinallunits" },
  { id: "ec_opt_iaq_iaq_2_12", pointNumber: "IAQ 2.12", text: "No carpet in main living area of all units", category: "Indoor Air Quality", points: 1, matchKey: "IAQ|nocarpetinmainlivingareaofallunits" },
  { id: "ec_opt_iaq_iaq_2_13", pointNumber: "IAQ 2.13", text: "Permanent walk-off mats installed at main entrances", category: "Indoor Air Quality", points: 1, matchKey: "IAQ|permanentwalkoffmatsinstalledatmainentrances" },
  { id: "ec_opt_be_1_13", pointNumber: "1.13", text: "Seal top plate to drywall at the attic level", category: "High Performance Building Envelope", points: 2, matchKey: "BE|sealtopplatetodrywallattheatticlevel" },
  { id: "ec_opt_be_be_1_14", pointNumber: "BE 1.14", text: "Comply with Air tight drywall approach (required if band area draft blocking is not used)", category: "High Performance Building Envelope", points: 4, matchKey: "BE|complywithairtightdrywallapproachrequiredifbandareadraftblockingisnotused" },
  { id: "ec_opt_be_be_1_15", pointNumber: "BE 1.15", text: "Gypcrete on all framed floors separating unit envelopes", category: "High Performance Building Envelope", points: 1, matchKey: "BE|gypcreteonallframedfloorsseparatingunitenvelopes" },
  { id: "ec_opt_be_be_1_16", pointNumber: "BE 1.16", text: "Two pour application of gypcrete to include areas blocked by drywall", category: "High Performance Building Envelope", points: 1, matchKey: "BE|twopourapplicationofgypcretetoincludeareasblockedbydrywall" },
  { id: "ec_opt_be_be_1_17", pointNumber: "BE 1.17", text: "Firewalls/party walls that eliminate air gap (UL-U370 or equivalent)", category: "High Performance Building Envelope", points: 2, matchKey: "BE|firewallspartywallsthateliminateairgapulu370orequivalent" },
  { id: "ec_opt_be_be_3_10_a", pointNumber: "BE 3.10 > A.", text: "Insulation installation quality (floors, walls and ceilings): > Grade I", category: "High Performance Building Envelope", points: 3, matchKey: "BE|insulationinstallationqualityfloorswallsandceilingsgradei", goldMandatoryOverlap: true },
  { id: "ec_opt_be_be_3_10_b", pointNumber: "BE 3.10 > B.", text: "Insulation installation quality (floors, walls and ceilings): > Grade II with insulated sheathing ≥ R-3 (100%)", category: "High Performance Building Envelope", points: 2, matchKey: "BE|insulationinstallationqualityfloorswallsandceilingsgradeiiwithinsulatedsheathingr3100", goldMandatoryOverlap: true },
  { id: "ec_opt_be_be_3_11", pointNumber: "BE 3.11", text: "Corners ≥ R-6", category: "High Performance Building Envelope", points: 1, matchKey: "BE|cornersr6", goldMandatoryOverlap: true },
  { id: "ec_opt_be_be_3_12", pointNumber: "BE 3.12", text: "Headers ≥ R-3", category: "High Performance Building Envelope", points: 1, matchKey: "BE|headersr3", goldMandatoryOverlap: true },
  { id: "ec_opt_be_be_3_13", pointNumber: "BE 3.13", text: "Fiberglass batts are unfaced/friction fit", category: "High Performance Building Envelope", points: 1, matchKey: "BE|fiberglassbattsareunfacedfrictionfit", goldMandatoryOverlap: true },
  { id: "ec_opt_be_be_3_14_1", pointNumber: "BE 3.14 > 1", text: "Insulate with foam: > Exterior walls including band area", category: "High Performance Building Envelope", points: 4, matchKey: "BE|insulatewithfoamexteriorwallsincludingbandarea" },
  { id: "ec_opt_be_be_3_14_2", pointNumber: "BE 3.14 > 2", text: "Insulate with foam: > Floor system over crawlspace, basement, or parking garage", category: "High Performance Building Envelope", points: 2, matchKey: "BE|insulatewithfoamfloorsystemovercrawlspacebasementorparkinggarage" },
  { id: "ec_opt_be_be_3_15_1", pointNumber: "BE 3.15 > 1", text: "Walls: > Seal and insulate crawlspace walls ≥ R-10 continuous", category: "High Performance Building Envelope", points: 2, matchKey: "BE|wallssealandinsulatecrawlspacewallsr10continuous" },
  { id: "ec_opt_be_be_3_15_2", pointNumber: "BE 3.15 > 2", text: "Walls: > Insulate unfinished basement walls instead of ceiling", category: "High Performance Building Envelope", points: 1, matchKey: "BE|wallsinsulateunfinishedbasementwallsinsteadofceiling" },
  { id: "ec_opt_be_be_3_15_3", pointNumber: "BE 3.15 > 3", text: "Walls: > Insulate basement walls with continuous insulation", category: "High Performance Building Envelope", points: 2, matchKey: "BE|wallsinsulatebasementwallswithcontinuousinsulation" },
  { id: "ec_opt_be_be_3_15_4", pointNumber: "BE 3.15 > 4", text: "Walls: > Insulate exterior walls and band joist ≥ R-19", category: "High Performance Building Envelope", points: 2, matchKey: "BE|wallsinsulateexteriorwallsandbandjoistr19" },
  { id: "ec_opt_be_be_3_15_5", pointNumber: "BE 3.15 > 5", text: "Walls: > Insulate with spray foam insulation: Flash and batt insulation including band area", category: "High Performance Building Envelope", points: 2, matchKey: "BE|wallsinsulatewithsprayfoaminsulationflashandbattinsulationincludingbandarea" },
  { id: "ec_opt_be_be_3_15_6", pointNumber: "BE 3.15 > 6", text: "Walls: > Insulate exterior walls and band joist ≥ R-20 or ≥ R-13 cavity plus R-5 insulated sheathing", category: "High Performance Building Envelope", points: 3, matchKey: "BE|wallsinsulateexteriorwallsandbandjoistr20orr13cavityplusr5insulatedsheathing" },
  { id: "ec_opt_be_be_3_16_1", pointNumber: "BE 3.16 > 1", text: "Continuous exterior insulation: > ≥R-3", category: "High Performance Building Envelope", points: 5, matchKey: "BE|continuousexteriorinsulationr3" },
  { id: "ec_opt_be_be_3_16_2", pointNumber: "BE 3.16 > 2", text: "Continuous exterior insulation: > ≥R-5", category: "High Performance Building Envelope", points: 7, matchKey: "BE|continuousexteriorinsulationr5" },
  { id: "ec_opt_be_be_3_17_1", pointNumber: "BE 3.17 > 1", text: "Ceilings: > Flat Vented: Climate Zone 4 ≥ R-60", category: "High Performance Building Envelope", points: 2, matchKey: "BE|ceilingsflatventedclimatezone4r60" },
  { id: "ec_opt_be_be_3_17_2", pointNumber: "BE 3.17 > 2", text: "Ceilings: > Continuous Roof Deck: Climate Zone 4 ≥ R-35", category: "High Performance Building Envelope", points: 2, matchKey: "BE|ceilingscontinuousroofdeckclimatezone4r35" },
  { id: "ec_opt_be_be_3_17_3", pointNumber: "BE 3.17 > 3", text: "Ceilings: > Sloped: Climate Zone 4  ≥ R-49", category: "High Performance Building Envelope", points: 2, matchKey: "BE|ceilingsslopedclimatezone4r49" },
  { id: "ec_opt_be_be_3_18", pointNumber: "BE 3.18", text: "Attic kneewall insulated ≥ R-22", category: "High Performance Building Envelope", points: 2, matchKey: "BE|attickneewallinsulatedr22" },
  { id: "ec_opt_be_4_4_1", pointNumber: "4.4 > 1", text: "Door U-factor: > Opaque door:  U factor≤ 0.21", category: "High Performance Building Envelope", points: 2, matchKey: "BE|doorufactoropaquedoorufactor021" },
  { id: "ec_opt_be_4_4_2", pointNumber: "4.4 > 2", text: "Door U-factor: > Door with ≤ 50% glass:  U-factor ≤ 0.27", category: "High Performance Building Envelope", points: 1, matchKey: "BE|doorufactordoorwith50glassufactor027" },
  { id: "ec_opt_be_4_4_3", pointNumber: "4.4 > 3", text: "Door U-factor: > Door with > 50% glass:  U-factor ≤ 0.32", category: "High Performance Building Envelope", points: 1, matchKey: "BE|doorufactordoorwith50glassufactor032" },
  { id: "ec_opt_be_4_5_1", pointNumber: "4.5 > 1", text: "Window U-factor and SHGC: > U-factor ≤0.32", category: "High Performance Building Envelope", points: 1, matchKey: "BE|windowufactorandshgcufactor032" },
  { id: "ec_opt_be_4_5_2", pointNumber: "4.5 > 2", text: "Window U-factor and SHGC: > SHGC ≤0.27", category: "High Performance Building Envelope", points: 2, matchKey: "BE|windowufactorandshgcshgc027" },
  { id: "ec_opt_be_4_6_1", pointNumber: "4.6 > 1", text: "Skylight U-factor and SHGC: > U-factor ≤0.50", category: "High Performance Building Envelope", points: 1, matchKey: "BE|skylightufactorandshgcufactor050" },
  { id: "ec_opt_be_4_6_2", pointNumber: "4.6 > 2", text: "Skylight U-factor and SHGC: > SHGC ≤0.27", category: "High Performance Building Envelope", points: 2, matchKey: "BE|skylightufactorandshgcshgc027" },
  { id: "ec_opt_be_4_7_1", pointNumber: "4.7 > 1", text: "Window U-factor and SHGC: > U-factor ≤0.25", category: "High Performance Building Envelope", points: 2, matchKey: "BE|windowufactorandshgcufactor025" },
  { id: "ec_opt_be_4_7_2", pointNumber: "4.7 > 2", text: "Window U-factor and SHGC: > SHGC ≤0.24", category: "High Performance Building Envelope", points: 3, matchKey: "BE|windowufactorandshgcshgc024" },
  { id: "ec_opt_be_4_8_1", pointNumber: "4.8 > 1", text: "Skylight U-factor and SHGC: > U-factor ≤0.43", category: "High Performance Building Envelope", points: 2, matchKey: "BE|skylightufactorandshgcufactor043" },
  { id: "ec_opt_be_4_8_2", pointNumber: "4.8 > 2", text: "Skylight U-factor and SHGC: > SHGC ≤0.24", category: "High Performance Building Envelope", points: 3, matchKey: "BE|skylightufactorandshgcshgc024" },
  { id: "ec_opt_be_4_11", pointNumber: "4.11", text: "Solar shade screens (min all east and west windows)", category: "High Performance Building Envelope", points: 2, matchKey: "BE|solarshadescreensminalleastandwestwindows" },
  { id: "ec_opt_be_be_5_0_a", pointNumber: "BE 5.0 > A.", text: "If Ducts located in unconditioned attic: > Attic Side Radiant Barrier", category: "High Performance Building Envelope", points: 2, matchKey: "BE|ifductslocatedinunconditionedatticatticsideradiantbarrier" },
  { id: "ec_opt_es_es_1_8_a", pointNumber: "ES 1.8 > A.", text: "Heating equipment efficiency: > ENERGY STAR qualified furnace(s) ≥95 AFUE and within 40% of load calculation", category: "Energy Efficient Systems", points: 2, matchKey: "ES|heatingequipmentefficiencyenergystarqualifiedfurnaces95afueandwithin40ofloadcalculation", goldMandatoryOverlap: true },
  { id: "ec_opt_es_es_1_8_b", pointNumber: "ES 1.8 > B.", text: "Heating equipment efficiency: > ENERGY STAR qualified heat pump(s) ≥8.5 HSPF and within 25% of load calculation", category: "Energy Efficient Systems", points: 2, matchKey: "ES|heatingequipmentefficiencyenergystarqualifiedheatpumps85hspfandwithin25ofloadcalculation", goldMandatoryOverlap: true },
  { id: "ec_opt_es_es_1_10", pointNumber: "ES 1.10", text: "ENERGY STAR qualified cooling equipment ≥SEER 15", category: "Energy Efficient Systems", points: 2, matchKey: "ES|energystarqualifiedcoolingequipmentseer15", goldMandatoryOverlap: true },
  { id: "ec_opt_es_es_1_15", pointNumber: "ES 1.15", text: "ENERGY STAR qualified cooling equipment ≥ SEER 16", category: "Energy Efficient Systems", points: 3, matchKey: "ES|energystarqualifiedcoolingequipmentseer16" },
  { id: "ec_opt_es_es_1_16", pointNumber: "ES 1.16", text: "Heat pump efficiency ≥9.0 HSPF", category: "Energy Efficient Systems", points: 2, matchKey: "ES|heatpumpefficiency90hspf" },
  { id: "ec_opt_es_es_1_18", pointNumber: "ES 1.18", text: "Condenser units are spaced 2 feet apart", category: "Energy Efficient Systems", points: 2, matchKey: "ES|condenserunitsarespaced2feetapart" },
  { id: "ec_opt_es_2_11_1", pointNumber: "2.11 > 1", text: "Minimize pressure imbalance within units: > Install fully ducted jumper ducts, transfer grills, or dedicated return for each bedroom", category: "Energy Efficient Systems", points: 2, matchKey: "ES|minimizepressureimbalancewithinunitsinstallfullyductedjumperductstransfergrillsordedicatedreturnforeachbedroom" },
  { id: "ec_opt_es_2_11_2", pointNumber: "2.11 > 2", text: "Minimize pressure imbalance within units: > Measured pressure differential ≤ 3 Pa between bedroom and return", category: "Energy Efficient Systems", points: 3, matchKey: "ES|minimizepressureimbalancewithinunitsmeasuredpressuredifferential3pabetweenbedroomandreturn" },
  { id: "ec_opt_es_2_12", pointNumber: "2.12", text: "Install rigid duct work or pull all flex ducts with no pinches and support at intervals ≤ 5'", category: "Energy Efficient Systems", points: 2, matchKey: "ES|installrigidductworkorpullallflexductswithnopinchesandsupportatintervals5" },
  { id: "ec_opt_es_2_15", pointNumber: "2.15", text: "HVAC system and ductwork is dry and clean", category: "Energy Efficient Systems", points: 1, matchKey: "ES|hvacsystemandductworkisdryandclean" },
  { id: "ec_opt_es_es_2_16", pointNumber: "ES 2.16", text: "Locate entire duct system within conditioned space", category: "Energy Efficient Systems", points: 5, matchKey: "ES|locateentireductsystemwithinconditionedspace" },
  { id: "ec_opt_es_es_2_17_1", pointNumber: "ES 2.17 > 1", text: "Duct design and installation: > Rigid metal supply trunk", category: "Energy Efficient Systems", points: 2, matchKey: "ES|ductdesignandinstallationrigidmetalsupplytrunk" },
  { id: "ec_opt_es_es_2_17_2", pointNumber: "ES 2.17 > 2", text: "Duct design and installation: > Space all supply duct take-offs ≥6\" apart", category: "Energy Efficient Systems", points: 1, matchKey: "ES|ductdesignandinstallationspaceallsupplyducttakeoffs6apart" },
  { id: "ec_opt_es_es_2_17_3", pointNumber: "ES 2.17 > 3", text: "Duct design and installation: > Install rigid circular duct as supply plenum", category: "Energy Efficient Systems", points: 2, matchKey: "ES|ductdesignandinstallationinstallrigidcircularductassupplyplenum" },
  { id: "ec_opt_es_es_2_18", pointNumber: "ES 2.18", text: "Duct insulation in unconditioned spaces ≥R-10", category: "Energy Efficient Systems", points: 1, matchKey: "ES|ductinsulationinunconditionedspacesr10" },
  { id: "ec_opt_es_es_2_19", pointNumber: "ES 2.19", text: "Return plenum duct take-off free area is 120% of supply plenum duct take-off free area", category: "Energy Efficient Systems", points: 2, matchKey: "ES|returnplenumducttakeofffreeareais120ofsupplyplenumducttakeofffreearea" },
  { id: "ec_opt_es_4_8", pointNumber: "4.8", text: "If installed, ceiling fans must be ENERGY STAR qualified (1/bedroom and 1 in living room)", category: "Energy Efficient Systems", points: 1, matchKey: "ES|ifinstalledceilingfansmustbeenergystarqualified1bedroomand1inlivingroom" },
  { id: "ec_opt_es_4_9", pointNumber: "4.9", text: "ENERGY STAR bath fans with properly sized ductwork and measured airflow ≥50 cfm", category: "Energy Efficient Systems", points: 2, matchKey: "ES|energystarbathfanswithproperlysizedductworkandmeasuredairflow50cfm" },
  { id: "ec_opt_es_4_1", pointNumber: "4.10", text: "Electric kitchen range vented to exterior ≥ 100 cfm fan", category: "Energy Efficient Systems", points: 3, matchKey: "ES|electrickitchenrangeventedtoexterior100cfmfan" },
  { id: "ec_opt_es_4_12", pointNumber: "4.12", text: "Install and label accessible ventilation controls, with override controls for continuously operating ventilation fans", category: "Energy Efficient Systems", points: 1, matchKey: "ES|installandlabelaccessibleventilationcontrolswithoverridecontrolsforcontinuouslyoperatingventilationfans" },
  { id: "ec_opt_es_4_13", pointNumber: "4.13", text: "Supply/exhaust fans rated at ≤3 sones (intermittent) and ≤1 sone (continuous)", category: "Energy Efficient Systems", points: 1, matchKey: "ES|supplyexhaustfansratedat3sonesintermittentand1sonecontinuous" },
  { id: "ec_opt_es_4_14_1", pointNumber: "4.14 > 1", text: "Radon resistant construction: > Passive, radon/soil gas vent system labeled on each floor", category: "Energy Efficient Systems", points: 1, matchKey: "ES|radonresistantconstructionpassiveradonsoilgasventsystemlabeledoneachfloor" },
  { id: "ec_opt_es_4_14_2", pointNumber: "4.14 > 2", text: "Radon resistant construction: > Radon test of building prior to occupancy", category: "Energy Efficient Systems", points: 1, matchKey: "ES|radonresistantconstructionradontestofbuildingpriortooccupancy" },
  { id: "ec_opt_es_4_15", pointNumber: "4.15", text: "Exhaust fan wired with light in bathroom", category: "Energy Efficient Systems", points: 1, matchKey: "ES|exhaustfanwiredwithlightinbathroom" },
  { id: "ec_opt_es_4_16", pointNumber: "4.16", text: "Duct all exhaust fans with rigid duct", category: "Energy Efficient Systems", points: 1, matchKey: "ES|ductallexhaustfanswithrigidduct" },
  { id: "ec_opt_es_4_17", pointNumber: "4.17", text: "Automatic (timer and/or humidistat) bathroom exhaust fan controls", category: "Energy Efficient Systems", points: 2, matchKey: "ES|automatictimerandorhumidistatbathroomexhaustfancontrols" },
  { id: "ec_opt_es_4_18", pointNumber: "4.18", text: "Energy recovery ventilator", category: "Energy Efficient Systems", points: 3, matchKey: "ES|energyrecoveryventilator" },
  { id: "ec_opt_es_4_19", pointNumber: "4.19", text: "Vent storage room to outside", category: "Energy Efficient Systems", points: 1, matchKey: "ES|ventstorageroomtooutside" },
  { id: "ec_opt_es_es_5_6_a", pointNumber: "ES 5.6  > A.", text: "Type of water heater: > Solar domestic (≥40% annual load based on unit demand)", category: "Energy Efficient Systems", points: 6, matchKey: "ES|typeofwaterheatersolardomestic40annualloadbasedonunitdemand" },
  { id: "ec_opt_es_es_5_6_b", pointNumber: "ES 5.6  > B.", text: "Type of water heater: > High efficiency tankless water heater (≥ .92 EF) with insulated buffer tank", category: "Energy Efficient Systems", points: 4, matchKey: "ES|typeofwaterheaterhighefficiencytanklesswaterheater92efwithinsulatedbuffertank" },
  { id: "ec_opt_es_es_5_7", pointNumber: "ES 5.7", text: "Hot water piping insulation ≥R-4 (100%)", category: "Energy Efficient Systems", points: 2, matchKey: "ES|hotwaterpipinginsulationr4100" },
  { id: "ec_opt_es_6_3", pointNumber: "6.3", text: "If installed, ENERGY STAR qualified clothes washer", category: "Energy Efficient Systems", points: 2, matchKey: "ES|ifinstalledenergystarqualifiedclotheswasher" },
  { id: "ec_opt_es_6_4", pointNumber: "6.4", text: "If installed, high efficiency clothes dryer with moisture sensor (not applicable to commercial dryers)", category: "Energy Efficient Systems", points: 2, matchKey: "ES|ifinstalledhighefficiencyclothesdryerwithmoisturesensornotapplicabletocommercialdryers" },
  { id: "ec_opt_es_es_6_5_a", pointNumber: "ES 6.5 > A.", text: "Fixtures and bulbs: > ENERGY STAR qualified compact fluorescent fixtures or LED bulbs (100%)", category: "Energy Efficient Systems", points: 2, matchKey: "ES|fixturesandbulbsenergystarqualifiedcompactfluorescentfixturesorledbulbs100" },
  { id: "ec_opt_es_es_6_5_b", pointNumber: "ES 6.5 > B.", text: "Fixtures and bulbs: > Ballasted compact fluorescents or LED bulbs at all recessed light fixtures", category: "Energy Efficient Systems", points: 1, matchKey: "ES|fixturesandbulbsballastedcompactfluorescentsorledbulbsatallrecessedlightfixtures" },
  { id: "ec_opt_es_7", pointNumber: "7", text: "100% LED bulbs in all corridor/breezeway and all common spaces", category: "Energy Efficient Systems", points: 2, matchKey: "ES|100ledbulbsinallcorridorbreezewayandallcommonspaces" },
  { id: "ec_opt_es_es_7_1_1", pointNumber: "ES 7.1 > 1", text: "Control systems: > Automatic indoor lighting controls", category: "Energy Efficient Systems", points: 2, matchKey: "ES|controlsystemsautomaticindoorlightingcontrols" },
  { id: "ec_opt_es_es_7_1_2", pointNumber: "ES 7.1 > 2", text: "Control systems: > Automatic outdoor lighting controls", category: "Energy Efficient Systems", points: 2, matchKey: "ES|controlsystemsautomaticoutdoorlightingcontrols" },
  { id: "ec_opt_es_es_7_2_3", pointNumber: "ES 7.2 > 3", text: "High Efficiency Exterior Lighting: > High efficiency exterior lighting using 100% LED bulbs", category: "Energy Efficient Systems", points: 2, matchKey: "ES|highefficiencyexteriorlightinghighefficiencyexteriorlightingusing100ledbulbs" },
  { id: "ec_opt_es_es_7_3", pointNumber: "ES 7.3", text: "High efficiency elevators", category: "Energy Efficient Systems", points: 2, matchKey: "ES|highefficiencyelevators" },
  { id: "ec_opt_we_we_1_3", pointNumber: "WE 1.3", text: "If installed, water treatment system NSF certified, ≥85% efficient", category: "Water Efficiency", points: 2, matchKey: "WE|ifinstalledwatertreatmentsystemnsfcertified85efficient" },
  { id: "ec_opt_we_we_1_4", pointNumber: "WE 1.4", text: "If installed, water softeners certified to NSF/ANSI Standard 44", category: "Water Efficiency", points: 2, matchKey: "WE|ifinstalledwatersoftenerscertifiedtonsfansistandard44" },
  { id: "ec_opt_we_we_1_6", pointNumber: "WE 1.6", text: "WaterSense labeled Showerhead (1.75 gpm)", category: "Water Efficiency", points: 1, matchKey: "WE|watersenselabeledshowerhead175gpm" },
  { id: "ec_opt_we_we_1_7", pointNumber: "WE 1.7", text: "Toilet (≤1.1 avg. gal/flush)", category: "Water Efficiency", points: 2, matchKey: "WE|toilet11avggalflush" },
  { id: "ec_opt_we_we_1_8", pointNumber: "WE 1.8", text: "Waterless urinals in common areas", category: "Water Efficiency", points: 2, matchKey: "WE|waterlessurinalsincommonareas" },
  { id: "ec_opt_we_we_1_9", pointNumber: "WE 1.9", text: "Greywater system for toilet flushing", category: "Water Efficiency", points: 4, matchKey: "WE|greywatersystemfortoiletflushing" },
  { id: "ec_opt_we_we_1_10", pointNumber: "WE 1.10", text: "Rainwater harvest system for indoor water use", category: "Water Efficiency", points: 4, matchKey: "WE|rainwaterharvestsystemforindoorwateruse" },
  { id: "ec_opt_we_we_1_12", pointNumber: "WE 1.12", text: "Hot water demand ≤0.13 gal of water between loop and fixture and ≤2 gal of water in loop between water heater and furthest fixture (not applicable to central systems)", category: "Water Efficiency", points: 2, matchKey: "WE|hotwaterdemand013galofwaterbetweenloopandfixtureand2galofwaterinloopbetweenwaterheaterandfurthestfixturenotapplicabletocentralsystems" },
  { id: "ec_opt_we_we_2_4_a", pointNumber: "WE 2.4 > A.", text: "Landscape design: > Turf ≤ 40% of landscaped area", category: "Water Efficiency", points: 2, matchKey: "WE|landscapedesignturf40oflandscapedarea" },
  { id: "ec_opt_we_we_2_5", pointNumber: "WE 2.5", text: "Vegetate slopes exceeding 4:1", category: "Water Efficiency", points: 1, matchKey: "WE|vegetateslopesexceeding41" },
  { id: "ec_opt_we_we_2_7", pointNumber: "WE 2.7", text: "Drought-tolerant/native landscaping turf and plants", category: "Water Efficiency", points: 1, matchKey: "WE|droughttolerantnativelandscapingturfandplants" },
  { id: "ec_opt_we_we_2_6_1", pointNumber: "WE 2.6 > 1", text: "If installed, irrigation system is: (Max 4 points) > Design, install, and audit irrigation system by WaterSense Irrigation Partner with no leaks", category: "Water Efficiency", points: 2, matchKey: "WE|ifinstalledirrigationsystemismax4pointsdesigninstallandauditirrigationsystembywatersenseirrigationpartnerwithnoleaks" },
  { id: "ec_opt_we_we_2_6_2", pointNumber: "WE 2.6 > 2", text: "If installed, irrigation system is: (Max 4 points) > Micro-irrigation system (e.g., drip irrigation) includes pressure regulator, filter and flush end assemblies", category: "Water Efficiency", points: 2, matchKey: "WE|ifinstalledirrigationsystemismax4pointsmicroirrigationsystemegdripirrigationincludespressureregulatorfilterandflushendassemblies" },
  { id: "ec_opt_we_we_2_6_3", pointNumber: "WE 2.6 > 3", text: "If installed, irrigation system is: (Max 4 points) > Distribution uniformity ≥65% lower quarter", category: "Water Efficiency", points: 2, matchKey: "WE|ifinstalledirrigationsystemismax4pointsdistributionuniformity65lowerquarter" },
  { id: "ec_opt_we_we_2_6_4", pointNumber: "WE 2.6 > 4", text: "If installed, irrigation system is: (Max 4 points) > Install sprinklers only on turfgrass, pop-up height ≥4\"", category: "Water Efficiency", points: 1, matchKey: "WE|ifinstalledirrigationsystemismax4pointsinstallsprinklersonlyonturfgrasspopupheight4" },
  { id: "ec_opt_we_we_2_6_5", pointNumber: "WE 2.6 > 5", text: "If installed, irrigation system is: (Max 4 points) > Establish grow-in phase and post landscape seasonal water schedules at irrigation controller", category: "Water Efficiency", points: 2, matchKey: "WE|ifinstalledirrigationsystemismax4pointsestablishgrowinphaseandpostlandscapeseasonalwaterschedulesatirrigationcontroller" },
  { id: "ec_opt_we_we_2_9_1", pointNumber: "WE 2.9 > 1", text: "Irrigation: (Max 5 points) > Greywater irrigation system", category: "Water Efficiency", points: 3, matchKey: "WE|irrigationmax5pointsgreywaterirrigationsystem" },
  { id: "ec_opt_we_we_2_9_2", pointNumber: "WE 2.9 > 2", text: "Irrigation: (Max 5 points) > Rainwater irrigation system", category: "Water Efficiency", points: 3, matchKey: "WE|irrigationmax5pointsrainwaterirrigationsystem" },
  { id: "ec_opt_we_we_2_9_3", pointNumber: "WE 2.9 > 3", text: "Irrigation: (Max 5 points) > Zone irrigation system for specific water needs in each planting area", category: "Water Efficiency", points: 2, matchKey: "WE|irrigationmax5pointszoneirrigationsystemforspecificwaterneedsineachplantingarea" },
  { id: "ec_opt_we_we_2_9_4", pointNumber: "WE 2.9 > 4", text: "Irrigation: (Max 5 points) > Provide weather station or soil moisture sensor on irrigation system", category: "Water Efficiency", points: 2, matchKey: "WE|irrigationmax5pointsprovideweatherstationorsoilmoisturesensoronirrigationsystem" },
  { id: "ec_opt_we_we_2_10", pointNumber: "WE 2.10", text: "Timer on exterior water spigots", category: "Water Efficiency", points: 1, matchKey: "WE|timeronexteriorwaterspigots" },
  { id: "ec_opt_eo_1_2", pointNumber: "1.2", text: "Community Recycling Facility", category: "Education & Operations", points: 2, matchKey: "EO|communityrecyclingfacility" },
  { id: "ec_opt_eo_eo_2_1", pointNumber: "EO 2.1", text: "Property Maintenance Staff representative attends design review and/or kick off meeting", category: "Education & Operations", points: 1, matchKey: "EO|propertymaintenancestaffrepresentativeattendsdesignreviewandorkickoffmeeting" },
  { id: "ec_opt_eo_eo_2_2", pointNumber: "EO 2.2", text: "Market EarthCraft Multifamily program", category: "Education & Operations", points: 1, matchKey: "EO|marketearthcraftmultifamilyprogram" },
  { id: "ec_opt_eo_eo_2_3", pointNumber: "EO 2.3", text: "Provide pre-occupancy briefing for tenant", category: "Education & Operations", points: 2, matchKey: "EO|providepreoccupancybriefingfortenant" },
  { id: "ec_opt_eo_eo_2_4", pointNumber: "EO 2.4", text: "Project participates in post occupancy project debriefing", category: "Education & Operations", points: 2, matchKey: "EO|projectparticipatesinpostoccupancyprojectdebriefing" },
  { id: "ec_opt_eo_eo_2_5", pointNumber: "EO 2.5", text: "Environmental management and building maintenance guidelines for staff", category: "Education & Operations", points: 2, matchKey: "EO|environmentalmanagementandbuildingmaintenanceguidelinesforstaff" },
  { id: "ec_opt_eo_eo_2_6", pointNumber: "EO 2.6", text: "Landscape maintenance guide for maintenance and management personnel", category: "Education & Operations", points: 2, matchKey: "EO|landscapemaintenanceguideformaintenanceandmanagementpersonnel" },
  { id: "ec_opt_eo_eo_3_0", pointNumber: "EO 3.0", text: "ENERGY STAR Multifamily New Construction", category: "Education & Operations", points: 2, matchKey: "EO|energystarmultifamilynewconstruction" },
  { id: "ec_opt_inv_1", pointNumber: "1", text: "On-site fuel cell or co-generation system", category: "Innovation", points: 4, matchKey: "INV|onsitefuelcellorcogenerationsystem" },
  { id: "ec_opt_inv_1_1", pointNumber: "1.1", text: "Solar-ready design", category: "Innovation", points: 2, matchKey: "INV|solarreadydesign" },
  { id: "ec_opt_inv_1_2", pointNumber: "1.2", text: "Wind and/or Solar electric system (10% of project requirements)", category: "Innovation", points: 5, matchKey: "INV|windandorsolarelectricsystem10ofprojectrequirements" },
  { id: "ec_opt_inv_in_1_3", pointNumber: "IN 1.3", text: "100% of stormwater kept on site and used for development operations", category: "Innovation", points: 4, matchKey: "INV|100ofstormwaterkeptonsiteandusedfordevelopmentoperations" },
  { id: "ec_opt_inv_in_1_4", pointNumber: "IN 1.4", text: "Common areas use solar and/or wind electric system (80% of demand)", category: "Innovation", points: 4, matchKey: "INV|commonareasusesolarandorwindelectricsystem80ofdemand" },
  { id: "ec_opt_du_du_2_16", pointNumber: "DU 2.16", text: "1.5' roof overhangs on all building elevations", category: "Durability & Moisture Management", points: 1, matchKey: "DU|15roofoverhangsonallbuildingelevations" },
  { id: "ec_opt_iaq_iaq_2_14", pointNumber: "IAQ 2.14", text: "All mechanically provided ventilation air crosses a MERV 13 or better filter prior to distribution", category: "Indoor Air Quality", points: 2, matchKey: "IAQ|allmechanicallyprovidedventilationaircrossesamerv13orbetterfilterpriortodistribution" },
  { id: "ec_opt_es_es_1_21", pointNumber: "ES 1.21", text: "Set all air handlers to design air flow", category: "Energy Efficient Systems", points: 1, matchKey: "ES|setallairhandlerstodesignairflow" },
];

function ecOptNormText(t) {
  return (t || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function ecCellStr(v) { return v === null || v === undefined ? "" : String(v).trim(); }

function ecIsRealPoints(v) {
  const s = ecCellStr(v);
  if (s === "" || s === "-") return false;
  return /^\d+(\.\d+)?$/.test(s);
}

// The workbook's "Planned" column defaults some unselected rows to 0 (formula result)
// rather than leaving them blank — only a truthy non-zero value means the project
// actually intends to pursue the item.
function ecIsPlanned(v) {
  const s = ecCellStr(v);
  if (s === "" || s === "-") return false;
  const n = Number(s);
  return isNaN(n) ? true : n !== 0;
}

const EC_CAT_HEADER_RE = /^[A-Z][A-Z &]+\([A-Z]{2,5}\)$/;

// Parses a populated EarthCraft Multifamily workbook (.xlsx) and returns the optional,
// point-bearing items this specific project has marked "Planned" that also appear in
// EARTHCRAFT_OPTIONAL_LIBRARY (curated by TA review — see note above). Row layout and
// hierarchy (category header -> subsection -> level note -> up to 3 levels of items,
// Points/Planned in columns E/F) verified directly against real V6.5 and V7 workbooks.
function parseEarthCraftWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = wb.SheetNames.includes("Worksheet") ? "Worksheet" : wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  if (!sheet) throw new Error("Couldn't find a 'Worksheet' sheet in this file.");
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });

  const libByKey = new Map(EARTHCRAFT_OPTIONAL_LIBRARY.map(i => [i.matchKey, i]));

  let catCode = null, optionalSection = false;
  let topCode = null, topText = null, midMarker = null, midText = null;

  const matched = [], seenIds = new Set(), unmatched = [];

  const handleLeaf = (text, pointsVal, plannedVal, statusVal, code) => {
    if (!optionalSection || !ecIsRealPoints(pointsVal) || !ecIsPlanned(plannedVal)) return;
    const key = `${catCode}|${ecOptNormText(text)}`;
    const libItem = libByKey.get(key);
    if (libItem) {
      if (!seenIds.has(libItem.id)) { seenIds.add(libItem.id); matched.push({ ...libItem, workbookCode: code, workbookPoints: pointsVal, workbookPlanned: plannedVal, workbookStatus: ecCellStr(statusVal) }); }
    } else {
      unmatched.push({ category: catCode, code, text, points: pointsVal, planned: plannedVal });
    }
  };

  for (const row of rows) {
    const a = ecCellStr(row[0]), b = ecCellStr(row[1]), c = ecCellStr(row[2]), d = ecCellStr(row[3]);
    const e = row[4], f = row[5], g = row[6];
    let restEmpty = true;
    for (let i = 1; i <= 6; i++) { if (ecCellStr(row[i]) !== "") { restEmpty = false; break; } }
    const upperA = a.toUpperCase();

    if (a && restEmpty && (upperA.includes("REQUIRED AT") || upperA.includes("OPTIONAL AT"))) {
      optionalSection = upperA.includes("OPTIONAL");
      topCode = topText = midMarker = midText = null;
      continue;
    }
    if (a && restEmpty && EC_CAT_HEADER_RE.test(a)) {
      catCode = a.slice(a.lastIndexOf("(") + 1, a.lastIndexOf(")")).trim();
      optionalSection = false;
      continue;
    }
    if (a && !b && a.includes(":")) continue; // subsection header, e.g. "SP 2: SITE DESIGN"

    if (a && b) {
      topCode = a; topText = b; midMarker = midText = null;
      handleLeaf(topText, e, f, g, topCode);
    } else if (!a && b && c) {
      midMarker = b; midText = c;
      handleLeaf(`${topText} > ${midText}`, e, f, g, `${topCode} > ${midMarker}`);
    } else if (!a && !b && c && d) {
      handleLeaf(`${topText} > ${midText} > ${d}`, e, f, g, `${topCode} > ${midText} > ${c}`);
    }
  }

  return { items: matched, unmatched };
}

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
  { id: "Innovation",                       code: "INV" },
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
  { id: "es_4_4", pointNumber: "4.4", text: "Drywall sealed to top plate at all unconditioned attic/wall interfaces using caulk, foam, drywall adhesive, or equivalent", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_5", pointNumber: "4.5", text: "Rough opening around windows & exterior doors sealed", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_6", pointNumber: "4.6", text: "Assemblies separating attached garages from occupiable space sealed; air barrier installed, sealed, and aligned", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_7", pointNumber: "4.7", text: "Doors adjacent to unconditioned space or ambient made substantially air-tight with doorsweep and weatherstripping", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_8", pointNumber: "4.8", text: "Attic access panels, roof hatches and drop-down stairs gasketed or equipped with durable gasketed covers", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_4_9", pointNumber: "4.9", text: "Unit entrance doors from corridor/stairwell made substantially air-tight with doorsweep and weatherstripping", category: "High Performance Building Envelope", mandatory: true },
  { id: "es_5_9", pointNumber: "5.9", text: "All heating and cooling systems serving a dwelling unit have thermostatic controls within the unit", category: "Energy Efficient Systems", mandatory: true },

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
  // SP 2.7 removed 2026-08-20: same bug as the V7 array had -- genuinely optional (already
  // correctly tracked as ec_opt_sp_sp_2_7 in EARTHCRAFT_OPTIONAL_LIBRARY, version-agnostic so
  // it matches a V6.5 or V7 workbook upload either way), not mandatory.
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
  { id: "ec_es5_1",      pointNumber: "ES 5.1",  tier: "ALL", text: "Heat trap on all storage water heaters; confirm by visual inspection or AHRI certificate", category: "Energy Efficient Systems", mergedWith: ["es_11_3"] },
  { id: "ec_es5_3",      pointNumber: "ES 5.3",  tier: "ALL", text: "Pipe insulation on first 2' of hot and cold water pipes at water heater", category: "Energy Efficient Systems" },
  // ── WATER EFFICIENCY ────────────────────────────────────────────────────────
  { id: "ec_we1_0",      pointNumber: "WE 1.0",  tier: "ALL", text: "Meet National Energy Policy Act low-flow standards for all fixtures", category: "Water Efficiency" },
  { id: "ec_we1_1",      pointNumber: "WE 1.1",  tier: "ALL", text: "Detect and repair all leaks at water-using fixtures, appliances, and equipment", category: "Water Efficiency" },
  { id: "ec_we1_2",      pointNumber: "WE 1.2",  tier: "ALL", text: "Low-flow fixtures: WaterSense toilet ≤1.28 gpf; urinal ≤0.5 gpf; lavatory faucet ≤1.5 gpm; showerhead ≤2.0 gpm", category: "Water Efficiency", mergedWith: ["es_13_2"] },
  { id: "ec_du2_6",      pointNumber: "WE 2.3",  tier: "ALL", text: "Newly installed and existing plants maintain distance ≥2' from building at maturity", category: "Water Efficiency" },
  // ── Added 2026-08-20: carried over from the V7 mandatory-checklist audit (see CLAUDE.md) ───
  { id: "ec_v65_sp_sp_3_8", pointNumber: "SP 3.8", text: "Label all storm drains or storm inlets to discourage dumping of pollutants", category: "Site Planning", tier: "ALL" },
  { id: "ec_v65_sp_sp_3_9", pointNumber: "SP 3.9", text: "Road/vehicle cleaning protocols posted and enforced", category: "Site Planning", tier: "ALL" },
  { id: "ec_v65_du_du_1", pointNumber: "DU 1", text: "All roof valleys direct water away from walls, dormers, chimneys, etc.", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v65_du_du_1_1", pointNumber: "DU 1.1", text: "Install drainage plane per manufacturer's specifications", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v65_du_du_1_2_1", pointNumber: "DU 1.2 > 1", text: "Integrate drainage plane with: > Window and door pan flashing at sills and side flashing", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v65_du_du_1_2_2", pointNumber: "DU 1.2 > 2", text: "Integrate drainage plane with: > Window and door head/top flashing", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v65_du_du_1_3", pointNumber: "DU 1.3", text: "Double layer of building paper or house wrap behind cementitious stucco, stone veneer or synthetic stone veneer on framed walls", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v65_du_du_1_4", pointNumber: "DU 1.4", text: "Roof gutters discharge water ≥5' from foundation", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v65_du_du_1_5_1", pointNumber: "DU 1.5 > 1", text: "Flashing: > Self-sealing bituminous membrane or equivalent at valleys and roof deck penetrations", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v65_du_du_1_5_2", pointNumber: "DU 1.5 > 2", text: "Flashing: > Step and kick-out flashing at wall/roof and wall/porch intersections, flashing ≥4” on wall surface and integrated with wall and roof/deck/porch drainage planes", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v65_du_du_2", pointNumber: "DU 2", text: "Gravel bed (57's, no fines) beneath sub-grade slabs, on grade slabs, or raised slabs", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v65_du_du_2_1", pointNumber: "DU 2.1", text: "100% coverage of ≥6mil vapor barrier beneath all slabs, in all crawlspaces", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v65_du_du_2_2", pointNumber: "DU 2.2", text: "Foundation drain on top of sub-grade footing", category: "Durability & Moisture Management", tier: "ALL" },
  { id: "ec_v65_be_be_1_1", pointNumber: "BE 1.1", text: "Seal bottom plates to subfloor or foundation for entire unit envelope", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_5_1", pointNumber: "BE 1.5 > 1", text: "Seal penetrations through: > Foundations and exterior wall assemblies", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_5_2", pointNumber: "BE 1.5 > 2", text: "Seal penetrations through: > Top and bottom plates", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_5_5", pointNumber: "BE 1.5 > 5", text: "Seal penetrations through: > Sheathing", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_5_7", pointNumber: "BE 1.5 > 7", text: "Seal penetrations through: > All ceilings", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_6_1", pointNumber: "BE 1.6 > 1", text: "Seal penetrations around: > Shower, sinks, toilets and tub drains", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_6_2", pointNumber: "BE 1.6 > 2", text: "Seal penetrations around: > HVAC supply and return boots sealed to subfloor or drywall (floor, walls, or ceilings)", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_6_3", pointNumber: "BE 1.6 > 3", text: "Seal penetrations around: > Window and door rough openings", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_6_4", pointNumber: "BE 1.6 > 4", text: "Seal penetrations around: > All drywall penetrations (common walls between attached units included)", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_6_5", pointNumber: "BE 1.6 > 5", text: "Seal penetrations around: > Exhaust fans to drywall", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_6_6", pointNumber: "BE 1.6 > 6", text: "Seal penetrations around: > Attic pull-down stairs, scuttle holes and kneewall doors", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_6_7", pointNumber: "BE 1.6 > 7", text: "Seal penetrations around: > Chases", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_7_1", pointNumber: "BE 1.7 > 1", text: "Seal seams and gaps in: > Band joist sheathing", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_7_2", pointNumber: "BE 1.7 > 2", text: "Seal seams and gaps in: > Exterior wall sheathing", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_8_1", pointNumber: "BE 1.8 > 1", text: "Install rigid air barriers: > Behind tubs and showers on insulated walls", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_9_1", pointNumber: "BE 1.9 > 1", text: "Install weather-stripping at: > All exterior doors (if not included in door assembly)", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_10", pointNumber: "BE 1.10", text: "All recessed can lights must be air tight, gasketed at all floors and also IC-rated in insulated ceilings;  in Climate Zone 4, insulate exterior surface of fixture to ≥R-10", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_1_12", pointNumber: "BE 1.12", text: "Units adjacent to CMU walls: framing and sub-floor at unit envelope, including interstitial space, must be sealed to CMU", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_1", pointNumber: "BE 3 > 1", text: "Floors: > Framed ≥ R-19", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_2", pointNumber: "BE 3 > 2", text: "Floors: > Cantilevered ≥ R-30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_3", pointNumber: "BE 3 > 3", text: "Floors: > Podium/Elevated Slab ≥ R-19", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_1_1", pointNumber: "BE 3.1 > 1", text: "Walls: > Exterior walls and band joists ≥ R-15", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_1_2", pointNumber: "BE 3.1 > 2", text: "Walls: > Elevator walls adjacent to dwelling units ≥ R-13", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_1_3", pointNumber: "BE 3.1 > 3", text: "Walls: > Foundation walls ≥ R-10 continuous or ≥ R-13 cavity; Climate Zone 2/3 ≥ R-5 continuous or ≥ R-13 cavity; Climate Zone 4 ≥ R-10 continuous or ≥ R-13 cavity", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_2_1", pointNumber: "BE 3.2 > 1", text: "Ceilings/Roof: > Vented: Climate Zone 4 ≥ R-49", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_2_2", pointNumber: "BE 3.2 > 2", text: "Ceilings/Roof: > Continuous Roof Deck: Climate Zone 4 ≥ R-30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_2_3", pointNumber: "BE 3.2 > 3", text: "Ceilings/Roof: > Cathedral: Climate Zone 4 ≥ R-38", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_3_1", pointNumber: "BE 3.3 > 1", text: "Attic/Roof: > Install wind baffles at eaves in every vented bay, or equivalent air barrier at edge of ceiling", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_3_2", pointNumber: "BE 3.3 > 2", text: "Attic/Roof: > Energy heel trusses or raised top plate", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_3_3", pointNumber: "BE 3.3 > 3", text: "Attic/Roof: > Attic platforms allow for full-depth insulation below", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_5", pointNumber: "BE 3.5", text: "Attic pull-down/scuttle hole ≥ R-49", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_3_7", pointNumber: "BE 3.7", text: "Steel framed buildings require thermal break ≥ R-7.5", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_4_1", pointNumber: "BE 4 > 1", text: "Door U-factors and SHGC: > U-factor ≤0.35", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_4_2", pointNumber: "BE 4 > 2", text: "Door U-factors and SHGC: > SHGC ≤ 0.30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_4_1_1", pointNumber: "BE 4.1 > 1", text: "Window U-factor and SHGC: > U-factor ≤0.35", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_4_1_2", pointNumber: "BE 4.1 > 2", text: "Window U-factor and SHGC: > SHGC ≤ 0.30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_4_2_1", pointNumber: "BE 4.2 > 1", text: "Skylight U-factor and SHGC: > U-factor ≤0.55", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_be_be_4_2_2", pointNumber: "BE 4.2 > 2", text: "Skylight U-factor and SHGC: > SHGC ≤ 0.30", category: "High Performance Building Envelope", tier: "ALL" },
  { id: "ec_v65_es_es_1_1", pointNumber: "ES 1.1", text: "If programmable thermostat installed for heat pump, include adaptive recovery technology", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_2", pointNumber: "ES 2", text: "Seal air handlers and duct systems with mastic", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_2_1", pointNumber: "ES 2.1", text: "Install ducts per ACCA Manual D duct design", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_2_2", pointNumber: "ES 2.2", text: "Fully duct all supply and return ducts", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_2_3_1", pointNumber: "ES 2.3 > 1", text: "Duct insulation: > ≥ R-6: Ducts in conditioned and interstitial spaces (between floors)", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_2_3_2", pointNumber: "ES 2.3 > 2", text: "Duct insulation: > ≥ R-8: Ducts in unconditioned space", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_2_4", pointNumber: "ES 2.4", text: "No ducts in exterior walls or vaulted ceilings and no plenum within 2' of roofline.", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_2_5", pointNumber: "ES 2.5", text: "Locate all air handlers within conditioned space", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_2_6", pointNumber: "ES 2.6", text: "Indoor coil protected until finished floor installed", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_2_8", pointNumber: "ES 2.8", text: "No duct take-offs within 6\" of supply plenum or supply trunk cap", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_2_9", pointNumber: "ES 2.9", text: "Design and construct mechanical closets accessible for service and maintenance requirements", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_4", pointNumber: "ES 4", text: "Install exhaust fans in all bathrooms and duct to outside", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_4_1", pointNumber: "ES 4.1", text: "Gas kitchen range vented to exterior  ≥100 cfm fan", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_4_3_2", pointNumber: "ES 4.3 > 2", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > ≥ 2' above grade", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_4_3_4", pointNumber: "ES 4.3 > 4", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > Fresh air duct may not be run to the roof", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_4_3_5", pointNumber: "ES 4.3 > 5", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > Fresh air shutoff may not be controlled by humidistat", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_4_3_6", pointNumber: "ES 4.3 > 6", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > Install rigid duct with insulation", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_4_3_7", pointNumber: "ES 4.3 > 7", text: "When installed to achieve ES 4.2, design and install fresh air intakes: > All intakes must be ducted to exterior of building", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_4_4", pointNumber: "ES 4.4", text: "Seal seams of all intake and exhaust ducts with mastic", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_4_5", pointNumber: "ES 4.5", text: "Duct clothes dryers to outside", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_4_7", pointNumber: "ES 4.7", text: "Back-draft dampers for kitchen and bathroom exhaust", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_5_0", pointNumber: "ES 5.0", text: "Water Heater must be installed in conditioned space. If gas, direct vent", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_6", pointNumber: "ES 6", text: "High-efficacy lighting in 100% of all permanent fixtures", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_6_1", pointNumber: "ES 6.1", text: "If installed, ENERGY STAR dishwasher", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_es_es_6_2", pointNumber: "ES 6.2", text: "If installed, ENERGY STAR refrigerator", category: "Energy Efficient Systems", tier: "ALL" },
  { id: "ec_v65_we_we_2", pointNumber: "WE 2", text: "Cover all exposed soil with 2\"-3\" mulch layer", category: "Water Efficiency", tier: "ALL" },
  { id: "ec_v65_we_we_2_1_1", pointNumber: "WE 2.1 > 1", text: "Irrigation system: > Must have rain sensor shutoff switch", category: "Water Efficiency", tier: "ALL" },
];

const EARTHCRAFT_GOLD_V6 = [
  ...EARTHCRAFT_CERTIFIED_V6,
  // ── RESOURCE EFFICIENCY: GOLD ───────────────────────────────────────────────
  { id: "ec_v7_re1_2_1", pointNumber: "RE 1.2 > 1", tier: "GOLD", text: "2-stud corners where structurally feasible", category: "Resource Efficiency", points: 3 },
  { id: "ec_v7_re1_2_2", pointNumber: "RE 1.2 > 2", tier: "GOLD", text: "Ladder T-walls where structurally feasible", category: "Resource Efficiency", points: 2 },
  { id: "ec_v7_re1_2_3", pointNumber: "RE 1.2 > 3", tier: "GOLD", text: "Size headers for loads (non-structural headers in non-load bearing walls)", category: "Resource Efficiency", points: 1 },
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
  // ── Added 2026-08-20: carried over from the V7 mandatory-checklist audit (see CLAUDE.md) ───
  { id: "ec_v65_be_be_1_13", pointNumber: "BE 1.13", text: "Seal top plate to drywall at the attic level", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v65_be_be_4_4_1", pointNumber: "BE 4.4 > 1", text: "Door U-factor: > Opaque door:  U factor≤ 0.21", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v65_be_be_4_4_2", pointNumber: "BE 4.4 > 2", text: "Door U-factor: > Door with ≤ 50% glass:  U-factor ≤ 0.27", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v65_be_be_4_4_3", pointNumber: "BE 4.4 > 3", text: "Door U-factor: > Door with > 50% glass:  U-factor ≤ 0.32", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v65_be_be_4_5_1", pointNumber: "BE 4.5 > 1", text: "Window U-factor and SHGC: > U-factor ≤0.32", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v65_be_be_4_5_2", pointNumber: "BE 4.5 > 2", text: "Window U-factor and SHGC: > SHGC ≤0.27", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v65_be_be_4_6_1", pointNumber: "BE 4.6 > 1", text: "Skylight U-factor and SHGC: > U-factor ≤0.50", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v65_be_be_4_6_2", pointNumber: "BE 4.6 > 2", text: "Skylight U-factor and SHGC: > SHGC ≤0.27", category: "High Performance Building Envelope", tier: "GOLD" },
  { id: "ec_v65_es_es_2_11_1", pointNumber: "ES 2.11 > 1", text: "Minimize pressure imbalance within units: > Install fully ducted jumper ducts, transfer grills, or dedicated return for each bedroom", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v65_es_es_2_12", pointNumber: "ES 2.12", text: "Install rigid duct work or pull all flex ducts with no pinches and support at intervals ≤ 5’", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v65_es_es_2_15", pointNumber: "ES 2.15", text: "HVAC system and ductwork is dry and clean", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v65_es_es_1_8", pointNumber: "ES 1.8", text: "Heating equipment efficiency: ENERGY STAR qualified furnace(s) ≥95 AFUE and within 40% of load calculation, OR ENERGY STAR qualified heat pump(s) ≥8.5 HSPF and within 25% of load calculation", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v65_es_es_1_10", pointNumber: "ES 1.10", text: "ENERGY STAR qualified cooling equipment ≥SEER 15", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v65_es_es_4_8", pointNumber: "ES 4.8", text: "If installed, ceiling fans must be ENERGY STAR qualified (1/bedroom and 1 in living room)", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v65_es_es_4_12", pointNumber: "ES 4.12", text: "Install and label accessible ventilation controls, with override controls for continuously operating ventilation fans", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v65_es_es_4_13", pointNumber: "ES 4.13", text: "Supply/exhaust fans rated at ≤3 sones (intermittent) and ≤1 sone (continuous)", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v65_es_es_6_3", pointNumber: "ES 6.3", text: "If installed, ENERGY STAR qualified clothes washer", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v65_es_es_6_4", pointNumber: "ES 6.4", text: "If installed, high efficiency clothes dryer with moisture sensor (not applicable to commercial dryers)", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v65_es_es_7", pointNumber: "ES 7", text: "100% LED bulbs in all corridor/breezeway and all common spaces", category: "Energy Efficient Systems", tier: "GOLD" },
  { id: "ec_v65_be_be_5_0_a", pointNumber: "BE 5.0 > A.", text: "If Ducts located in unconditioned attic: > Attic Side Radiant Barrier", category: "High Performance Building Envelope", tier: "GOLD" }
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
  { id: "ec_v7_re1_2_1",  pointNumber: "RE 1.2 > 1", tier: "GOLD", text: "2-stud corners where structurally feasible", category: "Resource Efficiency", points: 3 },
  { id: "ec_v7_re1_2_2",  pointNumber: "RE 1.2 > 2", tier: "GOLD", text: "Ladder T-walls where structurally feasible", category: "Resource Efficiency", points: 2 },
  { id: "ec_v7_re1_2_3",  pointNumber: "RE 1.2 > 3", tier: "GOLD", text: "Size headers for loads (non-structural headers in non-load bearing walls)", category: "Resource Efficiency", points: 1 },
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

function isEarthCraftGoldSelected(programSelections) {
  return (programSelections || []).some(s => s.programId === "earthcraft_gold");
}

function getItemsForSelection(programSelections, categoryId, extraItems) {
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
  // Project-specific EarthCraft optional points, matched from an uploaded workbook — see
  // parseEarthCraftWorkbook. An item flagged goldMandatoryOverlap is already on this list as a
  // mandatory pass/fail item for Gold (see EARTHCRAFT_GOLD_V7/V6), so it's excluded here — live,
  // from the project's CURRENT program selections — to avoid showing it twice. Its workbook
  // status still isn't wasted: see applyEarthCraftGoldOverlapAutoPass, which auto-passes the
  // mandatory item instead.
  const goldSelected = isEarthCraftGoldSelected(programSelections);
  for (const item of extraItems || []) {
    if (item.goldMandatoryOverlap && goldSelected) continue;
    if (item.category === categoryId && !seen.has(item.id)) { seen.add(item.id); result.push({ ...item, sourceKey: "earthcraft_optional_import" }); }
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

// A workbook "y" status means the TA already reviewed and approved the item with backup
// documentation on file — not just planned. Auto-marks those as passing so the field TA
// isn't re-verifying something already confirmed. Never overwrites a record that already
// exists (e.g. a TA's own field check on a re-uploaded workbook takes precedence).
function applyEarthCraftAutoPass(project, existingRecords) {
  const updates = {};
  const goldSelected = isEarthCraftGoldSelected(project.programs);
  for (const item of project.earthcraftOptionalItems || []) {
    if (item.goldMandatoryOverlap && goldSelected) continue; // tracked as a mandatory item instead — see applyEarthCraftGoldOverlapAutoPass
    if ((item.workbookStatus || "").trim().toLowerCase() !== "y") continue;
    const key = `${project.id}__${item.category}__${item.id}`;
    if (existingRecords[key]) continue;
    updates[key] = { status: "pass", fromWorkbook: true, updatedAt: new Date().toISOString() };
  }
  return updates;
}

// A workbook "y" on an item that's ALSO mandatory for Gold (goldMandatoryOverlap — see
// EARTHCRAFT_OPTIONAL_LIBRARY) isn't tracked as its own optional point (see getItemsForSelection),
// so the "already verified" signal has nowhere to land unless we route it to the matching
// mandatory item instead. Matched by category + point number; falls back to the code before " > "
// for items like BE 3.10 that are split into sub-options (A/B) in the optional library but exist
// as a single combined item in the mandatory checklist.
function applyEarthCraftGoldOverlapAutoPass(project, existingRecords) {
  const updates = {};
  const goldSel = (project.programs || []).find(s => s.programId === "earthcraft_gold");
  if (!goldSel) return updates;
  const mandatoryItems = CHECKLIST_REGISTRY[`${goldSel.programId}||${goldSel.version}||${goldSel.revision}`] || [];
  const norm = s => (s || "").trim().toLowerCase();
  const base = s => norm(s).split(">")[0].trim();
  for (const item of project.earthcraftOptionalItems || []) {
    if (!item.goldMandatoryOverlap) continue;
    if ((item.workbookStatus || "").trim().toLowerCase() !== "y") continue;
    const match = mandatoryItems.find(m => m.category === item.category && norm(m.pointNumber) === norm(item.pointNumber))
      || mandatoryItems.find(m => m.category === item.category && base(m.pointNumber) === base(item.pointNumber));
    if (!match) continue;
    const key = `${project.id}__${match.category}__${match.id}`;
    if (existingRecords[key]) continue;
    updates[key] = { status: "pass", fromWorkbook: true, updatedAt: new Date().toISOString() };
  }
  return updates;
}

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
  let total = 0, verified = 0, fail = 0, mismatches = 0;
  CATEGORIES.forEach(cat => {
    const items = getItemsForSelection(project.programs || [], cat.id, project.earthcraftOptionalItems);
    items.forEach(item => {
      total++;
      const r = records[`${project.id}__${cat.id}__${item.id}`];
      if (r?.status === "pass" || r?.status === "na") verified++;
      if (r?.status === "fail") fail++;
      if (r?.modelMismatch) mismatches++;
    });
  });
  return { pct: total ? Math.round((verified / total) * 100) : 0, fail, total, verified, mismatches };
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
// Shared by project creation and project editing — editing pre-fills every field from
// initialProject, re-verifies the linked SharePoint folder still exists, and preserves
// the project's id/createdAt rather than minting a new project.
function ProjectForm({ initialProject, onSave, onBack, auth, setAuth }) {
  const isEdit = !!initialProject;
  const [name, setName] = useState(initialProject?.name || "");
  const [advisor, setAdvisor] = useState(initialProject?.advisor || "");
  const [step, setStep] = useState("name"); // name | programs | version
  const [selections, setSelections] = useState(initialProject?.programs || []); // [{programId, version, revision}]
  const [pickingProgram, setPickingProgram] = useState(null); // programId being configured
  const [folderPath, setFolderPath] = useState(initialProject?.sharePointFolder || "");
  const [folderStatus, setFolderStatus] = useState(null); // null | "checking" | "ok" | "error"
  const [folderMsg, setFolderMsg] = useState("");
  const [energyModel, setEnergyModel] = useState(initialProject?.energyModel || null);
  const [energyModelFileName, setEnergyModelFileName] = useState(initialProject?.energyModelFileName || "");
  const [energyModelUploadedAt, setEnergyModelUploadedAt] = useState(initialProject?.energyModelUploadedAt || null);
  const [energyModelError, setEnergyModelError] = useState("");
  const emFileRef = useRef();
  // Raw matched items from the workbook, unfiltered — this is what actually gets saved onto the
  // project. Gold-mandatory-overlap items are filtered OUT of the optional checklist display at
  // render time instead (see getItemsForSelection), not here and not at save time, so their
  // workbook status survives to drive applyEarthCraftGoldOverlapAutoPass regardless of upload/
  // program-selection order.
  const [earthcraftRawItems, setEarthcraftRawItems] = useState(initialProject?.earthcraftOptionalItems || null);
  const [earthcraftWorkbookFileName, setEarthcraftWorkbookFileName] = useState(initialProject?.earthcraftWorkbookFileName || "");
  const [earthcraftWorkbookUploadedAt, setEarthcraftWorkbookUploadedAt] = useState(initialProject?.earthcraftWorkbookUploadedAt || null);
  const [earthcraftError, setEarthcraftError] = useState("");
  const ecFileRef = useRef();

  // A Gold project must already do every Gold-tier item as a mandatory pass/fail check (see
  // EARTHCRAFT_GOLD_V7/V6) — don't also track it as a separate optional point, or it shows up
  // twice. Certified-only projects don't have that mandatory item at all, so it's a genuine
  // optional bonus point for them. Re-derived on every render so it stays correct regardless
  // of upload/program-selection order.
  const isGoldSelected = selections.some(s => s.programId === "earthcraft_gold");
  const earthcraftOptionalItems = isGoldSelected
    ? (earthcraftRawItems || []).filter(i => !i.goldMandatoryOverlap)
    : (earthcraftRawItems || []);
  const earthcraftAlreadyMandatoryCount = (earthcraftRawItems || []).length - earthcraftOptionalItems.length;

  const startAddProgram = () => setPickingProgram("choose");

  const handleEnergyModelFile = (e) => {
    const file = e.target.files[0]; e.target.value = ""; if (!file) return;
    setEnergyModelError("");
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        setEnergyModel(parseEkotropeXml(ev.target.result));
        setEnergyModelFileName(file.name);
        setEnergyModelUploadedAt(new Date().toISOString());
      } catch (err) {
        setEnergyModelError(err.message || "Could not parse this file.");
      }
    };
    reader.onerror = () => setEnergyModelError("Could not read this file.");
    reader.readAsText(file);
  };

  const removeEnergyModel = () => { setEnergyModel(null); setEnergyModelFileName(""); setEnergyModelUploadedAt(null); setEnergyModelError(""); };

  const handleEarthCraftFile = (e) => {
    const file = e.target.files[0]; e.target.value = ""; if (!file) return;
    setEarthcraftError("");
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const { items } = parseEarthCraftWorkbook(ev.target.result);
        setEarthcraftRawItems(items);
        setEarthcraftWorkbookFileName(file.name);
        setEarthcraftWorkbookUploadedAt(new Date().toISOString());
      } catch (err) {
        setEarthcraftError(err.message || "Could not parse this file.");
      }
    };
    reader.onerror = () => setEarthcraftError("Could not read this file.");
    reader.readAsArrayBuffer(file);
  };

  const removeEarthcraftWorkbook = () => {
    setEarthcraftRawItems(null); setEarthcraftWorkbookFileName(""); setEarthcraftWorkbookUploadedAt(null);
    setEarthcraftError("");
  };

  const confirmVersionRevision = (programId, version, revision) => {
    setSelections(s => [...s, { programId, version, revision }]);
    setPickingProgram(null);
  };

  const removeSelection = (idx) => setSelections(s => s.filter((_, i) => i !== idx));

  const verifyFolder = async () => {
    const path = folderPath.trim();
    if (!path) { setFolderStatus(null); setFolderMsg(""); return; }
    if (!auth) { setFolderStatus("error"); setFolderMsg("Connect to SharePoint first to verify this folder."); return; }
    setFolderStatus("checking"); setFolderMsg("");
    try {
      const token = await getValidToken(auth, setAuth);
      if (!token) throw new Error("Could not get a valid SharePoint session.");
      const siteId = await getSharePointSiteId(token);
      await getSharePointFolderId(siteId, token, path);
      setFolderStatus("ok"); setFolderMsg("Folder found");
    } catch (e) {
      setFolderStatus("error"); setFolderMsg(e.message || "Folder not found.");
    }
  };

  // Editing an existing link — re-check it's still valid rather than assuming it still is
  // (this screen exists specifically because a linked folder can be renamed/moved later).
  useEffect(() => {
    if (isEdit && folderPath.trim() && auth) { verifyFolder(); }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (step === "name") {
    return (
      <div style={{ padding: "24px 20px" }}>
        <h2 style={{ margin: "0 0 24px", fontSize: 20, fontWeight: 700, color: "#111827" }}>{isEdit ? "Edit project" : "New project"}</h2>
        <label style={{ fontSize: 12, fontWeight: 700, color: "#6B7280", letterSpacing: "0.06em", textTransform: "uppercase" }}>Project name</label>
        <input value={name} onChange={e => setName(e.target.value)}
          placeholder="e.g. Green Park"
          style={{ display: "block", width: "100%", marginTop: 8, padding: "12px 14px", fontSize: 16, border: "1.5px solid #E5E7EB", borderRadius: 10, outline: "none", boxSizing: "border-box", fontFamily: "DM Sans, sans-serif" }}/>
        <label style={{ display: "block", marginTop: 20, fontSize: 12, fontWeight: 700, color: "#6B7280", letterSpacing: "0.06em", textTransform: "uppercase" }}>Technical Advisor</label>
        <input value={advisor} onChange={e => setAdvisor(e.target.value)}
          placeholder="Full name"
          style={{ display: "block", width: "100%", marginTop: 8, padding: "12px 14px", fontSize: 16, border: "1.5px solid #E5E7EB", borderRadius: 10, outline: "none", boxSizing: "border-box", fontFamily: "DM Sans, sans-serif" }}/>

        <label style={{ display: "block", marginTop: 20, fontSize: 12, fontWeight: 700, color: "#6B7280", letterSpacing: "0.06em", textTransform: "uppercase" }}>SharePoint Site Visits folder</label>
        {!auth && (
          <div style={{ marginTop: 8, padding: "10px 12px", background: "#F9FAFB", border: "1.5px solid #E5E7EB", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
            <p style={{ margin: 0, fontSize: 12, color: "#9CA3AF" }}>Connect to SharePoint to link a folder</p>
            <button onClick={startLogin}
              style={{ fontSize: 12, fontWeight: 600, color: "#FFF", background: "#0078D4", border: "none", borderRadius: 6, padding: "6px 12px", cursor: "pointer", flexShrink: 0, fontFamily: "DM Sans, sans-serif" }}>
              Connect
            </button>
          </div>
        )}
        <input value={folderPath} disabled={!auth}
          onChange={e => { setFolderPath(e.target.value); setFolderStatus(null); setFolderMsg(""); }}
          onBlur={verifyFolder}
          placeholder="e.g. Projects/123 Main St/Site Visits"
          style={{ display: "block", width: "100%", marginTop: 8, padding: "12px 14px", fontSize: 16, border: `1.5px solid ${folderStatus==="error"?"#EF4444":folderStatus==="ok"?"#10B981":"#E5E7EB"}`, borderRadius: 10, outline: "none", boxSizing: "border-box", fontFamily: "DM Sans, sans-serif", background: !auth ? "#F9FAFB" : "#FFF" }}/>
        {folderStatus === "checking" && <p style={{ margin: "6px 0 0", fontSize: 12, color: "#9CA3AF" }}>Checking…</p>}
        {folderStatus === "ok" && <p style={{ margin: "6px 0 0", fontSize: 12, color: "#10B981" }}>✓ {folderMsg}</p>}
        {folderStatus === "error" && <p style={{ margin: "6px 0 0", fontSize: 12, color: "#EF4444" }}>{folderMsg}</p>}

        <label style={{ display: "block", marginTop: 20, fontSize: 12, fontWeight: 700, color: "#6B7280", letterSpacing: "0.06em", textTransform: "uppercase" }}>
          Energy model <span style={{ fontWeight: 400, color: "#9CA3AF", textTransform: "none" }}>(optional — Ekotrope .xml export)</span>
        </label>
        {energyModelFileName ? (
          <div style={{ marginTop: 8, padding: "10px 12px", background: "#F0FDF4", border: "1.5px solid #10B981", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
            <div style={{ minWidth: 0 }}>
              <p style={{ margin: 0, fontSize: 12.5, color: "#065F46", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>✓ {energyModelFileName}</p>
              {energyModelUploadedAt && <p style={{ margin: "2px 0 0", fontSize: 11, color: "#059669" }}>Uploaded {fmtDate(energyModelUploadedAt)}</p>}
            </div>
            <button onClick={removeEnergyModel}
              style={{ fontSize: 11, color: "#065F46", background: "none", border: "1px solid #A7F3D0", borderRadius: 6, padding: "4px 10px", cursor: "pointer", flexShrink: 0, fontFamily: "DM Sans, sans-serif" }}>
              Remove
            </button>
          </div>
        ) : (
          <button onClick={() => emFileRef.current.click()}
            style={{ marginTop: 8, width: "100%", padding: "12px", border: "1.5px dashed #D1D5DB", borderRadius: 10, background: "#F9FAFB", color: "#6B7280", fontSize: 13, cursor: "pointer", fontFamily: "DM Sans, sans-serif" }}>
            + Upload energy model
          </button>
        )}
        {energyModelError && <p style={{ margin: "6px 0 0", fontSize: 12, color: "#EF4444" }}>{energyModelError}</p>}
        <input ref={emFileRef} type="file" accept=".xml" onChange={handleEnergyModelFile} style={{ display: "none" }}/>
        <p style={{ margin: "6px 0 0", fontSize: 11.5, color: "#9CA3AF" }}>Lets Technical Advisors see the modeled wall/ceiling/foundation/window/equipment values in the Minimum Rated Features section, and flag anything that doesn't match in the field.</p>

        <label style={{ display: "block", marginTop: 20, fontSize: 12, fontWeight: 700, color: "#6B7280", letterSpacing: "0.06em", textTransform: "uppercase" }}>
          EarthCraft optional points <span style={{ fontWeight: 400, color: "#9CA3AF", textTransform: "none" }}>(optional — populated EarthCraft workbook .xlsx)</span>
        </label>
        {earthcraftWorkbookFileName ? (
          <div style={{ marginTop: 8, padding: "10px 12px", background: "#F0FDF4", border: "1.5px solid #10B981", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
            <div style={{ minWidth: 0 }}>
              <p style={{ margin: 0, fontSize: 12.5, color: "#065F46", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>✓ {earthcraftWorkbookFileName}</p>
              <p style={{ margin: "2px 0 0", fontSize: 11, color: "#059669" }}>
                {(earthcraftOptionalItems||[]).length} planned point{(earthcraftOptionalItems||[]).length===1?"":"s"} added
                {(earthcraftOptionalItems||[]).filter(i => (i.workbookStatus||"").toLowerCase()==="y").length > 0 &&
                  <> ({(earthcraftOptionalItems||[]).filter(i => (i.workbookStatus||"").toLowerCase()==="y").length} already marked passing from workbook)</>}
                {earthcraftWorkbookUploadedAt && <> · Uploaded {fmtDate(earthcraftWorkbookUploadedAt)}</>}
              </p>
            </div>
            <button onClick={removeEarthcraftWorkbook}
              style={{ fontSize: 11, color: "#065F46", background: "none", border: "1px solid #A7F3D0", borderRadius: 6, padding: "4px 10px", cursor: "pointer", flexShrink: 0, fontFamily: "DM Sans, sans-serif" }}>
              Remove
            </button>
          </div>
        ) : (
          <button onClick={() => ecFileRef.current.click()}
            style={{ marginTop: 8, width: "100%", padding: "12px", border: "1.5px dashed #D1D5DB", borderRadius: 10, background: "#F9FAFB", color: "#6B7280", fontSize: 13, cursor: "pointer", fontFamily: "DM Sans, sans-serif" }}>
            + Upload EarthCraft workbook
          </button>
        )}
        {earthcraftError && <p style={{ margin: "6px 0 0", fontSize: 12, color: "#EF4444" }}>{earthcraftError}</p>}
        <input ref={ecFileRef} type="file" accept=".xlsx" onChange={handleEarthCraftFile} style={{ display: "none" }}/>
        <p style={{ margin: "6px 0 0", fontSize: 11.5, color: "#9CA3AF" }}>Adds this project's planned optional points as trackable checklist items, matched against our reviewed field-verifiable list.</p>
        {earthcraftAlreadyMandatoryCount > 0 && (
          <div style={{ marginTop: 8, padding: "10px 12px", background: "#EFF6FF", border: "1.5px solid #BFDBFE", borderRadius: 10 }}>
            <p style={{ margin: 0, fontSize: 11.5, fontWeight: 700, color: "#1D4ED8" }}>{earthcraftAlreadyMandatoryCount} item{earthcraftAlreadyMandatoryCount===1?"":"s"} already tracked as mandatory</p>
            <p style={{ margin: "4px 0 0", fontSize: 11, color: "#1D4ED8", lineHeight: 1.5 }}>These are required for EarthCraft Gold, so they're already on this project's checklist as pass/fail — not added again as separate optional points.</p>
          </div>
        )}
        {/* TEMP TEST-ONLY: SharePoint folder requirement dropped below so this preview branch is
            usable without OAuth (the app's redirect URI isn't authorized for preview domains).
            REVERT before merging to main — see git history on this branch. */}
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

      <button onClick={() => selections.length && onSave({
          id: initialProject?.id || Date.now().toString(),
          name: name.trim(), advisor: advisor.trim(), programs: selections, sharePointFolder: folderPath.trim(),
          energyModel, energyModelFileName, energyModelUploadedAt,
          earthcraftOptionalItems: earthcraftRawItems || [], earthcraftWorkbookFileName, earthcraftWorkbookUploadedAt,
          createdAt: initialProject?.createdAt || new Date().toISOString(),
        })}
        disabled={!selections.length}
        style={{ marginTop: 20, width: "100%", padding: 14, background: !selections.length?"#E5E7EB":"#1B4332", color: !selections.length?"#9CA3AF":"#FFF", border: "none", borderRadius: 12, fontSize: 15, fontWeight: 600, cursor: !selections.length?"not-allowed":"pointer", fontFamily: "DM Sans, sans-serif" }}>
        {isEdit ? "Save changes" : "Create project"}
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
            {item.points != null && (
              <span style={{ fontSize: 10, fontWeight: 700, color: "#92400E", background: "#FEF3C7", padding: "1px 7px", borderRadius: 20 }}>★ {item.points} pt{item.points===1?"":"s"}</span>
            )}
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
            {rec.modelMismatch && <span style={{ fontSize: 10, fontWeight: 600, color: "#991B1B", background: "#FEF2F2", padding: "1px 6px", borderRadius: 20 }}>⚡ model mismatch</span>}
            {rec.fromWorkbook && <span style={{ fontSize: 10, fontWeight: 600, color: "#1D4ED8", background: "#EFF6FF", padding: "1px 6px", borderRadius: 20 }}>📄 from workbook</span>}
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
function ProjectDashboard({ project, records, onSelectCategory, onSelectItem, onEdit, auth, setAuth, updateRecord }) {
  const pg = calcProjectProgress(project, records);
  const [query, setQuery] = useState("");
  const q = query.trim().toLowerCase();
  const [syncState, setSyncState] = useState({ running: false, done: 0, total: 0, errors: [] });

  // All items across every category, tagged with their source category
  const allProjectItems = CATEGORIES.flatMap(cat =>
    getItemsForSelection(project.programs||[], cat.id, project.earthcraftOptionalItems).map(i => ({ ...i, _cat: cat.id }))
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
    if (!project.sharePointFolder) { setSyncState({ running: false, done: 0, total: 0, errors: ["This project isn't linked to a SharePoint folder yet."] }); return; }
    if (!auth) { startLogin(); return; }
    const token = await getValidToken(auth, setAuth);
    if (!token) { setSyncState({ running: false, done: 0, total: 0, errors: ["Could not connect to SharePoint — please reconnect and try again."] }); return; }
    const jobs = pendingPhotoJobs();
    if (!jobs.length) return;
    setSyncState({ running: true, done: 0, total: jobs.length, errors: [] });

    // Resolve (or create) today's inspection-date subfolder inside this project's Site Visits
    // folder ONCE for the whole batch — every photo in this run lands in the same folder.
    let siteId, dateFolderId;
    try {
      siteId = await getSharePointSiteId(token);
      const siteVisitsFolderId = await getSharePointFolderId(siteId, token, project.sharePointFolder);
      dateFolderId = await createOrGetSubfolder(siteId, token, siteVisitsFolderId, formatDateFolderName(new Date()));
    } catch (e) {
      setSyncState({ running: false, done: 0, total: jobs.length, errors: [e.message] });
      return;
    }

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
        await uploadPhotoToFolder(siteId, token, dateFolderId, fileName, dataUrl);
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
    const items = getItemsForSelection(project.programs||[], cat.id, project.earthcraftOptionalItems);
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
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 16, minWidth: 0 }}>
            <div style={{ position: "relative", flexShrink: 0 }}>
              <ProgressRing pct={pg.pct} size={72} stroke={6} fail={pg.fail}/>
              <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: 700, color: pg.fail>0?"#EF4444":pg.pct===100?"#10B981":"#60A5FA" }}>{pg.pct}%</div>
            </div>
            <div style={{ minWidth: 0 }}>
              <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: "#FFF", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{project.name}</h2>
              <p style={{ margin: "3px 0 0", fontSize: 12, color: "#A7F3D0" }}>{pg.verified}/{pg.total} items verified</p>
              {project.advisor && <p style={{ margin: "2px 0 0", fontSize: 12, color: "#6EE7B7" }}>TA: {project.advisor}</p>}
              {pg.fail>0 && <p style={{ margin: "2px 0 0", fontSize: 12, color: "#FCA5A5", fontWeight: 600 }}>⚠ {pg.fail} item{pg.fail>1?"s":""} failing</p>}
              {pg.mismatches>0 && <p style={{ margin: "2px 0 0", fontSize: 12, color: "#FCA5A5", fontWeight: 600 }}>⚡ {pg.mismatches} model mismatch{pg.mismatches>1?"es":""}</p>}
            </div>
          </div>
          <button onClick={onEdit} title="Edit project"
            style={{ flexShrink: 0, width: 32, height: 32, borderRadius: 8, border: "1px solid rgba(255,255,255,0.3)", background: "rgba(255,255,255,0.1)", color: "#FFF", fontSize: 14, cursor: "pointer" }}>
            ✎
          </button>
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
            📁 {project.sharePointFolder ? `${project.sharePointFolder}/${formatDateFolderName(new Date())}` : "Not linked to a SharePoint folder"}
          </p>
          <p style={{ margin: "2px 0 0", fontSize: 11, color: "#9CA3AF" }}>
            {!project.sharePointFolder ? "Set up by your Project Manager when the project is created"
              : syncState.running ? `Uploading ${syncState.done}/${syncState.total}…`
              : !auth ? "Not connected"
              : pendingCount === 0 ? "All photos synced"
              : `${pendingCount} photo${pendingCount>1?"s":""} pending`}
          </p>
          {!syncState.running && syncState.errors.length>0 && (
            <p style={{ margin: "2px 0 0", fontSize: 11, color: "#EF4444" }}>{syncState.errors.length} failed — tap to retry</p>
          )}
        </div>
        <button onClick={handleUploadToSharePoint} disabled={!project.sharePointFolder || syncState.running || (auth && pendingCount===0)}
          style={{ fontSize: 12, fontWeight: 600, color: "#FFF", background: !project.sharePointFolder ? "#D1D5DB" : !auth ? "#0078D4" : (pendingCount===0 ? "#D1D5DB" : "#059669"), border: "none", borderRadius: 6, padding: "6px 14px", cursor: (!project.sharePointFolder || syncState.running) ? "not-allowed" : "pointer", flexShrink: 0, fontFamily: "DM Sans, sans-serif" }}>
          {!project.sharePointFolder ? "Not linked" : !auth ? "Connect" : syncState.running ? "Uploading…" : pendingCount===0 ? "Synced" : "Upload to SharePoint"}
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
  const allItems = getItemsForSelection(project.programs||[], category.id, project.earthcraftOptionalItems).map(i => ({ ...i, _cat: category.id }));
  const [query, setQuery] = useState("");
  const [modelNotesOpen, setModelNotesOpen] = useState(false);
  const p = calcCatProgress(allItems, records, project.id, category.id);
  const q = query.trim().toLowerCase();
  const isMRF = category.id === "Minimum Rated Features";
  const modelNotes = isMRF ? project.energyModel?.notes : null;

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
      {modelNotes && (
        <div style={{ padding: "10px 20px", borderBottom: "1px solid #F3F4F6", background: "#EFF6FF" }}>
          <button onClick={() => setModelNotesOpen(o => !o)}
            style={{ display: "flex", alignItems: "center", gap: 6, width: "100%", background: "none", border: "none", padding: 0, cursor: "pointer", fontFamily: "DM Sans, sans-serif" }}>
            <span style={{ fontSize: 11, fontWeight: 700, color: "#1D4ED8", textTransform: "uppercase", letterSpacing: "0.06em" }}>⚡ Energy model notes</span>
            <span style={{ fontSize: 10, color: "#1D4ED8", transform: modelNotesOpen ? "rotate(180deg)" : "none" }}>▾</span>
          </button>
          {modelNotesOpen && (
            <p style={{ margin: "8px 0 0", fontSize: 12.5, color: "#1E3A8A", lineHeight: 1.6, whiteSpace: "pre-wrap" }}>{modelNotes}</p>
          )}
        </div>
      )}
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
  const mismatchRef = useRef(!!record?.modelMismatch);
  const mismatchNoteRef = useRef(record?.modelMismatchNote || "");

  // Energy-model reference lines for this item, if one was uploaded and this item has a
  // known comparison point (see MRF_MODEL_FIELDS) — null otherwise, incl. non-MRF items.
  const modelRefLines = project.energyModel ? MRF_MODEL_FIELDS[item.id]?.(project.energyModel) : null;
  const [mismatch, setMismatch] = useState(!!record?.modelMismatch);
  const [mismatchNote, setMismatchNote] = useState(record?.modelMismatchNote || "");

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
    // eslint-disable-next-line react-hooks/exhaustive-deps
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
      modelMismatch: mismatchRef.current,
      modelMismatchNote: mismatchNoteRef.current,
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
      ? { status: record.status, note: record.note||"", mismatch: !!record.modelMismatch, mismatchNote: record.modelMismatchNote||"", updatedAt: record.updatedAt }
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

  const handleMismatchToggle = () => {
    const next = !mismatch;
    // Log every flag/unflag with a timestamp — lets you tell whether a mismatch was marked
    // before or after the model was last updated (see the model's upload date above).
    const archive = { status: record?.status||"", note: record?.note||"", mismatch: !!record?.modelMismatch, mismatchNote: record?.modelMismatchNote||"", updatedAt: record?.updatedAt };
    setMismatch(next);
    mismatchRef.current = next;
    save({ archive });
  };

  const handleMismatchNoteChange = (val) => {
    setMismatchNote(val);
    mismatchNoteRef.current = val;
  };

  const handleMismatchNoteBlur = () => save();

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
    const archive = changed ? { status: statusRef.current, note: noteSnapshot.current.note, mismatch: !!record?.modelMismatch, mismatchNote: record?.modelMismatchNote||"", updatedAt: noteSnapshot.current.updatedAt || record?.updatedAt } : undefined;
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
        {item.points != null && (
          <span style={{ display: "inline-block", marginBottom: 8, marginLeft: 6, fontSize: 12, fontWeight: 700, color: "#92400E", background: "#FEF3C7", padding: "2px 8px", borderRadius: 6 }}>
            ★ {item.points} pt{item.points===1?"":"s"}
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
          {record?.fromWorkbook && <span style={{ fontSize: 10, padding: "1px 7px", borderRadius: 20, background: "#EFF6FF", color: "#1D4ED8", fontWeight: 600 }}>📄 from workbook</span>}
        </div>
      </div>

      {/* Energy model reference — what the Ekotrope model assumes for this item */}
      {modelRefLines && modelRefLines.length > 0 && (
        <div style={{ marginBottom: 24, padding: "12px 14px", background: "#EFF6FF", border: "1px solid #BFDBFE", borderRadius: 10 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 8 }}>
            <p style={{ margin: "0 0 6px", fontSize: 11, fontWeight: 700, color: "#1D4ED8", textTransform: "uppercase", letterSpacing: "0.06em" }}>
              ⚡ Energy model says
            </p>
            {project.energyModelUploadedAt && <span style={{ fontSize: 10.5, color: "#60A5FA" }}>as of {fmtDate(project.energyModelUploadedAt)}</span>}
          </div>
          {modelRefLines.map((line, i) => (
            <p key={i} style={{ margin: i===0 ? 0 : "3px 0 0", fontSize: 13, color: "#1E3A8A", lineHeight: 1.5 }}>{line}</p>
          ))}
          <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 10, cursor: "pointer" }}>
            <input type="checkbox" checked={mismatch} onChange={handleMismatchToggle}
              style={{ width: 16, height: 16, cursor: "pointer", accentColor: "#EF4444" }}/>
            <span style={{ fontSize: 12.5, fontWeight: 600, color: mismatch ? "#991B1B" : "#1E3A8A" }}>
              Field doesn't match the model — flag for the energy modeler
            </span>
          </label>
          {mismatch && (
            <textarea value={mismatchNote} onChange={e => handleMismatchNoteChange(e.target.value)} onBlur={handleMismatchNoteBlur}
              placeholder="What's different in the field?" rows={2}
              style={{ width: "100%", marginTop: 8, padding: "8px 10px", border: "1.5px solid #FECACA", borderRadius: 8, fontSize: 13, fontFamily: "DM Sans, sans-serif", color: "#111827", resize: "none", outline: "none", boxSizing: "border-box", background: "#FFF" }}/>
          )}
        </div>
      )}

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
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: (h.note || h.mismatch) ? 6 : 0 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <StatusBadge status={h.status}/>
                      {h.mismatch && <span style={{ fontSize: 10, fontWeight: 600, color: "#991B1B", background: "#FEF2F2", padding: "1px 6px", borderRadius: 20 }}>⚡ flagged</span>}
                    </div>
                    <span style={{ fontSize: 11, color: "#9CA3AF" }}>{fmtDate(h.updatedAt)}</span>
                  </div>
                  {h.note && <p style={{ margin: 0, fontSize: 12, color: "#4B5563", lineHeight: 1.5 }}>{h.note}</p>}
                  {h.mismatch && h.mismatchNote && <p style={{ margin: "4px 0 0", fontSize: 12, color: "#991B1B", lineHeight: 1.5 }}>⚡ {h.mismatchNote}</p>}
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
    else if (screen === "edit") setScreen("dashboard");
  };

  const titles = { projects: "Field Documentation Tracker", create: "New project", edit: "Edit project", dashboard: activeProject?.name||"", checklist: activeCategory?.id||"", item: "Document item" };

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
      {screen === "create" && <ProjectForm onSave={proj=>{setData(d=>({...d,projects:[...d.projects,proj],records:{...d.records,...applyEarthCraftAutoPass(proj,d.records),...applyEarthCraftGoldOverlapAutoPass(proj,d.records)}}));setScreen("projects");}} onBack={navBack} auth={auth} setAuth={setAuth}/>}
      {screen === "edit" && activeProject && (
        <ProjectForm
          initialProject={activeProject}
          onSave={proj=>{setData(d=>({...d,projects:d.projects.map(p=>p.id===proj.id?proj:p),records:{...d.records,...applyEarthCraftAutoPass(proj,d.records),...applyEarthCraftGoldOverlapAutoPass(proj,d.records)}}));setActiveProject(proj);setScreen("dashboard");}}
          onBack={navBack}
          auth={auth}
          setAuth={setAuth}
        />
      )}
      {screen === "dashboard" && activeProject && (
        <ProjectDashboard
          project={activeProject}
          records={data.records}
          onSelectCategory={cat=>{setActiveCategory(cat);setScreen("checklist");}}
          onSelectItem={item=>{setActiveItem(item);setScreen("item");}}
          onEdit={()=>setScreen("edit")}
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