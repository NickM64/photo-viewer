"use strict";
/* Photo Viewer - view only (Photo Manager spec section 2.11).
   Stage 2: sign-in that stays signed in, the Viewer Catalogue File, Albums, Grid and Photo screens.
   Stage 3: the Filter (2.11.5) and the caption / place / date switches in Settings.

   The page talks only to Microsoft: the sign-in service and OneDrive (Microsoft Graph), with the single
   permission "read files".  It has no server of its own. */

const BUILD = "0.3.1";                          // also in version.json and in index.html's ?v= (see CLAUDE.md)
const VERSION = BUILD + " (stage 3)";
const CFG = Object.assign({ clientId: "", catalogueFolder: "PhotoManager/Shared", catalogueFile: "viewer_catalogue.json.gz" },
  window.PV_CONFIG || {});
const LOGIN = "https://login.microsoftonline.com/consumers/oauth2/v2.0";   // personal Microsoft accounts
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPE = "Files.Read offline_access";      // read-only files; offline_access lets the sign-in renew itself
const REDIRECT = location.origin + location.pathname.replace(/index\.html$/i, "");
const THUMB = "c300x300_crop";                  // the one size of small picture asked for (trial: fastest)
const CHUNK = 240;                              // grid cells added at a time
const LIST_ALBUM_FROM = 5;                      // this many pictures wanted from one album: list the album instead

const $ = (id) => document.getElementById(id);
const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };

// ---------------------------------------------------------------- storage (never relied on: all wrapped)
const store = {
  get(k, area = localStorage) { try { return area.getItem(k); } catch (e) { return null; } },
  set(k, v, area = localStorage) { try { area.setItem(k, v); } catch (e) { /* storage off */ } },
  del(k, area = localStorage) { try { area.removeItem(k); } catch (e) { /* storage off */ } },
};
const db = {                                    // IndexedDB: holds the copy of the Viewer Catalogue File
  open() {
    return new Promise((ok, fail) => {
      const r = indexedDB.open("photo-viewer", 1);
      r.onupgradeneeded = () => r.result.createObjectStore("kv");
      r.onsuccess = () => ok(r.result);
      r.onerror = () => fail(r.error);
    });
  },
  async run(mode, work) {
    try {
      const d = await this.open();
      return await new Promise((ok, fail) => {
        const tx = d.transaction("kv", mode);
        const req = work(tx.objectStore("kv"));
        tx.oncomplete = () => { d.close(); ok(req ? req.result : undefined); };
        tx.onerror = tx.onabort = () => { d.close(); fail(tx.error); };
      });
    } catch (e) { return undefined; }           // private window, storage cleared or blocked: carry on without
  },
  get(k) { return this.run("readonly", (s) => s.get(k)); },
  set(k, v) { return this.run("readwrite", (s) => s.put(v, k)); },
  clear() { return this.run("readwrite", (s) => s.clear()); },
};

// ---------------------------------------------------------------- notices
let noticeTimer = 0;
function notice(text, stay, action) {
  $("noticeText").textContent = text;
  $("noticeAct").hidden = !action;
  $("noticeAct").onclick = action ? action.run : null;
  if (action) $("noticeAct").textContent = action.label;
  $("noticeClose").textContent = action ? "Later" : "OK";
  $("notice").hidden = false;
  clearTimeout(noticeTimer);
  if (!stay) noticeTimer = setTimeout(() => { $("notice").hidden = true; }, 6000);
}

// ---------------------------------------------------------------- a new version of the viewer
/* A phone keeps the page's files for a while, and a home-screen icon may keep them even after the app is
   swiped away.  version.json is asked for past every cache; when it names another version the viewer offers
   to update.  Updating opens the page at a new address (?v=...), which no cache holds, and index.html asks
   for app.js and config.js with the same ?v=, so all three files come fresh. */
const update = {
  checkedAt: 0, ready: "",
  async check() {
    this.checkedAt = Date.now();
    try {
      const reply = await fetch("version.json?t=" + Date.now(), { cache: "no-store" });
      const v = reply.ok ? String((await reply.json()).version || "") : "";
      if (!v || v === BUILD) return "";
      this.ready = v;
      notice("A new version of the viewer is ready (" + v + ").", true, { label: "Update", run: () => this.now() });
      return v;
    } catch (e) { return ""; }               // no connection: try again next time
  },
  now() { location.replace(REDIRECT + "?v=" + encodeURIComponent(this.ready || BUILD)); },
};
document.addEventListener("visibilitychange", () => {   // a home-screen app is often brought back rather than reopened
  if (document.visibilityState === "visible" && Date.now() - update.checkedAt > 5 * 60 * 1000) update.check();
});

// ---------------------------------------------------------------- sign-in (authorization code with PKCE)
class SignInNeeded extends Error {}
const clientId = () => (CFG.clientId || store.get("pv_client") || "").trim();

function base64url(bytes) {
  let s = "";
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function randomText(n) { const a = new Uint8Array(n); crypto.getRandomValues(a); return base64url(a); }

/* Go to Microsoft's sign-in page.  quiet = true asks it to come straight back without showing anything,
   which works while Microsoft still remembers the account on this device. */
async function goToSignIn(quiet) {
  const verifier = randomText(48), state = randomText(16);
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
  store.set("pv_verifier", verifier, sessionStorage);
  store.set("pv_state", state, sessionStorage);
  store.set("pv_quiet", quiet ? "1" : "", sessionStorage);
  const q = new URLSearchParams({
    client_id: clientId(), response_type: "code", redirect_uri: REDIRECT, response_mode: "fragment",
    scope: SCOPE, state, code_challenge: base64url(new Uint8Array(digest)), code_challenge_method: "S256",
    prompt: quiet ? "none" : "select_account",
  });
  location.assign(LOGIN + "/authorize?" + q);
  return new Promise(() => {});                 // the page is leaving
}

async function tokenRequest(fields) {
  let reply;
  try {
    reply = await fetch(LOGIN + "/token", { method: "POST", body: new URLSearchParams(Object.assign({ client_id: clientId(), scope: SCOPE }, fields)) });
  } catch (e) {
    throw new Error("Could not reach Microsoft to sign in. Check the connection and try again.");
  }
  const data = await reply.json().catch(() => ({}));
  if (!reply.ok) {
    const err = new Error("Microsoft sign-in said: " + (data.error || reply.status) + "\n" + (data.error_description || ""));
    err.code = data.error;
    throw err;
  }
  store.set("pv_token", data.access_token);
  store.set("pv_expires", String(Date.now() + (data.expires_in - 120) * 1000));
  if (data.refresh_token) store.set("pv_refresh", data.refresh_token);
  store.set("pv_signed_in", "1");
  return data.access_token;
}

/* Called at start-up when Microsoft has sent the browser back here.  Returns "ok", "quiet-failed" or throws. */
async function finishSignIn(params) {
  history.replaceState(history.state, "", REDIRECT);
  const quiet = store.get("pv_quiet", sessionStorage) === "1";
  if (params.get("error")) {
    if (quiet) return "quiet-failed";            // Microsoft wants the account chosen again: show the button
    throw new Error("Microsoft sign-in said: " + params.get("error") + "\n" + (params.get("error_description") || ""));
  }
  if (!params.get("state") || params.get("state") !== store.get("pv_state", sessionStorage)) {
    throw new Error("The sign-in reply did not match this page's request. Please try again.");
  }
  await tokenRequest({ grant_type: "authorization_code", code: params.get("code"), redirect_uri: REDIRECT,
    code_verifier: store.get("pv_verifier", sessionStorage) || "" });
  store.del("pv_verifier", sessionStorage);
  store.del("pv_state", sessionStorage);
  return "ok";
}

let renewing = null;
async function accessToken(forceRenew) {
  const t = store.get("pv_token");
  if (!forceRenew && t && Date.now() < Number(store.get("pv_expires") || 0)) return t;
  const refresh = store.get("pv_refresh");
  if (!refresh) throw new SignInNeeded();
  if (!renewing) {
    renewing = tokenRequest({ grant_type: "refresh_token", refresh_token: refresh })
      .catch((e) => {
        if (e.code) { store.del("pv_refresh"); store.del("pv_token"); throw new SignInNeeded(); }   // Microsoft refused: sign in again
        throw e;                                                                                    // no connection: keep the sign-in
      })
      .finally(() => { renewing = null; });
  }
  return renewing;
}

function forgetSignIn() {
  for (const k of ["pv_token", "pv_expires", "pv_refresh", "pv_signed_in", "pv_addrtest", "pv_filter", "pv_show"]) store.del(k);
  FILTER = readFilter();
  SHOW = readShow();
}

// ---------------------------------------------------------------- OneDrive (Microsoft Graph)
const encPath = (p) => p.split("/").filter(Boolean).map(encodeURIComponent).join("/");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/* GET a Graph address (path after /v1.0, or a full address such as a next-page link) and return its JSON.
   Renews the sign-in once on 401; waits and retries when OneDrive says it is busy. */
async function graph(url) {
  const full = url.startsWith("https://") ? url : GRAPH + url;
  for (let attempt = 0; ; attempt++) {
    const token = await accessToken(attempt === 1 && graph.lastStatus === 401);
    let reply;
    try {
      reply = await fetch(full, { headers: { Authorization: "Bearer " + token } });
    } catch (e) {
      throw new Error("No connection to OneDrive.");
    }
    graph.lastStatus = reply.status;
    if (reply.status === 401 && attempt === 0) continue;
    if ((reply.status === 429 || reply.status === 503) && attempt < 3) {
      await sleep(Math.min(20, Number(reply.headers.get("Retry-After")) || 2 * (attempt + 1)) * 1000);
      continue;
    }
    if (reply.status === 401) throw new SignInNeeded();
    const data = await reply.json().catch(() => ({}));
    if (!reply.ok) {
      const err = new Error("OneDrive said: " + (data.error ? data.error.code + " - " + data.error.message : reply.status));
      err.status = reply.status;
      throw err;
    }
    return data;
  }
}

// ---------------------------------------------------------------- the Viewer Catalogue File
let CAT = null;        // { data, col, lists: Map }  - the catalogue in use
let catInfo = null;    // { etag, modified, fetchedAt }

async function gunzipText(buffer) {
  const b = new Uint8Array(buffer);
  if (b[0] !== 0x1f || b[1] !== 0x8b) return new TextDecoder().decode(b);     // already unpacked on the way
  if (!("DecompressionStream" in window)) throw new Error("This browser is too old to open the photo list. Please update it.");
  return new Response(new Blob([buffer]).stream().pipeThrough(new DecompressionStream("gzip"))).text();
}

function useCatalogue(data, info) {
  if (!data || data.format !== 1 || !Array.isArray(data.photos) || !Array.isArray(data.columns)) {
    throw new Error("The photo list in OneDrive is not in a form this version of the viewer understands (format " +
      (data && data.format) + "). The viewer may need updating.");
  }
  const col = {};
  data.columns.forEach((name, i) => { col[name] = i; });
  for (const need of ["album", "file", "video", "date", "date_text", "date_sort", "caption", "site", "city", "state", "country", "rating"]) {
    if (!(need in col)) throw new Error("The photo list in OneDrive is missing the column '" + need + "'.");
  }
  const counts = data.albums.map(() => 0);
  for (const p of data.photos) counts[p[col.album]]++;
  data.albums.forEach((a, i) => { a.count = counts[i]; a.name = a.path.split("/").pop() || a.title || a.path; });
  CAT = { data, col, lists: new Map() };
  catInfo = info;
  thumbs.reset();
  fullAddress.clear();
}

/* The photos of one album (a number) or of the whole collection ("all"), oldest first, undated last;
   or the photos the Filter matches ("filter"), in the order the Filter asks for. */
function listFor(key) {
  if (CAT.lists.has(key)) return CAT.lists.get(key);
  const { data, col } = CAT;
  let out = [];
  if (key === "filter") out = filterMatches(FILTER);
  else data.photos.forEach((p, i) => { if (key === "all" || p[col.album] === key) out.push(i); });
  const order = key === "filter" ? FILTER.order : "old";
  const when = (i) => data.photos[i][col.date] || data.photos[i][col.date_sort] || "";
  const byName = (a, b) => { const f = data.photos[a][col.file], g = data.photos[b][col.file]; return f < g ? -1 : f > g ? 1 : 0; };
  if (order === "name") out.sort(byName);
  else {
    const way = order === "new" ? -1 : 1;
    out.sort((a, b) => {
      const x = when(a), y = when(b);
      if (x !== y) return !x ? 1 : !y ? -1 : x < y ? -way : way;      // no date: last, whichever way
      return byName(a, b);
    });
  }
  CAT.lists.set(key, out);
  return out;
}

// ---------------------------------------------------------------- the Filter (spec 2.11.5)
/* The choices are kept on the device.  Matching needs no request: it only reads the catalogue. */
const NO_FILTER = { text: "", album: "", site: "", location: "", from: "", to: "", rating: 0, type: "", order: "old" };
function readFilter() {
  let kept = null;
  try { kept = JSON.parse(store.get("pv_filter") || "null"); } catch (e) { kept = null; }
  return Object.assign({}, NO_FILTER, kept && typeof kept === "object" ? kept : {});
}
let FILTER = readFilter();
function saveFilter(f) {
  FILTER = f;
  store.set("pv_filter", JSON.stringify(f));
  if (CAT) CAT.lists.delete("filter");
  if (grid.key === "filter") grid.key = undefined;            // the grid is drawn again when next shown
}

/* Capitals, accents and spaces at either end are ignored, as in the Captioning App. */
const plain = (s) => String(s == null ? "" : s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
const yearOf = (s) => { const m = /^\s*(\d{4})\s*$/.exec(s || ""); return m ? Number(m[1]) : 0; };

/* One line of text per photo for Search, made the first time Search is used. */
function searchText() {
  if (CAT.hay) return CAT.hay;
  const { data, col } = CAT;
  const albumText = data.albums.map((a) => a.path + "\n" + (a.title || ""));
  const extra = ["location_long", "keywords", "people"].filter((n) => n in col).map((n) => col[n]);
  CAT.hay = data.photos.map((p) => {
    let s = p[col.caption] + "\n" + p[col.site] + "\n" + p[col.city] + "\n" + p[col.state] + "\n" + p[col.country] + "\n" + albumText[p[col.album]];
    for (const c of extra) s += "\n" + (p[c] || "");
    return plain(s);
  });
  return CAT.hay;
}

/* The photos that match every choice that is filled in (not yet in order). */
function filterMatches(f) {
  const { data, col } = CAT;
  const words = plain(f.text).split(/\s+/).filter(Boolean);
  const hay = words.length ? searchText() : null;
  const album = f.album ? data.albums.findIndex((a) => a.path === f.album) : -1;
  const site = plain(f.site), place = plain(f.location), rating = Number(f.rating) || 0;
  let from = yearOf(f.from), to = yearOf(f.to);
  if (from && to && from > to) [from, to] = [to, from];
  const out = [];
  data.photos.forEach((p, i) => {
    if (f.album && p[col.album] !== album) return;
    if (f.type === "photo" && p[col.video]) return;
    if (f.type === "video" && !p[col.video]) return;
    if (rating && !(p[col.rating] >= rating)) return;
    if (from || to) {
      const y = Number(String(p[col.date] || p[col.date_sort] || "").slice(0, 4));
      if (!y || (from && y < from) || (to && y > to)) return;
    }
    if (site && !plain(p[col.site]).includes(site)) return;
    if (place && !plain(p[col.city]).includes(place) && !plain(p[col.state]).includes(place) && !plain(p[col.country]).includes(place)) return;
    for (const w of words) if (!hay[i].includes(w)) return;
    out.push(i);
  });
  return out;
}

/* The active choices on one line; "" when nothing is chosen and the order is the usual one. */
function filterSummary(f) {
  const parts = [];
  if (f.text.trim()) parts.push("\u201c" + f.text.trim() + "\u201d");
  if (f.album) parts.push("Album: " + f.album.split("/").pop());
  if (f.site.trim()) parts.push("Site: " + f.site.trim());
  if (f.location.trim()) parts.push("Location: " + f.location.trim());
  let from = yearOf(f.from), to = yearOf(f.to);
  if (from && to && from > to) [from, to] = [to, from];
  if (from && to) parts.push(from === to ? String(from) : from + "\u2013" + to);
  else if (from) parts.push("From " + from);
  else if (to) parts.push("Up to " + to);
  const r = Number(f.rating) || 0;
  if (r) parts.push(r === 5 ? "5 stars" : r + (r === 1 ? " star" : " stars") + " or more");
  if (f.type === "photo") parts.push("Photos only");
  if (f.type === "video") parts.push("Videos only");
  if (f.order === "new") parts.push("Newest first");
  if (f.order === "name") parts.push("By file name");
  return parts.join(" \u00b7 ");
}

const filterForm = {
  ready: null,                                                // the catalogue the lists were made for
  read() {
    return { text: $("fText").value, album: $("fAlbum").value, site: $("fSite").value, location: $("fPlace").value,
      from: $("fFrom").value, to: $("fTo").value, rating: Number($("fRating").value) || 0, type: $("fType").value, order: $("fOrder").value };
  },
  fill(f) {
    $("fText").value = f.text; $("fAlbum").value = f.album; $("fSite").value = f.site; $("fPlace").value = f.location;
    $("fFrom").value = f.from; $("fTo").value = f.to; $("fRating").value = String(f.rating || 0); $("fType").value = f.type; $("fOrder").value = f.order;
    if ($("fAlbum").selectedIndex < 0) $("fAlbum").selectedIndex = 0;      // an album no longer in the photo list
  },
  lists() {                                                   // album choices and the suggestions, once per catalogue
    if (this.ready === CAT) return;
    this.ready = CAT;
    const { data, col } = CAT;
    const sel = $("fAlbum");
    sel.replaceChildren(new Option("All albums", ""));
    for (const [y, list] of albumsByYear()) {
      const g = document.createElement("optgroup");
      g.label = y || "No year";
      for (const i of list) g.append(new Option(data.albums[i].name, data.albums[i].path));
      sel.append(g);
    }
    const sites = new Set(), places = new Set();
    let lo = 9999, hi = 0;
    for (const p of data.photos) {
      if (p[col.site]) sites.add(p[col.site]);
      for (const c of [col.city, col.state, col.country]) if (p[c]) places.add(p[c]);
      const y = Number(String(p[col.date] || p[col.date_sort] || "").slice(0, 4));
      if (y) { if (y < lo) lo = y; if (y > hi) hi = y; }
    }
    const options = (id, set) => $(id).replaceChildren(...[...set].sort((x, y) => x.localeCompare(y, undefined, { sensitivity: "base" })).map((v) => new Option(v, v)));
    options("siteNames", sites);
    options("placeNames", places);
    $("fFrom").placeholder = hi ? String(lo) : "";
    $("fTo").placeholder = hi ? String(hi) : "";
  },
  draw() {
    this.lists();
    this.fill(FILTER);
    if ($("fAlbum").value !== FILTER.album) saveFilter(this.read());
    this.count();
  },
  count() {
    const found = filterMatches(FILTER);
    $("filterCount").textContent = found.length ? countText(found) : "No photos match";
    $("filterShow").disabled = !found.length;
  },
  changed() { saveFilter(this.read()); this.count(); },
};
(function filterControls() {
  let timer = 0;
  $("filterForm").addEventListener("input", () => { clearTimeout(timer); timer = setTimeout(() => filterForm.changed(), 120); });
  $("filterForm").addEventListener("change", () => { clearTimeout(timer); filterForm.changed(); });
  const show = () => {
    clearTimeout(timer);
    filterForm.changed();
    if ($("filterShow").disabled) return;
    if (document.activeElement && document.activeElement.blur) document.activeElement.blur();   // puts the phone's keyboard away
    if (shown.fromGrid) history.back(); else go({ v: "grid", album: "filter", photo: null });
  };
  $("filterForm").addEventListener("submit", (e) => { e.preventDefault(); show(); });
  $("filterShow").addEventListener("click", show);
  $("filterClear").addEventListener("click", () => { clearTimeout(timer); filterForm.fill(NO_FILTER); filterForm.changed(); });
})();

/* Which lines the photo screen shows (Settings), kept on the device. */
function readShow() {
  let kept = null;
  try { kept = JSON.parse(store.get("pv_show") || "null"); } catch (e) { kept = null; }
  return Object.assign({ caption: true, place: true, date: true }, kept && typeof kept === "object" ? kept : {});
}
let SHOW = readShow();

const cataloguePath = () => "/me/drive/root:/" + encPath(CFG.catalogueFolder + "/" + CFG.catalogueFile);

/* Start-up: show the copy kept on this device at once, then ask OneDrive (one small request) whether there
   is a newer file and download it only if so.  Returns true when a catalogue is in use. */
async function loadCatalogue(onChange, force) {
  if (!CAT) {
    const kept = await db.get("catalogue");
    if (kept && kept.data) {
      try { useCatalogue(kept.data, kept.info); onChange(); } catch (e) { /* an old form: fetch again below */ }
    }
  }
  let meta;
  try {
    meta = await graph(cataloguePath() + "?select=id,eTag,size,lastModifiedDateTime,@microsoft.graph.downloadUrl");
  } catch (e) {
    if (e.status === 404) {
      throw new Error("The photo list was not found in OneDrive at " + CFG.catalogueFolder + "/" + CFG.catalogueFile +
        ".\nIn the Captioning App, choose File > Export for Phone Viewer, wait for OneDrive to upload it, then try again.");
    }
    if (CAT && !(e instanceof SignInNeeded)) { notice("Could not check for new information: " + e.message); return "kept"; }
    throw e;
  }
  if (CAT && !force && catInfo && catInfo.etag === meta.eTag) return "same";
  const reply = await fetch(meta["@microsoft.graph.downloadUrl"]);
  if (!reply.ok) throw new Error("The photo list could not be downloaded from OneDrive (" + reply.status + ").");
  const data = JSON.parse(await gunzipText(await reply.arrayBuffer()));
  const info = { etag: meta.eTag, modified: meta.lastModifiedDateTime, size: meta.size, fetchedAt: new Date().toISOString() };
  const had = !!CAT;
  useCatalogue(data, info);
  db.set("catalogue", { data, info });
  onChange();
  return had ? "updated" : "new";
}

// ---------------------------------------------------------------- small pictures for the grid
/* Requests to OneDrive are the slow part (about half a second each), so:
   - several pictures wanted from one album  -> one request lists the album's folder with every small picture;
   - a few pictures from an album            -> each is looked up by its file name. */
const thumbs = {
  url: new Map(),          // photo index -> address, or null when OneDrive has no such file
  waiting: new Map(),      // photo index -> [img, ...]
  listed: new Map(),       // album index -> promise of the folder list
  queue: new Set(),
  timer: 0,
  active: 0,
  pending: [],
  reset() { this.url.clear(); this.waiting.clear(); this.listed.clear(); this.queue.clear(); this.pending = []; },

  folder(a) { return CAT.data.pictures_path + "/" + CAT.data.albums[a].path; },
  filePath(i) { const p = CAT.data.photos[i]; return this.folder(p[CAT.col.album]) + "/" + p[CAT.col.file]; },

  want(i, img) {
    if (this.url.has(i)) { this.show(i, img); return; }
    if (!this.waiting.has(i)) this.waiting.set(i, []);
    this.waiting.get(i).push(img);
    this.queue.add(i);
    clearTimeout(this.timer);
    this.timer = setTimeout(() => this.flush(), 60);
  },

  flush() {
    const byAlbum = new Map();
    for (const i of this.queue) {
      const a = CAT.data.photos[i][CAT.col.album];
      if (!byAlbum.has(a)) byAlbum.set(a, []);
      byAlbum.get(a).push(i);
    }
    this.queue.clear();
    for (const [a, wanted] of byAlbum) {
      if (this.listed.has(a) || wanted.length >= LIST_ALBUM_FROM) this.listAlbum(a);
      else for (const i of wanted) this.byName(i);
    }
  },

  listAlbum(a) {
    if (this.listed.has(a)) return this.listed.get(a);
    const cat = CAT;
    const work = (async () => {
      const found = new Map();
      let url = "/me/drive/root:/" + encPath(this.folder(a)) + ":/children?$select=id,name&$expand=thumbnails($select=" + THUMB + ")&$top=1000";
      while (url) {
        const page = await graph(url);
        for (const item of page.value || []) {
          const t = item.thumbnails && item.thumbnails[0] && item.thumbnails[0][THUMB];
          if (t && t.url) found.set(item.name.toLowerCase(), t.url);
        }
        url = page["@odata.nextLink"];
      }
      if (cat !== CAT) return;
      cat.data.photos.forEach((p, i) => {
        if (p[cat.col.album] !== a) return;
        this.settle(i, found.get(p[cat.col.file].toLowerCase()) || null);
      });
      addressTest.offer(found.values().next().value);
    })();
    this.listed.set(a, work);
    work.catch((e) => {
      if (cat !== CAT) return;
      this.listed.delete(a);                                  // let a later scroll try again
      cat.data.photos.forEach((p, i) => { if (p[cat.col.album] === a) this.fail(i); });
      onLoadError(e);
    });
    return work;
  },

  byName(i, replace) {
    const cat = CAT;
    this.pool(async () => {
      try {
        const item = await graph("/me/drive/root:/" + encPath(this.filePath(i)) + "?$select=id&$expand=thumbnails($select=" + THUMB + ")");
        if (cat !== CAT) return;
        const t = item.thumbnails && item.thumbnails[0] && item.thumbnails[0][THUMB];
        if (replace) this.url.delete(i);
        this.settle(i, (t && t.url) || null);
      } catch (e) {
        if (cat !== CAT) return;
        if (e.status === 404) this.settle(i, null); else { this.fail(i); onLoadError(e); }
      }
    });
  },

  settle(i, url) {
    this.url.set(i, url);
    for (const img of this.waiting.get(i) || []) this.show(i, img);
    this.waiting.delete(i);
  },
  fail(i) {                                                   // leaves the cell grey and asks again in a while
    const imgs = this.waiting.get(i) || [];
    this.waiting.delete(i);
    setTimeout(() => {
      for (const img of imgs) if (img.isConnected && grid.watcher) grid.watcher.observe(img.parentElement);
    }, 8000);
  },

  show(i, img) {
    const url = this.url.get(i);
    if (!url) { img.parentElement && img.parentElement.classList.add("missing"); return; }
    img.onload = () => img.classList.add("in");
    img.onerror = () => {                                     // an address that has stopped working: ask once more
      img.onerror = null;
      this.url.delete(i);
      this.waiting.set(i, [img]);
      this.byName(i, true);
    };
    img.src = url;
  },
};
thumbs.pool = function (job) {                                // at most 8 look-ups by name at a time
  if (job) this.pending.push(job);
  while (this.active < 8 && this.pending.length) {
    this.active++;
    this.pending.shift()().finally(() => { this.active--; this.pool(); });
  }
};

/* The address of a Large file, found by its file name.  OneDrive's download addresses are short-lived,
   so one is reused for 40 minutes at most. */
const fullAddress = {
  kept: new Map(),
  clear() { this.kept.clear(); },
  get(i) {
    const k = this.kept.get(i);
    if (k && Date.now() - k.at < 40 * 60 * 1000) return k.promise;
    const cat = CAT;
    const promise = graph("/me/drive/root:/" + encPath(thumbs.filePath(i)) + "?select=id,size,@microsoft.graph.downloadUrl")
      .then((item) => item["@microsoft.graph.downloadUrl"]);
    this.kept.set(i, { at: Date.now(), promise });
    promise.catch(() => { if (cat === CAT) this.kept.delete(i); });
    return promise;
  },
};

/* Finding out how long OneDrive keeps a small picture's address valid (spec 2.11.6): one address is kept
   with the time it was given, and Settings tries it again each time it is opened. */
const addressTest = {
  offer(url) {
    if (!url || store.get("pv_addrtest")) return;
    store.set("pv_addrtest", JSON.stringify({ url, at: Date.now(), lastOk: Date.now(), failedAt: 0 }));
  },
  async check() {
    let t;
    try { t = JSON.parse(store.get("pv_addrtest") || "null"); } catch (e) { t = null; }
    if (!t) return "No picture address saved yet. Open an album first.";
    const hours = (ms) => { const h = ms / 3600000; return h < 1 ? Math.round(h * 60) + " min" : h.toFixed(1) + " h"; };
    if (!t.failedAt) {
      const ok = await new Promise((done) => { const i = new Image(); i.onload = () => done(true); i.onerror = () => done(false); i.src = t.url; });
      if (ok) t.lastOk = Date.now(); else t.failedAt = Date.now();
      store.set("pv_addrtest", JSON.stringify(t));
    }
    return t.failedAt
      ? "Worked for " + hours(t.lastOk - t.at) + "; had stopped by " + hours(t.failedAt - t.at)
      : "Still working after " + hours(t.lastOk - t.at);
  },
};

// ---------------------------------------------------------------- screens and moving between them
const LAYERS = ["signin", "albums", "grid", "filter", "settings"];
let shown = { v: "", album: null, photo: null };

function showLayer(name) {
  for (const id of LAYERS) {
    const on = id === name;
    $(id).classList.toggle("off", !on);
    $(id).inert = !on;
  }
}

/* state = { v: "albums" | "grid" | "filter" | "settings", album: number | "all" | "filter", photo: position in the list or null,
             fromGrid: true when the Filter screen was opened from the grid of Filter results } */
function go(state, replace) {
  if (replace) history.replaceState(state, ""); else history.pushState(state, "");
  render(state);
}
function render(state) {
  if (!CAT) return;
  state = state || { v: "albums" };
  if (state.v === "grid" && (state.album === "all" || state.album === "filter" || CAT.data.albums[state.album])) {
    if (grid.key !== state.album) grid.build(state.album);
    showLayer("grid");
    if (state.photo != null && state.photo < grid.list.length) photo.open(state.photo); else photo.close();
  } else if (state.v === "settings") {
    photo.close();
    showLayer("settings");
    drawSettings();
  } else if (state.v === "filter") {
    photo.close();
    showLayer("filter");
    filterForm.draw();
  } else {
    photo.close();
    drawFilterRow();
    showLayer("albums");
  }
  shown = state;
}
window.addEventListener("popstate", (e) => render(e.state));

function countText(indices) {
  let videos = 0;
  for (const i of indices) videos += CAT.data.photos[i][CAT.col.video] ? 1 : 0;
  const n = indices.length - videos;
  const part = (k, word) => k.toLocaleString() + " " + word + (k === 1 ? "" : "s");
  return !videos ? part(n, "photo") : n ? part(n, "photo") + ", " + part(videos, "video") : part(videos, "video");
}

// ---- Albums
function albumsByYear() {                         // [[year text, [album numbers]], ...] newest year first, no year last
  const { albums } = CAT.data;
  const years = new Map();
  albums.forEach((a, i) => {
    if (!a.count) return;
    const y = a.year == null ? "" : String(a.year);
    if (!years.has(y)) years.set(y, []);
    years.get(y).push(i);
  });
  const order = [...years.keys()].sort((x, y) => (x === "" ? 1 : y === "" ? -1 : Number(y) - Number(x)));
  return order.map((y) => [y, years.get(y).sort((i, j) => albums[i].name.localeCompare(albums[j].name, undefined, { numeric: true, sensitivity: "base" }))]);
}

/* The line under "All photos" that reopens the Filter's photos; only there while a choice is made. */
function drawFilterRow() {
  const b = $("filterRow");
  if (!b) return;
  const summary = filterSummary(FILTER);
  b.hidden = !summary;
  if (!summary) return;
  const text = el("span", "", "Filter");
  text.append(el("span", "what", summary));
  b.replaceChildren(text, el("span", "n", listFor("filter").length.toLocaleString()));
}

function drawAlbums() {
  const box = $("albumList");
  box.replaceChildren();
  const { albums, photos } = CAT.data;
  const all = el("div", "rows");
  const b = el("button");
  b.append(el("span", "", "All photos"), el("span", "n", photos.length.toLocaleString()));
  b.addEventListener("click", () => go({ v: "grid", album: "all", photo: null }));
  const f = el("button");
  f.id = "filterRow";
  f.hidden = true;
  f.addEventListener("click", () => go({ v: "grid", album: "filter", photo: null }));
  all.append(b, f);
  box.append(all);
  drawFilterRow();

  for (const [y, list] of albumsByYear()) {
    box.append(el("p", "year", y || "NO YEAR"));
    const rows = el("div", "rows");
    for (const i of list) {
      const r = el("button");
      r.append(el("span", "", albums[i].name), el("span", "n", albums[i].count.toLocaleString()));
      r.addEventListener("click", () => go({ v: "grid", album: i, photo: null }));
      rows.append(r);
    }
    box.append(rows);
  }
  if (!photos.length) box.append(el("p", "muted", "The photo list is empty."));
}

// ---- Grid
const grid = {
  key: undefined, list: [], drawn: 0, watcher: null, more: null,

  build(key) {
    this.key = key;
    this.list = listFor(key);
    this.drawn = 0;
    $("cells").replaceChildren();
    $("grid").scrollTop = 0;
    const title = $("gridTitle");
    const filtered = key === "filter";
    title.replaceChildren(document.createTextNode(key === "all" ? "All photos" : filtered ? "Filter" : CAT.data.albums[key].name),
      el("span", "sub", countText(this.list)));
    $("gridBack").textContent = filtered ? "\u2039 Back" : "\u2039 Albums";
    $("gridFilter").hidden = !filtered;
    $("gridSummary").hidden = !filtered;
    $("gridSummary").textContent = filtered ? filterSummary(FILTER) || "No choices made: all photos" : "";
    $("gridNote").hidden = this.list.length > 0;
    $("gridNote").textContent = filtered ? "No photos match the Filter." : "No photos here.";
    if (this.watcher) this.watcher.disconnect();
    if (this.more) this.more.disconnect();
    // a picture is asked for when its cell comes within two screens of view
    this.watcher = new IntersectionObserver((entries) => {
      for (const e of entries) {
        if (!e.isIntersecting) continue;
        this.watcher.unobserve(e.target);
        thumbs.want(Number(e.target.dataset.i), e.target.firstChild);
      }
    }, { root: $("grid"), rootMargin: "150% 0px" });
    this.more = new IntersectionObserver((entries) => { if (entries.some((e) => e.isIntersecting)) this.addChunk(); },
      { root: $("grid"), rootMargin: "300% 0px" });
    this.more.observe($("gridEnd"));
    this.addChunk();
    // an album, or a Filter that names one: one request for the lot.  Otherwise thumbs.want decides as cells come into view.
    const one = filtered ? (FILTER.album ? CAT.data.albums.findIndex((a) => a.path === FILTER.album) : -1) : key === "all" ? -1 : key;
    if (one >= 0 && this.list.length >= LIST_ALBUM_FROM) thumbs.listAlbum(one);
    else if (one >= 0 && !filtered && this.list.length) thumbs.listAlbum(one);
  },

  addChunk(upTo) {
    const { photos } = CAT.data, col = CAT.col;
    const end = Math.min(this.list.length, Math.max(this.drawn + CHUNK, upTo || 0));
    const frag = document.createDocumentFragment();
    for (let pos = this.drawn; pos < end; pos++) {
      const i = this.list[pos], p = photos[i];
      const cell = el("button", "cell");
      cell.dataset.i = i;
      cell.dataset.pos = pos;
      const img = el("img");
      img.alt = p[col.caption] || p[col.file];
      img.decoding = "async";
      cell.append(img);
      if (p[col.video]) cell.append(el("span", "tag", "▶ " + lengthText(p[col.length])));
      frag.append(cell);
      this.watcher.observe(cell);
    }
    $("cells").append(frag);
    this.drawn = end;
  },

  reveal(pos) {                                   // coming back from a photo: make sure its cell is in view
    if (pos >= this.drawn) this.addChunk(pos + 1);
    const cell = $("cells").children[pos];
    if (!cell) return;
    const box = $("grid").getBoundingClientRect(), r = cell.getBoundingClientRect();
    if (r.top < box.top + 60 || r.bottom > box.bottom) cell.scrollIntoView({ block: "center" });
  },
};
$("cells").addEventListener("click", (e) => {
  const cell = e.target.closest(".cell");
  if (cell) go({ v: "grid", album: grid.key, photo: Number(cell.dataset.pos) });
});

function lengthText(sec) {
  if (sec == null) return "";
  const m = Math.floor(sec / 60), s = String(sec % 60).padStart(2, "0");
  return m >= 60 ? Math.floor(m / 60) + ":" + String(m % 60).padStart(2, "0") + ":" + s : m + ":" + s;
}

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function dateText(p) {
  const d = p[CAT.col.date];
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(d || "");
  if (m) return Number(m[3]) + " " + MONTHS[Number(m[2]) - 1] + " " + m[1];
  return p[CAT.col.date_text] || "";
}

// ---- one photo
const photo = {
  pos: -1, isOpen: false, ticket: 0,

  open(pos) {
    this.isOpen = true;
    $("photo").classList.remove("off", "bare");
    $("photo").inert = false;
    $("grid").inert = true;
    this.show(pos);
  },
  close() {
    if (!this.isOpen) return;
    this.isOpen = false;
    this.ticket++;
    $("photo").classList.add("off");
    $("photo").inert = true;
    $("grid").inert = false;
    $("stage").replaceChildren();                 // stops a playing video and drops the picture
    if (this.pos >= 0) grid.reveal(this.pos);
  },

  show(pos) {
    this.pos = pos;
    const ticket = ++this.ticket;
    const i = grid.list[pos], p = CAT.data.photos[i], col = CAT.col;
    const isVideo = !!p[col.video];
    $("photo").classList.toggle("video", isVideo);
    $("photoCount").textContent = (pos + 1).toLocaleString() + " of " + grid.list.length.toLocaleString();
    $("prevBtn").hidden = pos === 0;
    $("nextBtn").hidden = pos === grid.list.length - 1;

    const info = $("photoInfo");
    info.replaceChildren();
    if (SHOW.caption && p[col.caption]) info.append(el("p", "cap", p[col.caption]));
    const place = [p[col.site], p[col.city], p[col.state], p[col.country]].filter(Boolean).join(", ");
    if (SHOW.place && place) info.append(el("p", "meta", place));
    const date = dateText(p);
    if (SHOW.date && date) info.append(el("p", "meta", date));
    if (p[col.rating] > 0) { const s = el("p", "stars", "★".repeat(p[col.rating])); s.setAttribute("aria-label", p[col.rating] + " stars"); info.append(s); }

    const stage = $("stage");
    stage.replaceChildren();
    const small = thumbs.url.get(i);
    if (small && !isVideo) {                      // the small picture straight away, the Large file over it when it arrives
      const blur = el("img", "blur");
      blur.alt = "";
      blur.src = small;
      stage.append(blur);
    } else if (!isVideo) {
      stage.append(el("p", "msg", "Loading…"));
    }
    fullAddress.get(i).then((url) => {
      if (ticket !== this.ticket) return;
      if (isVideo) {
        const v = el("video");
        v.controls = true; v.playsInline = true; v.preload = "metadata";
        if (small) v.poster = small;
        v.src = url;
        stage.replaceChildren(v);
        return;
      }
      const img = el("img");
      img.alt = p[col.caption] || p[col.file];
      img.onload = () => { if (ticket === this.ticket) stage.replaceChildren(img); };
      img.onerror = () => { if (ticket === this.ticket) stage.replaceChildren(el("p", "msg", "This photo could not be shown.")); };
      img.src = url;
    }).catch((e) => {
      if (ticket !== this.ticket) return;
      stage.replaceChildren(el("p", "msg", e.status === 404 ? "This photo was not found in OneDrive." : "This photo could not be fetched.\n" + e.message));
      if (e instanceof SignInNeeded) onLoadError(e);
    });
    this.fetchAhead(pos + 1);
  },

  fetchAhead(pos) {                               // so that the next swipe is quick
    if (pos >= grid.list.length) return;
    const i = grid.list[pos];
    if (CAT.data.photos[i][CAT.col.video]) return;
    fullAddress.get(i).then((url) => { new Image().src = url; }).catch(() => {});
  },

  step(by) {
    const to = this.pos + by;
    if (!this.isOpen || to < 0 || to >= grid.list.length) return;
    go({ v: "grid", album: grid.key, photo: to }, true);     // replaces: Back still returns to the grid in one go
  },
};

(function photoControls() {
  const box = $("photo");
  $("photoBack").addEventListener("click", () => history.back());
  $("prevBtn").addEventListener("click", (e) => { e.stopPropagation(); photo.step(-1); });
  $("nextBtn").addEventListener("click", (e) => { e.stopPropagation(); photo.step(1); });
  // a tap on the picture hides or shows the lines; a sideways swipe moves to the next or previous photo
  let start = null, swiped = false;
  const zoomed = () => window.visualViewport && window.visualViewport.scale > 1.01;
  box.addEventListener("touchstart", (e) => { start = e.touches.length === 1 ? { x: e.touches[0].clientX, y: e.touches[0].clientY, t: Date.now() } : null; swiped = false; }, { passive: true });
  box.addEventListener("touchend", (e) => {
    if (!start || zoomed() || e.target.closest("video")) { start = null; return; }
    const t = e.changedTouches[0], dx = t.clientX - start.x, dy = t.clientY - start.y;
    start = null;
    if (Math.abs(dx) > 50 && Math.abs(dx) > 2 * Math.abs(dy)) { swiped = true; photo.step(dx < 0 ? 1 : -1); }
  }, { passive: true });
  box.addEventListener("click", (e) => {
    if (swiped) { swiped = false; return; }
    if (e.target.closest("button, video")) return;
    if (!box.classList.contains("video")) box.classList.toggle("bare");
  });
  document.addEventListener("keydown", (e) => {
    if (!photo.isOpen) return;
    if (e.key === "ArrowRight") photo.step(1);
    else if (e.key === "ArrowLeft") photo.step(-1);
    else if (e.key === "Escape") history.back();
  });
})();

// ---- Settings
function drawSettings() {
  $("showCaption").checked = SHOW.caption; $("showPlace").checked = SHOW.place; $("showDate").checked = SHOW.date;
  const row = (dl, name, value) => { const d = el("div"); d.append(el("dt", "", name), el("dd", "", value)); dl.append(d); return d.lastChild; };
  const when = (iso) => (iso ? new Date(iso).toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" }) : "-");
  const about = $("aboutList");
  about.replaceChildren();
  row(about, "Photos", CAT.data.photos.length.toLocaleString() + " in " + CAT.data.albums.filter((a) => a.count).length.toLocaleString() + " albums");
  row(about, "Written", when(CAT.data.written_at) + (CAT.data.machine ? " on " + CAT.data.machine : ""));
  row(about, "Fetched", when(catInfo && catInfo.fetchedAt));
  row(about, "From", CFG.catalogueFolder + "/" + CFG.catalogueFile);
  const dev = $("deviceList");
  dev.replaceChildren();
  row(dev, "Viewer version", VERSION + (update.ready ? " (" + update.ready + " is ready)" : ""));
  const test = row(dev, "Picture address test", "Checking…");
  addressTest.check().then((text) => { test.textContent = text; });
}
$("openSettings").addEventListener("click", () => go({ v: "settings" }));
$("openFilter").addEventListener("click", () => go({ v: "filter" }));
$("gridFilter").addEventListener("click", () => go({ v: "filter", fromGrid: true }));
for (const [id, key] of [["showCaption", "caption"], ["showPlace", "place"], ["showDate", "date"]]) {
  $(id).addEventListener("change", () => { SHOW[key] = $(id).checked; store.set("pv_show", JSON.stringify(SHOW)); });
}
for (const b of document.querySelectorAll("[data-back]")) b.addEventListener("click", () => history.back());
$("noticeClose").addEventListener("click", () => { $("notice").hidden = true; });
$("checkNow").addEventListener("click", async () => {
  const b = $("checkNow");
  b.disabled = true;
  b.textContent = "Checking…";
  try {
    const result = await loadCatalogue(() => { drawAlbums(); grid.key = undefined; }, false);
    if (result !== "same") render(history.state);
    if (!(await update.check())) notice(result === "same" ? "The photo list is up to date." : result === "kept" ? "Could not check just now." : "New photo list loaded.");
  } catch (e) { onLoadError(e); }
  b.disabled = false;
  b.textContent = "Check for new information now";
});
$("signOut").addEventListener("click", async () => {
  forgetSignIn();
  await db.clear();
  CAT = null;
  history.replaceState(null, "", REDIRECT);
  showSignIn("");
});

// ---------------------------------------------------------------- start-up
function showSignIn(message) {
  photo.close();
  showLayer("signin");
  $("clientBox").hidden = !!CFG.clientId;
  $("clientId").value = store.get("pv_client") || "";
  $("signinError").textContent = message || "";
}
$("signInBtn").addEventListener("click", () => {
  if (!CFG.clientId) {
    const id = $("clientId").value.trim();
    if (!/^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(id)) {
      $("signinError").textContent = "Enter the Application (client) ID first. It looks like 8-4-4-4-12 letters and digits.";
      return;
    }
    store.set("pv_client", id);
  }
  if (!window.crypto || !crypto.subtle) {
    $("signinError").textContent = "This page must be opened from an https address (or http://localhost), not as a file.";
    return;
  }
  goToSignIn(false).catch((e) => { $("signinError").textContent = e.message; });
});

/* Something could not be fetched.  If the sign-in has run out, renew it: quietly first, the button if not. */
let renewAsked = false;
function onLoadError(e) {
  if (e instanceof SignInNeeded) {
    if (renewAsked) return;
    renewAsked = true;
    if (store.get("pv_signed_in") && clientId() && !store.get("pv_quiet_tried", sessionStorage)) {
      store.set("pv_quiet_tried", "1", sessionStorage);
      goToSignIn(true);
    } else {
      showSignIn("Please sign in again.");
    }
    return;
  }
  notice(e.message);
}

async function start() {
  if (location.search) history.replaceState(null, "", REDIRECT + location.hash);   // drop the ?v= an update used
  update.check();
  $("photo").inert = true;
  const params = new URLSearchParams(location.hash.slice(1));
  let cameBack = "";
  if (params.get("code") || params.get("error")) {
    try { cameBack = await finishSignIn(params); } catch (e) { showSignIn(e.message); return; }
  }
  if (cameBack === "ok") store.del("pv_quiet_tried", sessionStorage);
  if (cameBack === "quiet-failed") { showSignIn("Please sign in again."); return; }
  if (!clientId() || (!store.get("pv_refresh") && !store.get("pv_token"))) {
    if (store.get("pv_signed_in") && clientId() && !store.get("pv_quiet_tried", sessionStorage) && window.crypto && crypto.subtle) {
      store.set("pv_quiet_tried", "1", sessionStorage);
      await goToSignIn(true);
    }
    showSignIn("");
    return;
  }
  showLayer("albums");
  try {
    await loadCatalogue(() => {
      drawAlbums();
      grid.key = undefined;
      render(history.state);
    });
  } catch (e) {
    if (e instanceof SignInNeeded) { onLoadError(e); return; }
    if (!CAT) {
      const box = $("albumList");
      box.replaceChildren(el("p", "", e.message));
      box.firstChild.style.cssText = "padding-top:24px;white-space:pre-wrap";
      const again = el("button", "btn wide", "Try again");
      again.addEventListener("click", () => location.reload());
      box.append(again);
    } else {
      notice(e.message);
    }
  }
}
start();
