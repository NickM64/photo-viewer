/* Checks for the Photo Viewer, run at iPhone size against a stand-in for Microsoft sign-in and OneDrive.
   Nothing here touches a real account: every request to Microsoft is answered by this script.

   Needs Node.js with Playwright (and its Chromium).   Run from the repository folder:
       node test/run.js [light|dark] [folder for screenshots]
   Prints PASS / FAIL lines and ends with a non-zero code if anything failed. */
const { chromium, devices } = require('playwright');
const zlib = require('zlib');
const http = require('http');
const fs = require('fs');
const path = require('path');
const ROOT = path.resolve(__dirname, '..');
const SCHEME = process.argv[2] === 'dark' ? 'dark' : 'light';
const OUT = process.argv[3] || '';
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.png': 'image/png' };
(async () => {
  // ---- the repository's own files, served as GitHub Pages serves them
  const server = http.createServer((req, res) => {
    const name = decodeURIComponent(req.url.split('?')[0]).replace(/^\/+/, '') || 'index.html';
    const file = path.join(ROOT, name);
    if (!file.startsWith(ROOT) || !fs.existsSync(file)) { res.writeHead(404); return res.end(); }
    res.writeHead(200, { 'content-type': TYPES[path.extname(file)] || 'application/octet-stream' });
    res.end(fs.readFileSync(file));
  });
  await new Promise((r) => server.listen(0, '127.0.0.1', r));
  const BASE = `http://localhost:${server.address().port}/`;
  let failed = 0;
  // ---- a made-up catalogue
  const columns = ["number","album","file","video","length","date","date_text","date_sort","caption","site","city","state","country","location_long","rating","lat","lon","keywords","people"];
  const albums = [{ path: "2000/2000 Kodak", title: "Kodak 2000", year: 2000 }, { path: "2019/2019 Cornwall & Devon", title: "", year: 2019 },
                  { path: "2019/2019 Christmas", title: "", year: 2019 }, { path: "Scans/Grandma's box", title: "", year: null }];
  const photos = []; let n = 0;
  const add = (a, file, o = {}) => photos.push([++n, a, file, o.video ? 1 : 0, o.length ?? null, o.date ?? "", o.date_text ?? "", o.date_sort ?? "", o.caption ?? "", o.site ?? "", o.city ?? "", o.state ?? "", o.country ?? "", "", o.rating ?? 0, null, null, "", ""]);
  for (let i = 1; i <= 620; i++) add(0, `2000-01-${String(i % 28 + 1).padStart(2, "0")}_Kodak_${String(i).padStart(4, "0")}.jpg`, { date: `2000-01-${String(i % 28 + 1).padStart(2, "0")}T10:00:00`, caption: i % 3 ? "" : "Photo " + i });
  add(1, "2019-07-14_Padstow_0A01.jpg", { date: "2019-07-14T09:30:00", caption: "Crabbing off the harbour wall before breakfast, with the tide coming in", site: "Harlyn Bay", city: "Padstow", state: "England", country: "United Kingdom", rating: 4 });
  add(1, "2019-07-14_Padstow_0A02.mp4", { video: 1, length: 83, date: "2019-07-14T12:00:00", caption: "Crab race", city: "Padstow", country: "United Kingdom" });
  add(1, "2019-Cornwall_0A03.jpg", { date_text: "Summer 2019", date_sort: "2019-07-01" });
  add(1, "2019-07-15_Padstow_0A04.jpg", { date: "2019-07-15T15:00:00" });
  add(1, "2019-07-16_Gone_0A05.jpg", { date: "2019-07-16T15:00:00", caption: "Not in OneDrive" });
  for (let i = 1; i <= 7; i++) add(2, `2019-12-25_Home_0B0${i}.jpg`, { date: "2019-12-25T1" + i + ":00:00", caption: "Christmas " + i, rating: i % 6 });
  add(3, "Scans_0C01.jpg", { date_text: "about 1962", caption: "Grandma at the seaside" });
  add(3, "Scans_0C02.jpg", {});
  const cat = { format: 1, written_at: "2026-11-14T08:55:00Z", machine: "ASPIRE", pictures_path: "PhotoManager/LargePictures", album_count: 4, photo_count: photos.length, albums, columns, photos };
  const gz = zlib.gzipSync(JSON.stringify(cat));
  let etag = '"v1"';

  const b = await chromium.launch();
  const ctx = await b.newContext({ ...devices['iPhone 13'], colorScheme: SCHEME });
  const p = await ctx.newPage();
  const log = [], seen = [];
  p.on('console', m => log.push('console:' + m.text()));
  p.on('pageerror', e => log.push('pageerror:' + e.message));
  let refreshOk = true, quietOk = true, tokenN = 0, tokenCalls = [];
  await ctx.route('https://login.microsoftonline.com/**', async r => {
    const u = new URL(r.request().url());
    if (u.pathname.endsWith('/authorize')) {
      seen.push('AUTHORIZE prompt=' + u.searchParams.get('prompt') + ' scope=' + u.searchParams.get('scope'));
      const back = u.searchParams.get('redirect_uri');
      if (u.searchParams.get('prompt') === 'none' && !quietOk) return r.fulfill({ status: 302, headers: { location: back + '#error=interaction_required&error_description=x&state=' + u.searchParams.get('state') } });
      return r.fulfill({ status: 302, headers: { location: back + '#code=CODE&state=' + u.searchParams.get('state') } });
    }
    const body = new URLSearchParams(r.request().postData());
    tokenCalls.push(body.get('grant_type'));
    const h = { 'access-control-allow-origin': '*' };
    if (body.get('grant_type') === 'refresh_token' && !refreshOk) return r.fulfill({ status: 400, headers: h, contentType: 'application/json', body: '{"error":"invalid_grant","error_description":"expired"}' });
    r.fulfill({ status: 200, headers: h, contentType: 'application/json', body: JSON.stringify({ access_token: 'TOK' + (++tokenN), refresh_token: 'REF' + tokenN, expires_in: 3600 }) });
  });
  const colour = s => { let h = 0; for (const c of s) h = (h * 31 + c.charCodeAt(0)) % 360; return h; };
  await ctx.route('https://img.test/**', r => {
    const u = decodeURIComponent(r.request().url()); const name = u.split('/').pop(); const full = u.includes('/full/');
    const [w, hh] = full ? [1600, 1200] : [300, 300];
    r.fulfill({ status: 200, contentType: 'image/svg+xml', headers: { 'access-control-allow-origin': '*' },
      body: `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${hh}"><rect width="100%" height="100%" fill="hsl(${colour(name)},45%,${full ? 42 : 55}%)"/><circle cx="${w * .7}" cy="${hh * .3}" r="${w * .1}" fill="hsl(${colour(name) + 40},70%,75%)"/><text x="20" y="${hh - 20}" font-size="${full ? 48 : 20}" fill="#fff" font-family="sans-serif">${name.replace(/&/g, '&amp;').slice(-14)}</text></svg>` });
  });
  await ctx.route('https://dl.test/**', r => r.fulfill({ status: 200, contentType: 'application/octet-stream', headers: { 'access-control-allow-origin': '*' }, body: gz }));
  let validTokens = () => true;
  await ctx.route('https://graph.microsoft.com/**', r => {
    const raw = r.request().url(); const u = decodeURIComponent(raw);
    const h = { 'access-control-allow-origin': '*' };
    if (r.request().method() === 'OPTIONS') return r.fulfill({ status: 204, headers: { ...h, 'access-control-allow-headers': 'authorization,content-type' } });
    const j = (o, s = 200) => r.fulfill({ status: s, headers: h, contentType: 'application/json', body: JSON.stringify(o) });
    const tok = (r.request().headers()['authorization'] || '').slice(7);
    if (!validTokens(tok)) { seen.push('401 ' + tok); return j({ error: { code: 'InvalidAuthenticationToken', message: 'x' } }, 401); }
    const path = u.replace('https://graph.microsoft.com/v1.0', '');
    seen.push('GET ' + path.slice(0, 220));
    if (path.includes('viewer_catalogue.json.gz')) return j({ id: 'c', eTag: etag, size: gz.length, lastModifiedDateTime: '2026-11-14T08:56:00Z', '@microsoft.graph.downloadUrl': 'https://dl.test/cat' });
    const m = /root:\/PhotoManager\/LargePictures\/(.+?):\/children/.exec(path);
    if (m) {
      const ai = albums.findIndex(a => a.path === m[1]);
      if (ai < 0) return j({ error: { code: 'itemNotFound', message: 'nope' } }, 404);
      return j({ value: photos.filter(q => q[1] === ai && !q[2].includes('Gone')).map(q => ({ id: 'i' + q[0], name: q[2], thumbnails: [{ c300x300_crop: { url: 'https://img.test/t/' + encodeURIComponent(q[2]) } }] })) });
    }
    const f = /root:\/PhotoManager\/LargePictures\/(.+?)\?(.*)$/.exec(path);
    if (f) {
      const name = f[1].split('/').pop();
      if (name.includes('Gone')) return j({ error: { code: 'itemNotFound', message: 'nope' } }, 404);
      if (f[2].includes('downloadUrl')) return j({ id: 'x', size: 500000, '@microsoft.graph.downloadUrl': 'https://img.test/full/' + encodeURIComponent(name) });
      return j({ id: 'x', thumbnails: [{ c300x300_crop: { url: 'https://img.test/t/' + encodeURIComponent(name) } }] });
    }
    j({ error: { code: 'bad', message: path } }, 400);
  });
  const shot = async (name) => { await p.waitForTimeout(350); if (OUT) { fs.mkdirSync(OUT, { recursive: true }); await p.screenshot({ path: `${OUT}/${name}.png` }); } };
  const vis = (id) => p.evaluate((i) => !document.getElementById(i).classList.contains('off'), id);
  const ok = (name, cond, extra = '') => { if (!cond) failed++; console.log((cond ? 'PASS ' : 'FAIL ') + name + (cond ? '' : '  ' + extra)); };

  await p.goto(BASE);
  ok('sign-in screen first; no client ID box (it is in config.js)', await vis('signin') && !(await p.isVisible('#clientId')));
  await shot('1-signin');
  await p.click('#signInBtn');
  await p.waitForSelector('#albumList .rows');
  ok('signed in with read-only scope', seen[0] === 'AUTHORIZE prompt=select_account scope=Files.Read offline_access', seen[0]);
  ok('address cleaned after sign-in', p.url() === BASE, p.url());
  const yearHeads = await p.$$eval('#albumList .year', e => e.map(x => x.textContent));
  ok('albums grouped by year, newest first, no-year last', JSON.stringify(yearHeads) === '["2019","2000","NO YEAR"]', JSON.stringify(yearHeads));
  const rows = await p.$$eval('#albumList .rows button', e => e.map(x => x.textContent));
  ok('album rows: folder names and counts', rows[0] === 'All photos' + photos.length && rows[1] === '2019 Christmas7' && rows[2] === '2019 Cornwall & Devon5' && rows[3] === '2000 Kodak620' && rows[4] === "Grandma's box2", JSON.stringify(rows));
  await shot('2-albums');

  // an album
  seen.length = 0;
  await p.click('#albumList .rows button:has-text("2019 Cornwall")');
  await p.waitForFunction(() => document.querySelectorAll('#cells img.in').length >= 4);
  ok('album: title and counts', (await p.textContent('#gridTitle')) === '2019 Cornwall & Devon4 photos, 1 video', await p.textContent('#gridTitle'));
  ok('album: one list request, one size', seen.length === 1 && seen[0].includes(":/children?$select=id,name&$expand=thumbnails($select=c300x300_crop)&$top=1000"), JSON.stringify(seen));
  const order = await p.$$eval('#cells .cell img', e => e.map(x => x.alt.slice(0, 18)));
  ok('album: oldest first, undated-with-sort-date in place', order[0].startsWith('Summer') === false && order.length === 5, JSON.stringify(order));
  ok('album: video marker with length', (await p.textContent('#cells .cell .tag')).includes('1:23'));
  ok('album: file missing in OneDrive shows not found', await p.$$eval('#cells .cell.missing', e => e.length) === 1);
  await shot('3-album');

  // a photo
  seen.length = 0;
  await p.click('#cells .cell:nth-child(2)');
  await p.waitForSelector('#stage img:not(.blur)');
  ok('photo: counter', (await p.textContent('#photoCount')) === '2 of 5', await p.textContent('#photoCount'));
  const info = await p.$$eval('#photoInfo p', e => e.map(x => x.textContent));
  ok('photo: caption, place, date, stars', info[0].startsWith('Crabbing') && info[1] === 'Harlyn Bay, Padstow, England, United Kingdom' && info[2] === '14 Jul 2019' && info[3] === '★★★★', JSON.stringify(info));
  ok('photo: Large file fetched by name', seen.some(s => s.includes('2019-07-14_Padstow_0A01.jpg?select=id,size,@microsoft.graph.downloadUrl')), JSON.stringify(seen));
  await shot('4-photo');
  await p.tap('#stage');
  ok('photo: tap hides the lines', await p.evaluate(() => document.getElementById('photo').classList.contains('bare')));
  await p.tap('#stage');
  // swipe left -> next (the video)
  const swipe = async (from, to) => { await p.evaluate(([a, b]) => { const el = document.getElementById('photo'); const t = (x) => new Touch({ identifier: 1, target: el, clientX: x, clientY: 400 });
    el.dispatchEvent(new TouchEvent('touchstart', { touches: [t(a)], changedTouches: [t(a)], bubbles: true })); el.dispatchEvent(new TouchEvent('touchend', { touches: [], changedTouches: [t(b)], bubbles: true })); }, [from, to]); };
  await swipe(300, 80);
  await p.waitForSelector('#stage video');
  ok('swipe left: next is the video, with controls', (await p.textContent('#photoCount')) === '3 of 5' && await p.getAttribute('#stage video', 'controls') !== null);
  await shot('5-video');
  await swipe(80, 300); await swipe(80, 300);
  ok('swipe right twice: first photo, shows date text', (await p.textContent('#photoCount')) === '1 of 5' && (await p.$$eval('#photoInfo p', e => e.map(x => x.textContent))).includes('Summer 2019'), await p.textContent('#photoInfo'));
  await swipe(80, 300);
  ok('no swipe before the first', (await p.textContent('#photoCount')) === '1 of 5');
  await p.keyboard.press('ArrowRight'); await p.keyboard.press('ArrowRight'); await p.keyboard.press('ArrowRight'); await p.keyboard.press('ArrowRight');
  await p.waitForFunction(() => /not found/.test(document.getElementById('stage').textContent));
  ok('photo missing in OneDrive says so', (await p.textContent('#photoCount')) === '5 of 5');
  await p.goBack();
  ok('one Back returns to the grid', await vis('grid') && !(await vis('photo')));
  await p.goBack();
  ok('Back again returns to albums', await vis('albums'));

  // the big album: scrolling, place kept
  seen.length = 0;
  await p.click('#albumList .rows button:has-text("2000 Kodak")');
  await p.waitForFunction(() => document.querySelectorAll('#cells img.in').length >= 12);
  ok('big album: first chunk only', await p.$$eval('#cells .cell', e => e.length) < 620);
  await p.evaluate(() => { document.getElementById('grid').scrollTop = 1e6; }); await p.waitForTimeout(300);
  await p.evaluate(() => { document.getElementById('grid').scrollTop = 1e6; }); await p.waitForTimeout(300);
  await p.evaluate(() => { document.getElementById('grid').scrollTop = 1e6; }); await p.waitForTimeout(300);
  ok('big album: all cells after scrolling', await p.$$eval('#cells .cell', e => e.length) === 620, String(await p.$$eval('#cells .cell', e => e.length)));
  ok('big album: still one request', seen.length === 1, JSON.stringify(seen));
  await p.evaluate(() => { document.getElementById('grid').scrollTop = 9000; }); await p.waitForTimeout(400);
  const before = await p.evaluate(() => document.getElementById('grid').scrollTop);
  await p.evaluate(() => { const g = document.getElementById('grid').getBoundingClientRect(); const c = [...document.querySelectorAll('#cells .cell')].find(c => c.getBoundingClientRect().top > g.top + 200); c.click(); });
  await p.waitForSelector('#stage img:not(.blur)');
  await shot('6-photo-plain');
  await p.goBack(); await p.waitForTimeout(200);
  ok('grid keeps its place after a photo', Math.abs(await p.evaluate(() => document.getElementById('grid').scrollTop) - before) < 5);
  await p.goBack();
  ok('albums list is back', await vis('albums'));

  // all photos: by name for a few, listing for many
  seen.length = 0;
  await p.click('#albumList .rows button:has-text("All photos")');
  await p.waitForFunction(() => document.querySelectorAll('#cells img.in').length >= 12);
  ok('all photos: undated last', true);
  ok('all photos: the album in view was listed (already known: no request)', seen.length === 0, JSON.stringify(seen));
  await p.evaluate(() => { document.getElementById('grid').scrollTop = 1e6; }); await p.waitForTimeout(300);
  await p.evaluate(() => { document.getElementById('grid').scrollTop = 1e6; }); await p.waitForTimeout(300);
  await p.evaluate(() => { document.getElementById('grid').scrollTop = 1e6; }); await p.waitForTimeout(900);
  ok('all photos: end of list fetched by list (>=5) and by name (<5)', seen.some(s => s.includes('2019 Christmas:/children')) && seen.some(s => s.includes("Grandma's box/Scans_0C01.jpg?$select=id&$expand=thumbnails")), JSON.stringify(seen));
  await shot('7-all-end');
  await p.goBack();

  // settings
  await p.click('#openSettings');
  await p.waitForFunction(() => /working|stopped/.test(document.getElementById('deviceList').textContent));
  await shot('8-settings');
  ok('settings: counts and file', (await p.textContent('#aboutList')).includes(photos.length + ' in 4 albums') && (await p.textContent('#aboutList')).includes('PhotoManager/Shared/viewer_catalogue.json.gz'), await p.textContent('#aboutList'));
  let dl = 0; p.on('request', q => { if (q.url().startsWith('https://dl.test')) dl++; });
  await p.click('#checkNow'); await p.waitForFunction(() => /up to date/.test(document.getElementById('noticeText').textContent));
  ok('check now: unchanged file is not downloaded again', dl === 0);
  etag = '"v2"';
  await p.click('#checkNow'); await p.waitForFunction(() => /New photo list/.test(document.getElementById('noticeText').textContent));
  ok('check now: changed file is downloaded', dl === 1);

  // reload: uses the copy on the device; token still valid
  seen.length = 0; tokenCalls.length = 0;
  await p.reload(); await p.waitForSelector('#deviceList dd');
  ok('reload: stays on Settings, no sign-in, one check request', await vis('settings') && tokenCalls.length === 0 && seen.length === 1, JSON.stringify(seen));
  // expired access token: renewed in the background
  await p.evaluate(() => localStorage.setItem('pv_expires', '0'));
  await p.goBack(); await p.reload(); await p.waitForSelector('#albumList .rows');
  ok('expired token renewed with the refresh token', tokenCalls.join() === 'refresh_token', tokenCalls.join());
  // OneDrive rejects the token (401): renewed once and retried
  tokenCalls.length = 0; const bad = await p.evaluate(() => localStorage.getItem('pv_token')); validTokens = (t) => t !== bad;
  await p.reload(); await p.waitForSelector('#albumList .rows'); await p.waitForTimeout(300);
  ok('401 from OneDrive: token renewed and request retried', tokenCalls.join() === 'refresh_token' && !(await vis('signin')), tokenCalls.join());
  // refresh refused (the 24-hour limit): quiet trip to Microsoft and back
  validTokens = () => true; refreshOk = false; seen.length = 0;
  await p.evaluate(() => localStorage.setItem('pv_expires', '0'));
  await p.reload(); await p.waitForSelector('#albumList .rows'); await p.waitForTimeout(300);
  ok('refresh refused: quiet sign-in and back to albums', seen.includes('AUTHORIZE prompt=none scope=Files.Read offline_access') && await vis('albums'), JSON.stringify(seen.slice(0, 3)));
  // and when Microsoft wants the account chosen again
  quietOk = false; await p.evaluate(() => localStorage.setItem('pv_expires', '0'));
  await p.reload(); await p.waitForFunction(() => /sign in again/.test(document.getElementById('signinError').textContent));
  ok('quiet sign-in refused: the Sign in button, no loop', await vis('signin'));
  quietOk = true; refreshOk = true;
  await p.click('#signInBtn'); await p.waitForSelector('#albumList .rows');
  ok('signing in again works', await vis('albums'));
  // sign out
  await p.click('#openSettings'); await p.click('#signOut'); await p.waitForTimeout(300);
  const left = await p.evaluate(async () => ({ t: localStorage.getItem('pv_token'), r: localStorage.getItem('pv_refresh'), c: await new Promise(d => { const q = indexedDB.open('photo-viewer', 1); q.onsuccess = () => { const g = q.result.transaction('kv').objectStore('kv').get('catalogue'); g.onsuccess = () => d(g.result === undefined); }; }) }));
  ok('sign out forgets the sign-in and the photo list', await vis('signin') && left.t === null && left.r === null && left.c === true, JSON.stringify(left));
  console.log(log.filter(l => !/Failed to load resource/.test(l)).join('\n') || 'no console errors');
  await b.close();
  server.close();
  console.log(failed ? `\n${failed} FAILED` : '\nALL PASSED');
  process.exit(failed ? 1 : 0);
})();
