/* Checks for the Photo Viewer, run at iPhone size against a stand-in for Microsoft sign-in and OneDrive.
   Nothing here touches a real account: every request to Microsoft is answered by this script.

   Needs Node.js with Playwright (and its Chromium).   Run from the repository folder:
       node test/run.js [light|dark] [folder for screenshots]
   (Set PW_CHROME to a Chrome executable to use it instead of Playwright's own Chromium.)
   Prints PASS / FAIL lines and ends with a non-zero code if anything failed. */
const { chromium, devices } = require('playwright');
const zlib = require('zlib');
const http = require('http');
const fs = require('fs');
const path = require('path');
const ROOT = path.resolve(__dirname, '..');
const SCHEME = process.argv[2] === 'dark' ? 'dark' : 'light';
const OUT = process.argv[3] || '';
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.png': 'image/png', '.json': 'application/json' };
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
  const add = (a, file, o = {}) => photos.push([++n, a, file, o.video ? 1 : 0, o.length ?? null, o.date ?? "", o.date_text ?? "", o.date_sort ?? "", o.caption ?? "", o.site ?? "", o.city ?? "", o.state ?? "", o.country ?? "", o.long ?? "", o.rating ?? 0, null, null, o.keywords ?? "", o.people ?? ""]);
  for (let i = 1; i <= 620; i++) add(0, `2000-01-${String(i % 28 + 1).padStart(2, "0")}_Kodak_${String(i).padStart(4, "0")}.jpg`, { date: `2000-01-${String(i % 28 + 1).padStart(2, "0")}T10:00:00`, caption: i % 3 ? "" : "Photo " + i });
  add(1, "2019-07-14_Padstow_0A01.jpg", { date: "2019-07-14T09:30:00", caption: "Crabbing off the harbour wall before breakfast, with the tide coming in", site: "Harlyn Bay", city: "Padstow", state: "England", country: "United Kingdom", rating: 4, long: "North coast path", keywords: "seaside; crabs", people: "Zoë; Tom" });
  add(1, "2019-07-14_Padstow_0A02.mp4", { video: 1, length: 83, date: "2019-07-14T12:00:00", caption: "Crab race", city: "Padstow", country: "United Kingdom" });
  add(1, "2019-Cornwall_0A03.jpg", { date_text: "Summer 2019", date_sort: "2019-07-01" });
  add(1, "2019-07-15_Padstow_0A04.jpg", { date: "2019-07-15T15:00:00" });
  add(1, "2019-07-16_Gone_0A05.jpg", { date: "2019-07-16T15:00:00", caption: "Not in OneDrive" });
  for (let i = 1; i <= 7; i++) add(2, `2019-12-25_Home_0B0${i}.jpg`, { date: "2019-12-25T1" + i + ":00:00", caption: "Christmas " + i, rating: i % 6 });
  add(2, "2019-Christmas_0B08.jpg", { date_text: "Christmas 2019", date_sort: "2019-12-25", caption: "Christmas, no camera date" });
  add(3, "Scans_0C01.jpg", { date_text: "about 1962", caption: "Grandma at the seaside", site: "Café Régence", city: "Besançon", country: "France", rating: 5 });
  add(3, "Scans_0C02.jpg", {});
  const cat = { format: 1, written_at: "2026-11-14T08:55:00Z", machine: "ASPIRE", pictures_path: "PhotoManager/LargePictures", album_count: 4, photo_count: photos.length, albums, columns, photos };
  const gz = zlib.gzipSync(JSON.stringify(cat));
  let etag = '"v1"';

  const b = await chromium.launch(process.env.PW_CHROME ? { executablePath: process.env.PW_CHROME } : {});
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
  const rows = await p.$$eval('#albumList .rows button:not([hidden])', e => e.map(x => x.textContent));
  ok('album rows: folder names and counts', rows[0] === 'All photos' + photos.length && rows[1] === '2019 Christmas8' && rows[2] === '2019 Cornwall & Devon5' && rows[3] === '2000 Kodak620' && rows[4] === "Grandma's box2", JSON.stringify(rows));
  await shot('2-albums');

  // an album
  seen.length = 0;
  await p.click('#albumList .rows button:has-text("2019 Cornwall")');
  await p.waitForFunction(() => document.querySelectorAll('#cells img.in').length >= 4);
  ok('album: title and counts', (await p.textContent('#gridTitle')) === '2019 Cornwall & Devon4 photos, 1 video', await p.textContent('#gridTitle'));
  ok('album: one list request, one size', seen.length === 1 && seen[0].includes(":/children?$select=id,name&$expand=thumbnails($select=c300x300_crop)&$top=1000"), JSON.stringify(seen));
  const order = await p.$$eval('#cells .cell img', e => e.map(x => x.alt.slice(0, 18)));
  ok('album: oldest first, undated-with-sort-date in place', order[0].startsWith('Summer') === false && order.length === 5, JSON.stringify(order));
  const xmas = await p.evaluate(() => listFor(2).map((i) => CAT.data.photos[i][CAT.col.file].slice(-8)));
  ok('album: same day, the photo with only a sort date follows the camera photos', xmas.length === 8 && xmas[7] === '0B08.jpg' && xmas[0] === '0B01.jpg', JSON.stringify(xmas));
  const xmasNew = await p.evaluate(() => { const f = FILTER; FILTER = { ...f, album: '2019/2019 Christmas', order: 'new' }; CAT.lists.delete('filter'); const r = listFor('filter').map((i) => CAT.data.photos[i][CAT.col.file].slice(-8)); FILTER = f; CAT.lists.delete('filter'); return r; });
  ok('newest first: the same day reversed, the sort-date-only photo first', xmasNew[0] === '0B08.jpg' && xmasNew[7] === '0B01.jpg', JSON.stringify(xmasNew));
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

  // ---- the Filter (spec 2.11.5)
  const count = () => p.textContent('#filterCount');
  const type = async (id, text) => { await p.fill(id, text); await p.waitForTimeout(200); };
  const pick = async (id, value) => { await p.selectOption(id, value); await p.waitForTimeout(50); };
  seen.length = 0;
  ok('albums: no Filter line while nothing is chosen', !(await p.isVisible('#filterRow')));
  await p.click('#openFilter');
  ok('filter: opens with everything matching', await vis('filter') && (await count()) === '634 photos, 1 video', await count());
  const albumChoices = await p.$$eval('#fAlbum option', e => e.map(x => x.textContent));
  ok('filter: album choices in the Albums order', JSON.stringify(albumChoices) === JSON.stringify(['All albums', '2019 Christmas', '2019 Cornwall & Devon', '2000 Kodak', "Grandma's box"]), JSON.stringify(albumChoices));
  ok('filter: suggestions from the Sites and places in use', JSON.stringify(await p.$$eval('#siteNames option', e => e.map(x => x.value))) === '["Café Régence","Harlyn Bay"]' && (await p.$$eval('#placeNames option', e => e.map(x => x.value))).join() === 'Besançon,England,France,Padstow,United Kingdom');
  ok('filter: year hints from the photo list', await p.getAttribute('#fFrom', 'placeholder') === '2000' && await p.getAttribute('#fTo', 'placeholder') === '2019');
  await type('#fText', 'padstow');
  ok('search: city (file names are not searched)', (await count()) === '1 photo, 1 video', await count());
  await type('#fText', '  HARBOUR  crabbing ');
  ok('search: several words, capitals and spaces ignored', (await count()) === '1 photo', await count());
  await type('#fText', 'zoe');
  ok('search: people, accents ignored', (await count()) === '1 photo', await count());
  await type('#fText', 'crabs');
  ok('search: keywords', (await count()) === '1 photo', await count());
  await type('#fText', 'coast path');
  ok('search: Location (long)', (await count()) === '1 photo', await count());
  await type('#fText', 'devon');
  ok('search: album name', (await count()) === '4 photos, 1 video', await count());
  await type('#fText', 'grandma padstow');
  ok('search: every word must be found', (await count()) === 'No photos match' && await p.isDisabled('#filterShow'), await count());
  await type('#fText', '');
  await type('#fSite', 'cafe reg');
  ok('site: part of the name, accents ignored', (await count()) === '1 photo', await count());
  await type('#fSite', '');
  await type('#fPlace', 'besancon');
  ok('location: city', (await count()) === '1 photo', await count());
  await type('#fPlace', 'kingdom');
  ok('location: country', (await count()) === '1 photo, 1 video', await count());
  await type('#fPlace', '');
  await type('#fFrom', '2019');
  ok('years: from; sort date used, undated left out', (await count()) === '12 photos, 1 video', await count());
  await type('#fTo', '2000'); 
  ok('years: the wrong way round is put right', (await count()) === '632 photos, 1 video', await count());
  await type('#fFrom', ''); 
  ok('years: up to', (await count()) === '620 photos', await count());
  await type('#fTo', '');
  await pick('#fRating', '4');
  ok('rating: stars or more', (await count()) === '4 photos', await count());
  await pick('#fRating', '0');
  await pick('#fType', 'video');
  ok('type: videos only', (await count()) === '1 video', await count());
  await pick('#fType', 'photo');
  await pick('#fAlbum', '2019/2019 Cornwall & Devon');
  ok('album and type combined with and', (await count()) === '4 photos', await count());
  ok('filter: counting made no request to OneDrive', seen.length === 0, JSON.stringify(seen));
  await shot('7a-filter');
  await p.click('#filterClear');
  ok('clear empties every choice', (await count()) === '634 photos, 1 video' && await p.inputValue('#fAlbum') === '' && await p.inputValue('#fType') === '');
  // results from several albums, newest first
  await pick('#fRating', '3'); await pick('#fOrder', 'new');
  await p.click('#filterShow');
  await p.waitForFunction(() => document.querySelectorAll('#cells img.in').length >= 5);
  ok('results: title, count and summary', (await p.textContent('#gridTitle')) === 'Filter5 photos' && (await p.textContent('#gridSummary')) === '3 stars or more · Newest first', (await p.textContent('#gridTitle')) + ' | ' + await p.textContent('#gridSummary'));
  const got = await p.$$eval('#cells .cell img', e => e.map(x => x.alt));
  ok('results: newest first, undated last', JSON.stringify(got) === JSON.stringify(['Christmas 5', 'Christmas 4', 'Christmas 3', 'Crabbing off the harbour wall before breakfast, with the tide coming in', 'Grandma at the seaside']), JSON.stringify(got));
  await shot('7b-filter-results');
  await p.click('#cells .cell:nth-child(4)');
  await p.waitForSelector('#stage img:not(.blur)');
  ok('results: photo screen moves through the Filter list', (await p.textContent('#photoCount')) === '4 of 5');
  await p.goBack();
  // changing the choices from the results and coming back
  await p.click('#gridFilter');
  ok('results: Filter button reopens the choices as left', await vis('filter') && await p.inputValue('#fRating') === '3' && await p.inputValue('#fOrder') === 'new');
  await pick('#fOrder', 'name'); await type('#fText', 'christmas');
  await p.click('#filterShow');
  await p.waitForFunction(() => document.getElementById('gridSummary').textContent.includes('christmas'));
  const got2 = await p.$$eval('#cells .cell img', e => e.map(x => x.alt));
  ok('results redrawn: by file name', await vis('grid') && got2.join() === 'Christmas 3,Christmas 4,Christmas 5', got2.join());
  await p.goBack();
  ok('one Back from the results returns to the Filter', await vis('filter'));
  await p.goBack();
  ok('and another to Albums, which now has a Filter line', await vis('albums') && (await p.textContent('#filterRow')) === 'Filter“christmas” · 3 stars or more · By file name3', await p.textContent('#filterRow'));
  await shot('7c-albums-filter-line');
  await p.reload(); await p.waitForSelector('#albumList .rows');
  ok('the Filter is remembered on the device', await p.isVisible('#filterRow'));
  await p.click('#filterRow'); await p.waitForFunction(() => document.querySelectorAll('#cells .cell').length === 3);
  ok('the Filter line opens its photos', (await p.textContent('#gridTitle')) === 'Filter3 photos');
  await p.goBack();
  // a Filter that names one album lists that album; a few photos from elsewhere are found by name
  await p.evaluate(() => thumbs.reset());
  await p.click('#openFilter'); await p.click('#filterClear'); await pick('#fAlbum', '2000/2000 Kodak'); await type('#fText', 'photo');
  seen.length = 0;
  await p.click('#filterShow'); await p.waitForFunction(() => document.querySelectorAll('#cells img.in').length >= 12);
  ok('one album chosen: one list request', seen.length === 1 && seen[0].includes('2000 Kodak:/children'), JSON.stringify(seen));
  await p.click('#gridFilter'); await p.click('#filterClear'); await type('#fPlace', 'padstow');
  seen.length = 0;
  await p.click('#filterShow'); await p.waitForFunction(() => document.querySelectorAll('#cells img.in').length >= 2);
  ok('a few photos: fetched by file name', seen.length === 2 && seen.every(s => s.includes('?$select=id&$expand=thumbnails')), JSON.stringify(seen));
  await p.goBack(); await p.goBack();

  // settings
  await p.click('#openSettings');
  ok('settings: caption, place and date switches all on to start with', await p.isChecked('#showCaption') && await p.isChecked('#showPlace') && await p.isChecked('#showDate'));
  await p.click('#showPlace'); await p.click('#showDate');
  await p.goBack();
  await p.click('#albumList .rows button:has-text("2019 Cornwall")'); await p.click('#cells .cell:nth-child(2)');
  await p.waitForSelector('#stage img:not(.blur)');
  const lines = await p.$$eval('#photoInfo p', e => e.map(x => x.textContent));
  ok('switches: place and date hidden, caption and stars kept', lines.length === 2 && lines[0].startsWith('Crabbing') && lines[1] === '★★★★', JSON.stringify(lines));
  await p.goBack(); await p.goBack();
  await p.click('#openSettings');
  ok('switches are remembered', !(await p.isChecked('#showPlace')) && await p.isChecked('#showCaption'));
  await p.click('#showPlace'); await p.click('#showDate');
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
  // ---- a new version of the viewer
  const build = /const BUILD = "([^"]+)"/.exec(fs.readFileSync(path.join(ROOT, 'app.js'), 'utf8'))[1];
  const vfile = JSON.parse(fs.readFileSync(path.join(ROOT, 'version.json'), 'utf8')).version;
  const vhtml = [...fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8').matchAll(/\.js\?v=([^"]+)"/g)].map(m => m[1]);
  ok('version: app.js, version.json and index.html agree', vfile === build && vhtml.length === 2 && vhtml.every(v => v === build), [build, vfile, vhtml].join(' / '));
  ok('version: same version, no notice', !(await p.isVisible('#noticeAct')));
  await ctx.route(BASE + 'version.json*', r => r.fulfill({ status: 200, contentType: 'application/json', body: '{"version":"9.9"}' }));
  await p.reload(); await p.waitForSelector('#noticeAct:not([hidden])');
  ok('version: a newer one is offered', (await p.textContent('#noticeText')).includes('9.9') && (await p.textContent('#noticeAct')) === 'Update');
  await shot('9-update');
  const loaded = []; p.on('request', q => { if (/\.js\?v=|\?v=9\.9/.test(q.url())) loaded.push(q.url().replace(BASE, '')); });
  await p.click('#noticeAct'); await p.waitForSelector('#albumList .rows');
  ok('version: Update opens the page at a new address, scripts with ?v=', loaded.includes('?v=9.9') && loaded.some(u => u.startsWith('app.js?v=')), JSON.stringify(loaded));
  ok('version: the ?v= is dropped from the address, still signed in', p.url() === BASE && await vis('albums'), p.url());
  await ctx.unroute(BASE + 'version.json*');
  // sign out
  await p.click('#openSettings'); await p.click('#signOut'); await p.waitForTimeout(300);
  const left = await p.evaluate(async () => ({ t: localStorage.getItem('pv_token'), r: localStorage.getItem('pv_refresh'), f: localStorage.getItem('pv_filter'), s: localStorage.getItem('pv_show'), c: await new Promise(d => { const q = indexedDB.open('photo-viewer', 1); q.onsuccess = () => { const g = q.result.transaction('kv').objectStore('kv').get('catalogue'); g.onsuccess = () => d(g.result === undefined); }; }) }));
  ok('sign out forgets the sign-in, the photo list, the Filter and the switches', await vis('signin') && left.t === null && left.r === null && left.f === null && left.s === null && left.c === true, JSON.stringify(left));
  console.log(log.filter(l => !/Failed to load resource/.test(l)).join('\n') || 'no console errors');
  await b.close();
  server.close();
  console.log(failed ? `\n${failed} FAILED` : '\nALL PASSED');
  process.exit(failed ? 1 : 0);
})();
