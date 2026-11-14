# Photo Viewer – how it is built

**Applies to:** version 0.3.1 (stage 3), November 2026.
**What it must do:** section 2.11 of `docs/caption_spec.md` in the private repository `NickM64/PhotoManagerPrototype`. That specification is the authority; this document explains the code.

## 1. The pieces

| Piece | Where | Notes |
| --- | --- | --- |
| The page | `index.html` | Markup and all styles. Five screens (`signin`, `albums`, `grid`, `filter`, `settings`) and the photo overlay (`photo`). Light and dark follow the device. |
| The viewer | `app.js` | Plain JavaScript, no libraries, no build step. |
| Version | `version.json` | The viewer's version, read past every cache to offer updates (section 6a). Must match `BUILD` in `app.js` and the `?v=` in `index.html`. |
| Settings | `config.js` | `clientId` (the Microsoft registration) and `catalogueFolder` (the OneDrive folder holding `viewer_catalogue.json.gz`). Neither is a secret. |
| Icon | `apple-touch-icon.png` | Used when the page is added to a phone's home screen. |
| Checks | `test/run.js` | Playwright at iPhone size against a stand-in for Microsoft and OneDrive (section 8). |

It is published by GitHub Pages from the `main` branch at `https://nickm64.github.io/photo-viewer/`. Merging to `main` publishes. The repository is public and holds no photos and no catalogue data.

## 2. Where the information comes from

- **Viewer Catalogue File.** `viewer_catalogue.json.gz` is written by the desktop applications (`pm_viewer.py` in the Photo Manager repository, spec 2.11.2) into the shared OneDrive folder. Format 1: a header (`format`, `written_at`, `machine`, `pictures_path`), `albums` (`path`, `title`, `year`), `columns`, and `photos` (one list per photo, one value per column).
- **A photo's file in OneDrive** is `pictures_path / albums[album].path / file`.
- **`loadCatalogue`** shows the copy kept on the device at once (IndexedDB), then asks OneDrive for the file's `eTag` (one request) and downloads it only when that has changed. `useCatalogue` checks the format and columns and builds the album counts.
- **`listFor(key)`** gives the photos of one album (a number), of everything (`"all"`) or of the Filter (`"filter"`, section 5a). Albums and everything are oldest first by `date` or `date_sort`, undated last, then by file name; the Filter's list is in the order the Filter asks for. Lists are kept in `CAT.lists`.

## 3. Signing in

- Microsoft's authorization-code flow with PKCE for single-page applications, written by hand (`goToSignIn`, `finishSignIn`, `tokenRequest`). No Microsoft library is used: version 5 of theirs needs extra pages, and this is about 80 lines.
- Authority `https://login.microsoftonline.com/consumers` (personal accounts). Scope `Files.Read offline_access`.
- The redirect address is the page's own address without `index.html`. Each address the viewer is opened from must be in the registration as a *Single-page application* redirect URI. Registered: `https://nickm64.github.io/photo-viewer/` and `http://localhost:8000/`.
- **Renewal.** `accessToken` renews with the refresh token when the access token has run out. When Microsoft refuses (its 24-hour limit for web pages), `onLoadError` makes one quiet trip to Microsoft (`prompt=none`) and back; if that is refused too, the Sign in button is shown. `pv_quiet_tried` in session storage stops a loop.
- `graph(url)` adds the token, renews once on 401, and waits and retries on 429 / 503.

## 4. Getting pictures (spec 2.11.6)

Each request to OneDrive costs about half a second whatever its size (trial, 13 November 2026), so the code makes as few as it can.

| Need | Request | Code |
| --- | --- | --- |
| Small pictures for an album | `/me/drive/root:/<folder>:/children?$select=id,name&$expand=thumbnails($select=c300x300_crop)&$top=1000` (one request, about 2 s for 220 photos) | `thumbs.listAlbum` |
| A small picture by file name | `/me/drive/root:/<file>?$select=id&$expand=thumbnails($select=c300x300_crop)` | `thumbs.byName`, at most 8 at a time |
| A Large file | `/me/drive/root:/<file>?select=id,size,@microsoft.graph.downloadUrl`, then the address it returns | `fullAddress.get`, reused for 40 minutes |

- `thumbs.want(i, img)` is called when a grid cell comes within two screens of view. Requests are gathered for 60 ms, then grouped by album: 5 or more wanted from one album lists the album, fewer are looked up by name.
- Opening an album lists it straight away. So does a Filter that names one album, when it matches 5 photos or more. Other Filter results use `thumbs.want` like *All photos*: photos spread over many albums are looked up by name, 8 at a time, and an album with 5 or more in view is listed.
- A file in the catalogue but not in OneDrive shows *not found*. A small-picture address that stops working is looked up again once.
- Rejected in the trial: asking for every picture size (4 s against 1.6 s), and Graph's `$batch` (slowest route on laptop and phone).

## 5. Screens and moving between them

- Each screen is a fixed, separately scrolling layer; a hidden one keeps its scroll position (`showLayer`).
- `go(state)` pushes a history entry and `render(state)` draws it, so the phone's Back gesture works. State is `{ v, album, photo, fromGrid }`; `album` is an album number, `"all"` or `"filter"`. Moving between photos replaces the entry, so one Back returns to the grid.
- **Grid** (`grid`): cells are added 240 at a time as the end comes near (`addChunk`); an `IntersectionObserver` asks for each picture.
- **Photo** (`photo`): the small picture is shown blurred until the Large file arrives; the next photo is fetched ahead. Swipe, arrow keys, tap to hide the text. Enlarging uses the browser's own pinch; swipes are ignored while enlarged. A video is a `<video>` element with the device's controls.
- **Settings** (`drawSettings`): the *Show caption*, *Show place* and *Show date* switches (`SHOW`, used by `photo.show`; the stars are always shown), catalogue details, *Check for new information now*, *Sign out*, and the picture address test.

## 5a. The Filter (spec 2.11.5)

- **The choices** are one object, `FILTER` (`text`, `album` (the album's path, so it survives a new catalogue), `site`, `location`, `from`, `to`, `rating`, `type`, `order`). `saveFilter` keeps it on the device and drops the Filter's list, so the grid is drawn again when next shown.
- **`filterMatches(f)`** reads the catalogue only; it makes no request. `plain()` removes capitals, accents and spaces at either end before comparing. Search uses one line of text per photo (`searchText`: caption, Site, City, State, Country, Location (long), album path and title, keywords, people), made the first time Search is used and kept until the catalogue changes. File names are not searched.
- **Years** use the first four characters of `date`, or `date_sort`; a photo with neither is left out while a year is given. Years typed the wrong way round are swapped. Anything that is not four digits is ignored.
- **Order:** oldest first, newest first (undated last in both), or by file name alone.
- **The Filter screen** (`filterForm`): every change is saved and counted at once (text after 120 ms). Album choices are in the Albums screen's order; Site and Location suggestions (`<datalist>`) and the year hints come from the catalogue and are made once per catalogue. *Show* is greyed when nothing matches.
- **Where it shows:** the Filter's photos are a list of their own, `{ v: "grid", album: "filter" }`. Opening an album from the Albums screen is never filtered. While a choice is made, the Albums screen has a *Filter* line under *All photos* with the summary and count (`drawFilterRow`).
- **Back:** Albums > Filter > results is three history entries. The *Filter* button above the results opens the Filter screen with `fromGrid`, and *Show* then goes back one entry instead of adding one, so Back from the results always leads to the Filter screen and then Albums.

## 6. What is kept on the device

| Key | Where | What |
| --- | --- | --- |
| `catalogue` | IndexedDB `photo-viewer` / `kv` | The catalogue and its `eTag` |
| `pv_token`, `pv_expires`, `pv_refresh`, `pv_signed_in` | localStorage | The sign-in |
| `pv_client` | localStorage | Client ID typed in, only used when `config.js` has none |
| `pv_addrtest` | localStorage | The picture address test (section 7) |
| `pv_filter` | localStorage | The Filter choices |
| `pv_show` | localStorage | The caption / place / date switches |
| `pv_verifier`, `pv_state`, `pv_quiet`, `pv_quiet_tried` | sessionStorage | One sign-in in progress |

*Sign out* removes the sign-in, the catalogue, the Filter choices and the switches. Every storage call is wrapped, so the viewer still works when storage is unavailable.

## 6a. Getting a new version onto the phone

Found in use (14 November 2026): after a merge, an iPhone's home-screen icon kept showing the old version, even after the app was swiped away, while a laptop browser showed the new one. GitHub Pages lets browsers keep its files for 10 minutes, and a home-screen app may keep them longer.

- `version.json` holds the version. `update.check` asks for it with `cache: "no-store"` and a time on the address, so neither the phone nor GitHub's servers answer from a cache.
- It is checked at start-up, when the viewer is brought back to the screen (at most every 5 minutes), and by *Check for new information now* in Settings.
- When it differs from `BUILD`, a notice offers **Update**. That opens the page at `?v=<version>`, an address no cache holds, and `index.html` asks for `app.js?v=` and `config.js?v=`, so all three come fresh. The `?v=` is removed from the address at start-up. The sign-in, photo list and choices are kept.
- So **each release raises the version in `app.js` (`BUILD`), `version.json` and `index.html` (both `?v=`)**. `test/run.js` fails if they differ.

## 7. Still to be confirmed in use

- **The 24-hour sign-in: found in use (14 November 2026).** After 24 hours the viewer asks for the account to be confirmed and does not ask for the password. That is acceptable and nothing more is planned.
- **The Filter on the full collection.** The automatic checks use 634 photos. With a made-up list of 50,000 photos on a laptop, the first Search took about 80 ms and later counts about 25 ms, so the count should follow the typing on a phone; to be checked there.
- **How long a small-picture address stays valid.** Settings shows the result of `addressTest`. If it is hours, addresses can be kept between visits so an album reopens at once.
- **Videos on a phone.** Playing straight from OneDrive's download address has only been checked against the stand-in.

## 8. Checks

`node test/run.js [light|dark] [folder for screenshots]` needs Node.js with Playwright. It serves the repository's files itself and answers every Microsoft address from the script, with a made-up catalogue of 634 photos in four albums. It covers sign-in, renewal, the 24-hour path, albums, grid, scrolling, photo, swipe, video, missing files, every Filter choice, the Filter's results and history, the Settings switches and sign out. It cannot show how real OneDrive or a real phone behaves; that is tested by hand after each merge.

## 9. Stages (spec 2.11.10)

| Stage | What | State |
| --- | --- | --- |
| 1 | Viewer Catalogue File written by the desktop applications | Done (Photo Manager PR #39) |
| 2 | Sign-in, catalogue, Albums, Grid, Photo; videos play | Done (PR #1) |
| 3 | Filter, and the caption / place / date switches in Settings | Done (PR #3) |
| 4 | Anything found in use; keeping addresses between visits if section 7 allows | Next |
