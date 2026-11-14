# Photo Viewer – how it is built

**Applies to:** version 0.2 (stage 2), November 2026.
**What it must do:** section 2.11 of `docs/caption_spec.md` in the private repository `NickM64/PhotoManagerPrototype`. That specification is the authority; this document explains the code.

## 1. The pieces

| Piece | Where | Notes |
| --- | --- | --- |
| The page | `index.html` | Markup and all styles. Four screens (`signin`, `albums`, `grid`, `settings`) and the photo overlay (`photo`). Light and dark follow the device. |
| The viewer | `app.js` | Plain JavaScript, no libraries, no build step. |
| Settings | `config.js` | `clientId` (the Microsoft registration) and `catalogueFolder` (the OneDrive folder holding `viewer_catalogue.json.gz`). Neither is a secret. |
| Icon | `apple-touch-icon.png` | Used when the page is added to a phone's home screen. |
| Checks | `test/run.js` | Playwright at iPhone size against a stand-in for Microsoft and OneDrive (section 8). |

It is published by GitHub Pages from the `main` branch at `https://nickm64.github.io/photo-viewer/`. Merging to `main` publishes. The repository is public and holds no photos and no catalogue data.

## 2. Where the information comes from

- **Viewer Catalogue File.** `viewer_catalogue.json.gz` is written by the desktop applications (`pm_viewer.py` in the Photo Manager repository, spec 2.11.2) into the shared OneDrive folder. Format 1: a header (`format`, `written_at`, `machine`, `pictures_path`), `albums` (`path`, `title`, `year`), `columns`, and `photos` (one list per photo, one value per column).
- **A photo's file in OneDrive** is `pictures_path / albums[album].path / file`.
- **`loadCatalogue`** shows the copy kept on the device at once (IndexedDB), then asks OneDrive for the file's `eTag` (one request) and downloads it only when that has changed. `useCatalogue` checks the format and columns and builds the album counts.
- **`listFor(key)`** gives the photos of one album (a number) or of everything (`"all"`), oldest first by `date` or `date_sort`, undated last, then by file name. Lists are kept in `CAT.lists`.

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
- Opening an album lists it straight away.
- A file in the catalogue but not in OneDrive shows *not found*. A small-picture address that stops working is looked up again once.
- Rejected in the trial: asking for every picture size (4 s against 1.6 s), and Graph's `$batch` (slowest route on laptop and phone).

## 5. Screens and moving between them

- Each screen is a fixed, separately scrolling layer; a hidden one keeps its scroll position (`showLayer`).
- `go(state)` pushes a history entry and `render(state)` draws it, so the phone's Back gesture works. State is `{ v, album, photo }`. Moving between photos replaces the entry, so one Back returns to the grid.
- **Grid** (`grid`): cells are added 240 at a time as the end comes near (`addChunk`); an `IntersectionObserver` asks for each picture.
- **Photo** (`photo`): the small picture is shown blurred until the Large file arrives; the next photo is fetched ahead. Swipe, arrow keys, tap to hide the text. Enlarging uses the browser's own pinch; swipes are ignored while enlarged. A video is a `<video>` element with the device's controls.
- **Settings** (`drawSettings`): catalogue details, *Check for new information now*, *Sign out*, and the picture address test.

## 6. What is kept on the device

| Key | Where | What |
| --- | --- | --- |
| `catalogue` | IndexedDB `photo-viewer` / `kv` | The catalogue and its `eTag` |
| `pv_token`, `pv_expires`, `pv_refresh`, `pv_signed_in` | localStorage | The sign-in |
| `pv_client` | localStorage | Client ID typed in, only used when `config.js` has none |
| `pv_addrtest` | localStorage | The picture address test (section 7) |
| `pv_verifier`, `pv_state`, `pv_quiet`, `pv_quiet_tried` | sessionStorage | One sign-in in progress |

*Sign out* removes the sign-in and the catalogue. Every storage call is wrapped, so the viewer still works when storage is unavailable.

## 7. Still to be confirmed in use

- **The 24-hour sign-in.** Whether the quiet trip to Microsoft works on an iPhone with a personal account, from Safari and from the home-screen icon.
- **How long a small-picture address stays valid.** Settings shows the result of `addressTest`. If it is hours, addresses can be kept between visits so an album reopens at once.
- **Videos on a phone.** Playing straight from OneDrive's download address has only been checked against the stand-in.

## 8. Checks

`node test/run.js [light|dark] [folder for screenshots]` needs Node.js with Playwright. It serves the repository's files itself and answers every Microsoft address from the script, with a made-up catalogue of 634 photos in four albums. It covers sign-in, renewal, the 24-hour path, albums, grid, scrolling, photo, swipe, video, missing files, settings and sign out. It cannot show how real OneDrive or a real phone behaves; that is tested by hand after each merge.

## 9. Stages (spec 2.11.10)

| Stage | What | State |
| --- | --- | --- |
| 1 | Viewer Catalogue File written by the desktop applications | Done (Photo Manager PR #39) |
| 2 | Sign-in, catalogue, Albums, Grid, Photo; videos play | Done (PR #1) |
| 3 | Filter, and the caption / place / date switches in Settings | Next |
| 4 | Anything found in use; keeping addresses between visits if section 7 allows | Later |
