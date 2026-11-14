# Photo Viewer

A view-only web page for a family photo collection in OneDrive. Read `docs/architecture.md` before changing anything: it explains the code, the requests to OneDrive that the trial showed to be fast, and what is still to be confirmed.

- **What it must do** is section 2.11 of `docs/caption_spec.md` in the private repository `NickM64/PhotoManagerPrototype`. Behaviour changes go into that specification (version number and change log) as well as here.
- **View only.** The only Microsoft permission is `Files.Read`. Never add a permission or anything that writes to OneDrive.
- **No libraries, no build step, no server.** Plain `index.html`, `app.js`, `config.js`. The page talks only to Microsoft.
- **The repository is public.** Never commit photos, catalogue data or anything personal.
- **Changes go on a branch with a pull request.** The owner tests and merges; merging to `main` publishes the page.
- **Every change that alters `app.js`, `index.html` or `config.js` raises the version** in three places that must agree: `BUILD` in `app.js`, `version.json`, and the two `?v=` in `index.html`. Phones only pick up a new version through this (architecture section 6a).
- **Before a pull request,** run `node test/run.js` and `node test/run.js dark`, and keep `docs/architecture.md` in step with the code.
- **Writing for the owner:** plain words, British spelling, step-by-step instructions for anything on GitHub or the Microsoft registration.
