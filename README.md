# Photo Viewer

A view-only web page for looking through a family photo collection kept in OneDrive, made for a phone.

- It signs in with a Microsoft account and asks for one permission: **read files**. It cannot change or delete anything.
- It reads a catalogue file and the pictures straight from OneDrive. There is no server.
- This repository holds only the viewer's code. It contains no photos and no information about them.

The specification is section 2.11 of the Photo Manager specification, which is kept with the desktop applications.

| File | What it is |
| --- | --- |
| `index.html` | The page and its styles |
| `app.js` | The viewer |
| `config.js` | Settings: the Microsoft registration's client ID and the OneDrive folder holding the catalogue file |
| `apple-touch-icon.png` | The icon used when the page is added to a phone's home screen |
