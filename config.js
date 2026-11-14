// Settings for the Photo Viewer (spec 2.11.8).  Nothing here is a password:
// neither value gives access to anything without signing in to the Microsoft account.
window.PV_CONFIG = {
  // Application (client) ID of the "Photo Viewer" registration with Microsoft.
  // While this is empty the viewer asks for it once on each device and remembers it there.
  clientId: "",

  // Folder in OneDrive that holds viewer_catalogue.json.gz (the shared_sync folder), with forward slashes.
  catalogueFolder: "PhotoManager/Shared",
};
