// DriveHelper.gs
// Shared helper for uploading a data-URL payload (as produced by the
// frontend's FileReader-based readAttachmentFile() / the signature canvas's
// toDataURL()) into the attachments Drive folder, returning a URL usable
// directly as an <img src> or download link.

function getAttachmentsFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(ATTACHMENTS_FOLDER_PROP_);
  if (!id) throw new Error('Attachments folder not set up yet. Run setupSpreadsheet() first.');
  return DriveApp.getFolderById(id);
}

// payload: { name, type, data } where data is a data-URL string
// ("data:<mime>;base64,<...>"). Returns a direct-content URL — NOT
// file.getUrl()'s "view" page, which does not work as an <img src> — using
// the uc?export=view form so it renders inline for images and embeds.
function uploadDataUrlToDrive_(payload, filenamePrefix) {
  const match = /^data:([^;]+);base64,(.+)$/.exec(payload.data);
  if (!match) throw new Error('uploadDataUrlToDrive_: payload.data is not a base64 data URL.');
  const mimeType = match[1];
  const base64 = match[2];
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType, filenamePrefix + '_' + (payload.name || 'file'));

  const folder = getAttachmentsFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}
