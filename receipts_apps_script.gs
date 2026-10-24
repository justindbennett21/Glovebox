/**
 * Glovebox receipts helper.
 * Runs as YOU, so receipts are saved in your own Google Drive
 * (folder "Glovebox Receipts", created automatically).
 *
 * Setup: paste into script.google.com, set TOKEN, run `setup` once,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */

// Make up a long random phrase. Paste the SAME value into Streamlit secrets as receipts_token.
const TOKEN = 'CHANGE-ME-to-a-long-random-phrase';
const FOLDER_NAME = 'Glovebox Receipts';

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== TOKEN) return reply({ ok: false, error: 'Not allowed' });
    const folder = getFolder();

    if (body.action === 'upload') {
      const bytes = Utilities.base64Decode(body.data);
      const blob = Utilities.newBlob(bytes, body.mimeType || 'application/octet-stream', body.filename || 'receipt');
      const file = folder.createFile(blob);
      return reply({ ok: true, id: file.getId(), url: file.getUrl() });
    }

    if (body.action === 'trash') {
      const file = DriveApp.getFileById(body.id);
      const parents = file.getParents();
      while (parents.hasNext()) {
        if (parents.next().getId() === folder.getId()) {
          file.setTrashed(true);
          return reply({ ok: true });
        }
      }
      return reply({ ok: false, error: 'File is not in the receipts folder' });
    }

    return reply({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  }
}

function getFolder() {
  const found = DriveApp.getFoldersByName(FOLDER_NAME);
  return found.hasNext() ? found.next() : DriveApp.createFolder(FOLDER_NAME);
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Run this once from the editor to grant Drive permission and create the folder.
function setup() {
  Logger.log('Receipts folder: ' + getFolder().getUrl());
}
