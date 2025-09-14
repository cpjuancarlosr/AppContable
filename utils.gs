//== File: utils.gs (validaciones, fechas, bitácora simple)

/**
 * A collection of utility functions for validation, date manipulation, and logging.
 */

function log_(message) {
  try {
    const ss = SpreadsheetApp.getActive();
    let logSheet = ss.getSheetByName("Logs");
    if (!logSheet) {
      logSheet = ss.insertSheet("Logs");
      logSheet.getRange("A1:B1").setValues([["Timestamp", "Message"]]).setFontWeight("bold");
    }
    logSheet.appendRow([new Date(), message]);
  } catch (e) {
    Logger.log(`Could not write to Logs sheet: ${e.message}. Fallback log: ${message}`);
  }
}

/**
 * Recursively collects all XML files from a given Drive folder and its subfolders.
 * @param {GoogleAppsScript.Drive.Folder} folder The root folder to start crawling from.
 * @return {Array<GoogleAppsScript.Drive.File>} An array of file objects.
 */
function recolectarXMLs_(folder) {
  const out = [];
  const crawl = (fol) => {
    const files = fol.getFiles();
    while (files.hasNext()) {
      let file = files.next();
      try {
        file = resolveShortcutFile_(file);
      } catch (e) {
        // Ignore if shortcut can't be resolved
      }
      const mt = (file.getMimeType() || "").toLowerCase();
      if (mt.includes('xml') || file.getName().toLowerCase().endsWith('.xml')) {
        out.push(file);
      } else if (mt.includes('zip')) {
        try {
          const blobs = Utilities.unzip(file.getBlob());
          blobs.forEach(b => {
            if (b.getName().toLowerCase().endsWith('.xml')) {
              // Note: This creates a temporary file object. It's not ideal but works.
              // For a more robust solution, we'd handle blobs directly.
              out.push({
                getName: () => b.getName(),
                getText: (enc) => b.getDataAsString(enc || 'UTF-8')
              });
            }
          });
        } catch (e) {
          log_(`Could not unzip file: ${file.getName()}. Error: ${e.message}`);
        }
      }
    }
    const subs = fol.getFolders();
    while (subs.hasNext()) {
      crawl(subs.next());
    }
  };
  crawl(folder);
  return out;
}


/**
 * If the given file is a Drive Shortcut, returns the target file. Otherwise, returns the file itself.
 * Requires the Advanced Drive Service to be enabled.
 * @param {GoogleAppsScript.Drive.File} file The file to resolve.
 * @return {GoogleAppsScript.Drive.File} The resolved file.
 */
function resolveShortcutFile_(file) {
  const MIME_SHORTCUT = 'application/vnd.google-apps.shortcut';
  try {
    if (file.getMimeType() !== MIME_SHORTCUT) {
      return file;
    }
    const targetId = Drive.Files.get(file.getId()).shortcutDetails.targetId;
    if (targetId) {
      return DriveApp.getFileById(targetId);
    }
  } catch (e) {
    // If the advanced service is not enabled or fails, return the original file.
  }
  return file;
}
