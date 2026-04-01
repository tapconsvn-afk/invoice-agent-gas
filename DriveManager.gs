function saveAttachmentToMonthlyFolder_(attachment, emailDate) {
  var rootFolder = getOrCreateRootFolder_();
  var safeDate = ensureDate_(emailDate);

  var monthFolderName = Utilities.formatDate(
    safeDate,
    Session.getScriptTimeZone(),
    'yyyy-MM'
  );

  var monthFolder = getOrCreateChildFolder_(rootFolder, monthFolderName);

  var originalName = safeString_(attachment.getName());
  var finalName = buildSafePdfFileName_(originalName, safeDate);

  var blob = attachment.copyBlob().setName(finalName);
  var file = monthFolder.createFile(blob);

  Logger.log('saved_file=' + file.getName());
  return file;
}

function getOrCreateRootFolder_() {
  var folders = DriveApp.getFoldersByName(CONFIG.DRIVE_ROOT_FOLDER);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(CONFIG.DRIVE_ROOT_FOLDER);
}

function getOrCreateChildFolder_(parentFolder, childFolderName) {
  var folders = parentFolder.getFoldersByName(childFolderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(childFolderName);
}

function buildSafePdfFileName_(originalName, fileDate) {
  var safeDate = ensureDate_(fileDate);
  var name = safeString_(originalName);

  if (!name) {
    name = 'invoice.pdf';
  }

  if (name.toLowerCase().slice(-4) !== '.pdf') {
    name += '.pdf';
  }

  var stamp = Utilities.formatDate(
    safeDate,
    Session.getScriptTimeZone(),
    'yyyyMMdd_HHmmss'
  );

  name = name.replace(/[\\\/:*?"<>|#\[\]]/g, '_');

  return stamp + '_' + name;
}
