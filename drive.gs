/**
 * Google Drive Control for COB-CHA
 */

/**
 * Get Spreadsheet's folder
 * @return Object
 */
function getCurrentFolder() {
  if (getCurrentFolder.folder) return getCurrentFolder.folder;
  var ss = getSpreadSheet();
  var ssId = ss.getId();
  var parentFolder = DriveApp.getFileById(ssId).getParents();
  getCurrentFolder.folder = parentFolder.next();
  return getCurrentFolder.folder; 
};

/**
 * Get target folder object
 * @param String target [resourceFolderName, exportFolderName, imagesFolderName]
 * @return Object
 */
function getTargetFolder(target) {
  if (getTargetFolder.folder && getTargetFolder.folder[target]) return getTargetFolder.folder[target];
  var currentFolder = getCurrentFolder();
  var children = currentFolder.getFolders();
  getTargetFolder.folder = {};
  
  // is already exists?
  var folders = [resourceFolderName, exportFolderName, imagesFolderName]
  while (children.hasNext()){
    var folder = children.next();
    for (var i = 0; i < folders.length; i++) {
      if (folder.getName().indexOf(folders[i]) != -1) {
        getTargetFolder.folder[folders[i]] = folder;
      }
    }
    if (getTargetFolder.folder[target]) return getTargetFolder.folder[target];
  }
  
  // create folder
  for (var i = 0; i < folders.length; i++) {
    getTargetFolder.folder[folders[i]] = currentFolder.createFolder(folders[i]);
  }
  return getTargetFolder.folder[target];
};

/**
 * delete file if exists
 * @param String target
 * @param String name
 * @return Void
 */
function deleteFileIfExists(targetFolder, name) {
  var targetFolder = getTargetFolder(targetFolder);
  var children = targetFolder.getFiles();
  while (children.hasNext()) {
    var current = children.next();
    if (current.getName() == name) {
      targetFolder.removeFile(current);
    }
  };
};

/**
 * save HTML
 * @param String target
 * @param String name
 * @param String html
 * @param Bool overwrite
 * @return Void
 */
function saveHtml(targetFolder, name, html, overwrite) {
  if (overwrite) {
    deleteFileIfExists(targetFolder, name);
  }
  var targetFolder = getTargetFolder(targetFolder);
  targetFolder.createFile(name, html, 'text/html');
};

/**
 * file upload
 * @param String targetFolder
 * @param Object formObj
 * @param String nameAttr
 * @return Array [fileName, fileId]
 */
function fileUpload(targetFolder, formObj, nameAttr) {
  if (formObj[nameAttr].length == 0) throw new Error('Empty File Uploaded');
  var formBlob = formObj[nameAttr];
  var driveFile = DriveApp.createFile(formBlob);
  var targetFolder = getTargetFolder(targetFolder);
  deleteFileIfExists(targetFolder, driveFile.getName());
  targetFolder.addFile(driveFile);
  driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  DriveApp.getRootFolder().removeFile(driveFile);
  return [driveFile.getName() ,driveFile.getId()];
}

/**
 * add screenshot
 * @param Object formObj
 * @return String
 */
function uploadScreenshot(formObj) {
  var activeSheet = getActiveSheet();
  if (activeSheet.getName().charAt(0) == '*') return getUiLang('current-sheet-is-not-for-webpage', "Current Sheet is not for webpage.");
  
  var file = fileUpload(imagesFolderName, formObj, "imageFile");
  activeSheet.getRange(2, 6).setValue(file[0]);
  activeSheet.getRange(2, 7).setValue('=IMAGE("https://drive.google.com/uc?export=download&id='+file[1]+'",1)');
  
  return getUiLang('image-uploaded', "Image Uploaded.");
}
