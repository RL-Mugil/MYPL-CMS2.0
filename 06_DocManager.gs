/**
 * ============================================================
 * 06_DocumentManager.gs — Drive File Operations
 * ============================================================
 */

/**
 * Get documents from a client's folder, organized by subfolder
 */
function getClientDocuments(clientId) {
  var client = getClientById(clientId);
  if (!client || !client.CLIENT_FOLDER_ID) {
    return { error: "Client folder not found" };
  }

  try {
    var clientFolder = DriveApp.getFolderById(client.CLIENT_FOLDER_ID);
    var result = {};

    CONFIG.CLIENT_SUBFOLDERS.forEach(function(subName) {
      result[subName] = [];
      var subFolders = clientFolder.getFoldersByName(subName);
      if (subFolders.hasNext()) {
        var files = subFolders.next().getFiles();
        while (files.hasNext()) {
          var f = files.next();
          result[subName].push({
            name: f.getName(),
            url: f.getUrl(),
            size: f.getSize(),
            mimeType: f.getMimeType(),
            dateCreated: f.getDateCreated(),
            lastUpdated: f.getLastUpdated()
          });
        }
      }
    });

    return result;
  } catch (e) {
    return { error: "Error accessing files: " + e.message };
  }
}

/**
 * Upload a file to a client's subfolder (called from admin UI)
 */
function uploadToClientFolder(clientId, subfolderName, fileBlob, fileName) {
  var client = getClientById(clientId);
  if (!client || !client.CLIENT_FOLDER_ID) {
    return { success: false, message: "Client folder not found" };
  }

  try {
    var clientFolder = DriveApp.getFolderById(client.CLIENT_FOLDER_ID);
    var subFolders = clientFolder.getFoldersByName(subfolderName);

    var targetFolder;
    if (subFolders.hasNext()) {
      targetFolder = subFolders.next();
    } else {
      targetFolder = clientFolder.createFolder(subfolderName);
    }

    var file = targetFolder.createFile(fileBlob);
    if (fileName) file.setName(fileName);

    logActivity_("UPLOAD_FILE", "DOCUMENT", clientId, "File: " + file.getName() + " -> " + subfolderName);

    return {
      success: true,
      fileUrl: file.getUrl(),
      message: "File uploaded to " + subfolderName
    };
  } catch (e) {
    return { success: false, message: "Upload error: " + e.message };
  }
}

/**
 * Get case-specific documents from the case folder
 */
function getCaseDocuments(caseId) {
  var sheet = getSheet_("CASE");
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === caseId) {
      var folderId = data[i][10]; // CASE_FOLDER_ID
      if (!folderId) return { error: "Case folder not found" };

      try {
        var folder = DriveApp.getFolderById(folderId);
        var result = {};
        var subFolders = folder.getFolders();

        while (subFolders.hasNext()) {
          var sub = subFolders.next();
          var subName = sub.getName();
          result[subName] = [];
          var files = sub.getFiles();
          while (files.hasNext()) {
            var f = files.next();
            result[subName].push({
              name: f.getName(),
              url: f.getUrl(),
              size: f.getSize(),
              dateCreated: f.getDateCreated()
            });
          }
        }
        return result;
      } catch (e) {
        return { error: "Error accessing case folder: " + e.message };
      }
    }
  }
  return { error: "Case not found" };
}
