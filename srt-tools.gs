function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SRT Tools')
    .addItem('Show Upload Instructions', 'showUploadInstructions')
    .addItem('Process Uploaded SRTs', 'processUploadedSRTs')
    .addItem('Export Active Sheet as SRT', 'exportActiveSheetAsSRT')
    .addToUi();
}

function setupFolders() {
  ['SRT', 'SRT/Process', 'SRT/Process/Processed', 'SRT/Exports'].forEach(function(folderPath) {
    createFolderIfNotExist(folderPath);
  });
}

function createFolderIfNotExist(folderPath) {
  var folders = folderPath.split('/');
  var currentFolder = DriveApp.getRootFolder();
  
  for (var i = 0; i < folders.length; i++) {
    var folderName = folders[i];
    var nextFolders = currentFolder.getFoldersByName(folderName);
    if (!nextFolders.hasNext()) {
      currentFolder = currentFolder.createFolder(folderName);
    } else {
      currentFolder = nextFolders.next();
    }
  }
  return currentFolder;
}

function showUploadInstructions() {
  setupFolders(); // Ensure folders are set up before showing instructions
  var processFolder = createFolderIfNotExist('SRT/Process');
  var html = HtmlService.createHtmlOutput(`
    <p>Please upload your SRT files to the following Google Drive folder:</p>
    <a href="${processFolder.getUrl()}" target="_blank" rel="noopener noreferrer">Upload SRT Files Here</a>
    <p>After uploading, return to this sheet and use the 'Process Uploaded SRTs' option from the SRT Tools menu.</p>
  `)
  .setWidth(400)
  .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload SRT Files');
}

function processUploadedSRTs() {
  setupFolders(); // Ensure folders are set up before processing
  var processFolder = createFolderIfNotExist('SRT/Process');
  var processedFolder = createFolderIfNotExist('SRT/Process/Processed');
  var files = processFolder.getFiles();
  
  while (files.hasNext()) {
    var file = files.next();
    var content = file.getBlob().getDataAsString();
    var lines = content.split('\n');
    var data = [];
    for (var i = 0; i < lines.length;) {
      if (lines[i].trim() === '') { i++; continue; } // Skip empty lines
      var number = lines[i++];
      var timestamp = lines[i++];
      var text = '';
      while (lines[i] && lines[i].trim() !== '') {
        text += lines[i++] + '\n';
      }
      i++; // Skip the blank line after text
      data.push([number, timestamp, text.trim()]);
    }
    var sheetName = file.getName().replace('.srt', '');
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
    sheet.clear(); // Clear any existing content
    sheet.getRange(1, 1, data.length, 3).setValues(data);
    
    // Move the processed file to the 'Processed' folder
    file.moveTo(processedFolder);
  }
}

function exportActiveSheetAsSRT() {
  setupFolders(); // Ensure folders are set up before exporting
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var srtContent = data.map(function(row) {
    return `${row[0]}\n${row[1]}\n${row[2]}`;
  }).join('\n\n');
  
  var exportFolder = createFolderIfNotExist('SRT/Exports');
  var fileName = `${sheet.getName()}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.srt`;
  var file = exportFolder.createFile(fileName, srtContent, MimeType.PLAIN_TEXT);
  
  var html = HtmlService.createHtmlOutput(`
    <p>Your file has been processed. You can download it from the link below:</p>
    <a href="${file.getUrl()}" target="_blank" rel="noopener noreferrer">${fileName}</a>
  `).setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Download Processed SRT File');
}
