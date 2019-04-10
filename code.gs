// === Fetch weather data from DWD.de and save it in the Google Sheet ===

var RESOURCE_URL = 'https://www.dwd.de/DE/leistungen/_config/leistungsteckbriefPublication.zip?view=nasPublication&nn=16102&imageFilePath=157242051950877752011408908330930139598321949458353665338080907323067407358570939094063695252196106887026681320814526060536595251033063153662108661981192575715848535028223489905255018373803265100435386215906807083067009139783338575081559719801855679129507120696592788599&download=true',
    ARCHIVE_SAVE_FOLDER_ID = '11cQZ9vc5jJDUaEG2_QI9_c35aYQgIHpw',
    TEMPERATURE_FILE_NAME_PATTERN = 'produkt_klima_monat_',
    SPREADSHEET_ID = '16joYgkhMi6a4YGKQ1adNT2du7AIFSCkMiQS4ewe8vKs'

var ss = SpreadsheetApp.openById(SPREADSHEET_ID),
    meteoSheet = ss.getSheetByName("Meteo Cottbus")

var lastNonEmptyRow = lastRowForColumn(meteoSheet, 1),
    lastUpdatedMonth = meteoSheet.getRange(lastNonEmptyRow, 2).getValue()


function updateMonthlyMeteoData() {
  var zipFile = fetchAndSaveArchive(RESOURCE_URL, ARCHIVE_SAVE_FOLDER_ID);
  var updateFile = unzipTemperatureFile(TEMPERATURE_FILE_NAME_PATTERN, ARCHIVE_SAVE_FOLDER_ID, zipFile.getName());

  var newRows = getNewRows(updateFile, lastUpdatedMonth); 
  if (newRows.length > 0) { updateSheet(meteoSheet, newRows); }

  zipFile.setTrashed(true);
  updateFile.setTrashed(true);
}

// ======================================================================


function fetchAndSaveArchive(url, folderID) {
  var file = UrlFetchApp.fetch(url).getBlob();
  var folder = DriveApp.getFolderById(folderID);
  return folder.createFile(file);
}


function unzipTemperatureFile(filenamePattern, folderID, fileZipName) {
  var folder = DriveApp.getFolderById(folderID);
  var fileZip = folder.getFilesByName(fileZipName);
  var fileExtractedBlob, fileZipBlob, i;

  // Decompression
  fileZipBlob = fileZip.next().getBlob();
  fileZipBlob.setContentType("application/zip");
  fileExtractedBlob = Utilities.unzip(fileZipBlob);
  for (i=0; i < fileExtractedBlob.length; i++){
    filename = fileExtractedBlob[i].getName();
    if (filename.indexOf(filenamePattern) > -1) { 
      return folder.createFile(fileExtractedBlob[i]); 
    }
  }
}


// parse the txt and return an array or row(s) (ready to be pasted into the sheet)
function getNewRows(file, lastUpdate) {
  var csv = file.getBlob().getDataAsString();
  var csvData = Utilities.parseCsv(csv);

  // build array of new rows (newer than the one existing in the file)
  result = [];
  csvData.forEach(function(row, i) {
    row = String(row).split(";");
    if (row[1] > lastUpdate) { result.push(row); }
  });
  return result;
}


function updateSheet(sheet, rows) {
  sheet.getRange(lastNonEmptyRow + 1, 1, rows.length, rows[0].length).setValues(rows);
}


// helper function: returns the last row for the specific column
function lastRowForColumn(sheet, column){
  // Get the last row with data for the whole sheet.
  var numRows = sheet.getLastRow();
  
  // Get all data for the given column
  var data = sheet.getRange(1, column, numRows).getValues();
  
  // Iterate backwards and find first non empty cell
  for(var i = data.length - 1 ; i >= 0 ; i--){
    if (data[i][0] != null && data[i][0] != ""){
      return i + 1;
    }
  }
}


// add function to to menu in the spreadsheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Weather data')
      .addItem('Fetch meteo data from DWD.de','updateMonthlyMeteoData')
      .addToUi();
}



// ========================================
