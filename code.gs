// === Fetch weather data from DWD.de and save it (update existing) in the Google Sheet ===

var RESOURCE_MONTHLY_URL = 'https://www.dwd.de/DE/leistungen/_config/leistungsteckbriefPublication.zip?view=nasPublication&nn=16102&imageFilePath=157242051950877752011408908330930139598321949458353665338080907323067407358570939094063695252196106887026681320814526060536595251033063153662108661981192575715848535028223489905255018373803265100435386215906807083067009139783338575081559719801855679129507120696592788599&download=true',
    RESOURCE_DAILY_URL = 'https://www.dwd.de/DE/leistungen/_config/leistungsteckbriefPublication.zip?view=nasPublication&nn=495662&imageFilePath=157242051950877752011408908330930139598321949458353665338080907323067407358570939094063695252196106887026681320814526060536595251033063153662108661981192575715848535028223489905255018373803265100435386215906807083067009139783338572007248601377183425875240168840296754295&download=true',
    ARCHIVE_SAVE_FOLDER_ID = '11cQZ9vc5jJDUaEG2_QI9_c35aYQgIHpw',
    TEMPERATURE_MONTHLY_FILENAME_PATTERN = 'produkt_klima_monat',
    TEMPERATURE_DAILY_FILENAME_PATTERN = 'produkt_klima_tag',    
    SPREADSHEET_ID = '16joYgkhMi6a4YGKQ1adNT2du7AIFSCkMiQS4ewe8vKs',
    DAILY_UPDATE_ROW_POSITION = 8

var ss = SpreadsheetApp.openById(SPREADSHEET_ID),
    meteoMonthlySheet = ss.getSheetByName("Meteo Cottbus"),
    meteoDailySheet = ss.getSheetByName("Meteo Cottbus daily")

var lastNonEmptyRow = lastRowForColumn(meteoMonthlySheet, 1),
    lastUpdatedMonth = meteoMonthlySheet.getRange(lastNonEmptyRow, 2).getValue()

function updateMeteoData() {
  updateMonthlyMeteoData();
  updateDailyMeteoData();
}

// ======================================================================


function updateMonthlyMeteoData() {
  var zipFile = fetchAndSaveArchive(RESOURCE_MONTHLY_URL, ARCHIVE_SAVE_FOLDER_ID);
  var updateFile = unzipTemperatureFile(TEMPERATURE_MONTHLY_FILENAME_PATTERN, ARCHIVE_SAVE_FOLDER_ID, zipFile.getName());

  var newRows = getNewMonthlyRows(updateFile, lastUpdatedMonth); 
  if (newRows.length > 0) { updateSheet(meteoMonthlySheet, newRows, lastNonEmptyRow + 1); }

  zipFile.setTrashed(true);
  updateFile.setTrashed(true);
}


function updateDailyMeteoData() {
  var zipFile = fetchAndSaveArchive(RESOURCE_DAILY_URL, ARCHIVE_SAVE_FOLDER_ID);
  var updateFile = unzipTemperatureFile(TEMPERATURE_DAILY_FILENAME_PATTERN, ARCHIVE_SAVE_FOLDER_ID, zipFile.getName());

  var newRows = getNewDailyRows(updateFile);
  updateSheet(meteoDailySheet, newRows, DAILY_UPDATE_ROW_POSITION);

  zipFile.setTrashed(true);
  updateFile.setTrashed(true);
}


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


// parse the txt and return an array of row(s) ready to paste into the sheet
function getNewMonthlyRows(file, lastUpdate) {
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


function getNewDailyRows(file) {
  var csv = file.getBlob().getDataAsString();
  var csvData = Utilities.parseCsv(csv);
  var result = [];
  var i, row;
  // skip 1st row (labels) (i=1 instead of 0)
  for (i = 1; i < csvData.length; i++) {
    row = String(csvData[i]).split(";");
    result.push(row);
  }  
  return result;
}


function updateSheet(sheet, rows, position) {
  sheet.getRange(position, 1, rows.length, rows[0].length).setValues(rows);
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
      .addItem('Fetch meteo data from DWD.de','updateMeteoData')
      .addItem('Fetch 5-day forecast','fetchForecast')
      .addToUi();
}



// ========================================

function fetchForecast() {
  // Call the OpenWeatherMap API for Cottbus data
  var response = UrlFetchApp.fetch("api.openweathermap.org/data/2.5/forecast?id=2939811&units=metric&APPID=103b510d36ec31f578d575d938d65581");
  // Call the OpenWeatherMap API for Hoyerswerda data
//  var response = UrlFetchApp.fetch("api.openweathermap.org/data/2.5/forecast?id=2898304&units=metric&APPID=103b510d36ec31f578d575d938d65581");
  var data = JSON.parse(response);

  // extract only the 'list' hash
  var data = data["list"];
    
  // set the right spreadsheet and sheet
  var ss = SpreadsheetApp.openById("16joYgkhMi6a4YGKQ1adNT2du7AIFSCkMiQS4ewe8vKs");
  SpreadsheetApp.setActiveSpreadsheet(ss); 
  var meteoMonthlySheet = ss.getSheetByName("Meteo Cottbus");
  ss.setActiveSheet(meteoMonthlySheet);

  var output = [];
  
  data.forEach(function(elem,i) {
    date = new Date(elem["dt"]*1000); // convert date-time from Unix format
    output.push([date, elem["main"]["temp"]]); // one row has 2 columns (date, temp)
    Logger.log(date + " " + elem["main"]["temp"]); // just logging
  });
  
  // paste into the sheet
  meteoMonthlySheet.getRange(8, 20, output.length, 2).setValues(output);
}

//===================
