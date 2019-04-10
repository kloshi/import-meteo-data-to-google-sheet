# import-meteo-data-to-google-sheet
Fetch weather data from DWD.de and save it in the Google Sheet

## intro
This script is used to fetch meteo data from DWD.de and save (update) existing data in a GoogleSheet spreadsheet. 

## installation
Open a Google Spreadsheet, go to `Tools/Script editor` and paste the code over there.

Install a menu in the spreadsheet by executing the function: `Run/Run function/OnOpen`

## usage
To fetch new meteo data, run the function `updateMonthlyMeteoData` manually from within the spreadsheet. Setting up a periodic, automatic trigger in GoogleSheets is also possible.

Set environmental variables:

    RESOURCE_URL = url of the resource at DWD.de
    ARCHIVE_SAVE_FOLDER_ID = GoogleDrive folder id, where temp files will be stored (zip archive and 1 extracted data file) 
    TEMPERATURE_FILE_NAME_PATTERN = beginning of the data file filename 
    SPREADSHEET_ID = GoogleSheet spreadsheet id, where the meteo data is stored

