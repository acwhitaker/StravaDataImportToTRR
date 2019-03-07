var CLIENT_ID = '[your_client_id]';
var CLIENT_SECRET = '[your_client_secret]';
var SPREADSHEET_ID = '[your_spreadsheet_id]';
var TRAINING_LOG_SHEET_NAME = "Training Log";
var CONFIG_SHEET_NAME = "Config";
var DEBUG = false;
var AFTER_DATE_CELL = 'B4';
var BEFORE_DATE_CELL = 'B5';


/**
 * Configures the service.
 */
function getService() {
  return OAuth2.createService('Strava')
      // Set the endpoint URLs.
      .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
      .setTokenUrl('https://www.strava.com/oauth/token')

      // Set the client ID and secret.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the name of the callback function that should be invoked to complete
      // the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())
  
      //Include private activities when retrieving activities.
      .setScope ("view_private")
  
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied');
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  var service = getService();
  service.reset();
}


/**
 * Authorizes and makes a request to the Strava API.
 */
function run() {
  var service = getService();
  if (service.hasAccess()) {
    var url = 'https://www.strava.com/api/v3/athlete';
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      }
    });
    var result = JSON.parse(response.getContentText());
    Logger.log(JSON.stringify(result, null, 2));
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
}

function retrieveData() {
  //if sheet is empty retrieve all data
  var service = getService();
  if (service.hasAccess()) {
    var testImportSheet = getTestImportSheet();
    var trainingLogSheet = getTrainingLogSheet();
    var unixTimeAfter = retrieveConfigUnixTime('after');
    var unixTimeBefore = retrieveConfigUnixTime('before');
    
    var url = 'https://www.strava.com/api/v3/athlete/activities?per_page=200' + '&before=' + unixTimeBefore + '&after=' + unixTimeAfter;
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      }
    });
    
    var result = JSON.parse(response.getContentText());

    if (result.length == 0) {
      Logger.log("No new data");
      return;
    }
    
    var data = convertData(result);
    
    if (data.length == 0) {
      Logger.log("No new data with heart rate");
      return;
    }
    
    logData(testImportSheet, data);
    insertDataToTRRTrainingLog(trainingLogSheet,data);
   
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
}

function retrieveLastDate(sheet) {
  var lastRow = sheet.getLastRow();
  var unixTime = 0; 
  if (lastRow > 0) { 
      var dateCell = sheet.getRange(lastRow, 1);
      var dateString = dateCell.getValue();
      Logger.log((dateString || "").replace(/-/g,"/").replace(/[TZ]/g," "));
      var date = new Date((dateString || "").replace(/-/g,"/").replace(/[TZ]/g," "));
      unixTime = date/1000;
   }
   return unixTime;
}

function retrieveConfigUnixTime(param) {
  var sheet = getConfigSheet();
  var unixTime = 0;
  var dateString = '';
  var dateToImport;
  
  switch(param) {
    case 'after':
      dateString = sheet.getRange(AFTER_DATE_CELL).getValue();
      dateToImport = new Date(dateString);
      unixTime = dateToImport/1000;
      break;
    case 'before':
      dateString = sheet.getRange(BEFORE_DATE_CELL).getValue();
      dateToImport = new Date(dateString); 
      unixTime = dateToImport/1000;
      //dateToImport = dateToImport + (24 * 60 * 60 * 1000); //add one day in milliseconds
      unixTime = unixTime + (24 * 60 * 60);
      break;
  }
  

  
  // Convert the date from GMT to user's time zone.
  //var timezoneOffsetMilliseconds = dateToImportAfter.getTimezoneOffset() * 60 * 1000;  //getTimezoneOffset() returns minutes
  //var millisecondsToFind = dateToImportAfter.getTime() + timezoneOffsetMilliseconds;
  //dateToImportAfter = new Date(millisecondsToFind); //assign dateToFind to a date variable created using the milliseconds corrected to reflect the user's time zone
  
  return unixTime;
}


function convertData(result) {
  var data = [];
  
  for (var i = 0; i < result.length; i++) {
    if (result[i]["type"] == "Run") {
    //if (1) {
      var item = [result[i]['start_date_local'],
                  roundToTwoPlaces(result[i]['distance']/1609.34),
                  result[i]['type'],
                  roundToTwoPlaces((result[i]['moving_time']/60)),
                  Math.round(result[i]['total_elevation_gain']*3.281),
                  result[i]['id'],
                  'https://www.strava.com/activities/' + result[i]['id'],
                 ];
      data.push(item);
    }
      
  }
  
  return data;
}

function getTrainingLogSheet() {
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(spreadsheet, TRAINING_LOG_SHEET_NAME);
  return sheet;
}

function getConfigSheet() {
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(spreadsheet, CONFIG_SHEET_NAME);
  return sheet;
}

function getTestImportSheet() {
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(spreadsheet, "StravaImportLog");
  return sheet;
}

function logData(sheet, data) {
  var header = ["Date", "Distance","Type", "Time", "Vert", "Activity ID", "Activity URL"];
  sheet.clearContents();
  ensureHeader(header, sheet);
  
  
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(lastRow+1,1,data.length,7);
  range.setValues(data); 
}

function insertDataToTRRTrainingLog(sheet, data) {
  
  var trainingLogRangeRow = 0;
  var dateString = '';
  var date = new Date();
  var subjectiveFeedbackStravaTest = false;
  var subjectiveFeedbackStravaTestString = '';
  var stravaActivityTime, stravaActivityDistance, stravaActivityVert, stravaActivityURL;
  //for each activity i in the imported run data
  Logger.log("length of data is %s", data.length);
  for (var i = 0; i < data.length; i++) {
    //find the row in the Training Log spreadsheet which matches the date of the imported run
    dateString = data[i][0];
    Logger.log(dateString);
    
    Logger.log((dateString || "").replace(/-/g,"/").replace(/[TZ]/g," "));
    date = new Date((dateString || "").replace(/-/g,"/").replace(/[TZ]/g," "));
    Logger.log(date);
    trainingLogRangeRow = getRowFromDate(date, sheet);
    Logger.log(trainingLogRangeRow);
    var trainingLogRunnerRange = sheet.getRange('O'+trainingLogRangeRow+':V'+trainingLogRangeRow);
    
    //put the strava data fields in human readable variable names
    stravaActivityTime = data[i][3];
    stravaActivityDistance = data[i][1];
    stravaActivityVert = data[i][4];
    stravaActivityURL = data[i][6];
    
    //set up a test variable to check whether a row's Subjective Feedback cell has a previous strava import
    subjectiveFeedbackStravaTestString = trainingLogRunnerRange.getCell(1,6).getValue();
    subjectiveFeedbackStravaTest = (subjectiveFeedbackStravaTestString.indexOf("https://www.strava.com/activities/") > -1);
    //if ALL the cells are blank, then copy the data into the relevant cells
    if (trainingLogRunnerRange.isBlank()){ //prevent overwriting data by checking that the range is blank before adding responses to the Triaining log
      
      //assign the data from the strava import to the proper column in the Training Log sheet
      trainingLogRunnerRange.getCell(1,2).setValue(stravaActivityTime);
      trainingLogRunnerRange.getCell(1,3).setValue(stravaActivityDistance);
      trainingLogRunnerRange.getCell(1,4).setValue(stravaActivityVert);
      trainingLogRunnerRange.getCell(1,6).setValue(stravaActivityURL);
      
    } else if (subjectiveFeedbackStravaTest) { //the row contains a run previously imported from strava; accumulate the run totals if unique
      
      if (!(trainingLogRunnerRange.getCell(1,6).getValue().indexOf(stravaActivityURL) > -1)) { //check for a duplicate strava activity
        
        trainingLogRunnerRange.getCell(1,2).setValue(trainingLogRunnerRange.getCell(1,2).getValue() + stravaActivityTime);
        trainingLogRunnerRange.getCell(1,3).setValue(trainingLogRunnerRange.getCell(1,3).getValue() + stravaActivityDistance);
        trainingLogRunnerRange.getCell(1,4).setValue(trainingLogRunnerRange.getCell(1,4).getValue() + stravaActivityVert);
        trainingLogRunnerRange.getCell(1,6).setValue(trainingLogRunnerRange.getCell(1,6).getValue() + '\n' + stravaActivityURL);
      }
      
    }
    else {
      //Ideally would show a pop-up to the user to alert that the spreadsheet entry did not work, but GAS does not make it possible to show an alert to the user of a form, just the editor.
      
      Logger.log("The cells in row %s (date: %s) of the Training Log spreadsheet are not blank. The form data was not entered in the spreadsheet.", trainingLogRangeRow, dateString);
      
    }
  }
  
}

function clearSheet(sheet) {
  sheet.clearContents();
}

function ensureHeader(header, sheet) {
  // Only add the header if sheet is empty
  if (sheet.getLastRow() == 0) {
    if (DEBUG) Logger.log('Sheet is empty, adding header.')    
    sheet.appendRow(header);
    return true;
    
  } else {
    if (DEBUG) Logger.log('Sheet is not empty, not adding header.')
    return false;
  }
}


function getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    if (DEBUG) Logger.log('Sheet "%s" does not exists, adding new one.', sheetName);
    sheet = spreadsheet.insertSheet(sheetName)
  } 
  
  return sheet;
}

function getRowFromDate(dateToFind, s){
  
  //finds a row in the Training Log sheet matching the date provided. The input spreadsheet ss should be a TeamRunRun Training Log spreadsheet.
  
  //var ss = SpreadsheetApp.openById(FormApp.getActiveForm().getDestinationId()); //ok to use if spreadsheet will be linked to the form responses
  //var s = ss.getSheetByName('Training Log'); // Must be changed if the name of the sheet changes
  var startRow = 3; // The first SS row containing dates and plan/log data. Should be changed if rows are inserted or deleted from the header.
  var values = s.getRange('A'+startRow+':B').getValues();
  var endRow = s.getMaxRows();
  var row;
  var currentRowDate;
  var dateToFindYYYYDD;
  var currentRowDateYYYYDD;
  var dateToFindYYYY;
  var currentRowDateYYYY;
  var currentRowDateDD;
  var dateToFindDD;
  var datesMatch = false;
  var dateToFindRow;
  

  //iterate over rows to find the first row in the values spreadsheet that corresponds with dateToFind
  row = 0; //this row variable represents rows of the values[][] array, not spreadsheet rows. values[0]0] corresponds to the startRow-th row of the spreadsheet.
  do {
    
    //Get some different date number formats for both dateToFind and the date in the current row to help with an efficient search.
    currentRowDate = values[row][0];
    currentRowDateYYYYDD = Utilities.formatDate(currentRowDate,"PST","yyyyDD");
    dateToFindYYYYDD = Utilities.formatDate(dateToFind,"PST","yyyyDD");
    currentRowDateYYYY = Utilities.formatDate(currentRowDate,"PST","yyyy");
    dateToFindYYYY = Utilities.formatDate(dateToFind,"PST","yyyy");
    currentRowDateDD = Utilities.formatDate(currentRowDate,"PST","DD");
    dateToFindDD = Utilities.formatDate(dateToFind,"PST","DD");
    
    
    if (currentRowDateYYYYDD == dateToFindYYYYDD) {
      //Success! dateToFind's date is in this row.
      datesMatch = true;
      dateToFindRow = row;
    }
    else if (currentRowDateYYYY > dateToFindYYYY) { 
      //error condition, dateToFind should always be after the date in the array
      //Browser.msgBox("Error, see log.");
      Logger.log("dateToFind should always be after the date in the array dateYYYY")
      return;
    }
    else if (currentRowDateYYYY < dateToFindYYYY) { 
      //the date in the current row is before dateToFind, so
      //need to jump ahead in the array to the first day of the next year by adding to the current row
      
      if (( (currentRowDateYYYY % 4) == 0) && (currentRowDateYYYY % 100 == 0) || (currentRowDateYYYY % 400) == 0) { 
        // is leap year
        //jump ahead in the array by the number of days left in the year
        row = row + 366-currentRowDateDD+1; 
      }
      else { 
        //not leap year
        //jump ahead in the array by the number of days left in the year
        row = row + 365-currentRowDateDD+1; 
      }      
    }
    else if (currentRowDateDD > dateToFindDD) { 
      //error condition, dateToFind should always be after the date in the array
      Browser.msgBox("Error, see log.");
      Logger.log("dateToFind should always be after the date in the array dateDD")
      return;
    }
    else if (currentRowDateDD < dateToFindDD) {
      row = row + (dateToFindDD - currentRowDateDD);
    }
    else { //error condition
      Browser.msgBox("Error, see log.");
      Logger.log("error during date search")
      return;
    }
    
  } while (datesMatch == false);  
  
  return dateToFindRow + startRow;
}

function roundToTwoPlaces(number) {
  return Math.floor(number*100)/100;
}
