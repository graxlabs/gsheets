const APIName = 'https://EXAMPLE.secure.grax.io/';
const APIToken = '<REDACTED>';
const GRAXListName = 'GRAX_DATA_LIST';
const GRAXDemoName = 'GRAX_DEMO';
const timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

function getHeader(method){
 var options = {
    'method' : method,
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + APIToken
    }
  };
  return options;
}

function getReturnData(data){
  if (data.length > 0) {
        const header = Object.keys(data[0])
        const values = data.map(Object.values);
        values.unshift(header);
        return values
      } else {
        throw new Error('No Data')
    }
}

function getAPIData(urlPath, method){
  var response = UrlFetchApp.fetch(APIName + urlPath, getHeader(method));
  var obj = JSON.parse(response.getContentText());
  return obj;
}

function getGRAXSearches() {
  var obj = getAPIData('api/v2/searches', 'get');
  return getReturnData(obj.searches);
}

function executeSearch(objectName,dateField,timeFrom,timeTo,status){
  var options = getHeader('post');
  var payload = {
    'object': objectName,
    'status': status,
    'timeField': dateField,
    'timeFieldMax': timeTo,
    'timeFieldMin': timeFrom
  }
  options.payload = JSON.stringify(payload);
  var response = UrlFetchApp.fetch(APIName + 'api/v2/searches',options);
  var data = JSON.parse(response.getContentText());
  Logger.log("Search Id: " + data.id);
  return data.id; 
}

function getSearchCsv(searchId,fields){
  var sUrl = APIName + 'api/v2/searches/' + searchId + '/download';
  if (fields!=null && fields!=''){
    sUrl+='?latest=true&fields=' + fields;
  }
  var response = UrlFetchApp.fetch(sUrl, getHeader('get'));
  var unZippedfile = Utilities.unzip(response);
  var blob = unZippedfile[1];
  var csvData = Utilities.parseCsv(blob.getDataAsString());
  return csvData;
}

function getSearchData(searchId){
  var response = UrlFetchApp.fetch(APIName + 'api/v2/searches/' + searchId + '/download?latest=true&fields=Id', getHeader('get'));
  var unZippedfile = Utilities.unzip(response);
  var blob = unZippedfile[1];
  var csvData = Utilities.parseCsv(blob.getDataAsString());
  const header = Object.keys(csvData[0])
  const values = csvData.map(Object.values);
  values.unshift(header);
  return values;
}

//@OnlyCurrentDoc
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("GRAX Data 👉️")
    .addItem("Refresh All GRAX Data", "refreshAllSearchData")
    .addToUi();
}

//Displays an alert as a Toast message
function displayToastAlert(message) {
  SpreadsheetApp.getActive().toast(message + "\n\nContact sales@grax.com with questions.", "GRAX Alert"); 
}

function getDate(){
  return Utilities.formatDate(new Date(), timezone, 'M/d/yyyy HH:mm');
}

function makeListActive(){
  var ss = SpreadsheetApp.getActive();
  var listSheet = ss.getSheetByName(GRAXListName);
  listSheet.activate();
  return listSheet;
}

function refreshAllSearchData() {
  var ss = SpreadsheetApp.getActive();
  var listSheet = makeListActive();
  if (listSheet.getName()==GRAXListName){
    var data = ss.getDataRange().getValues();
    for(var i = 0; i < data.length; i++){
      listSheet = makeListActive();
      if (i!=0){
        var objectName = data[i][0];
        var fields = data[i][1];
        var dateMin = data[i][2];
        var dateMax = data[i][3];
        var dateField = data[i][4];
        var status = data[i][5];
        var latestSearchId = data[i][6].toString().trim();
        if (latestSearchId=="" || latestSearchId==null){
          displayToastAlert('Executing GRAX search for ' + objectName);
          latestSearchId = executeSearch(objectName,dateField,dateMin,dateMax,status);
          ss.getRange('G' + (i+1)).setValue(latestSearchId);
          ss.getRange('H' + (i+1)).setValue(getDate());
        }else{
          Logger.log('Skipping Search Execution (ID Exists) ' + objectName + ' SearchID="' + latestSearchId + '"')
        }
        displayToastAlert('Refreshing ' + objectName + ' From Existing Search "' + latestSearchId + '"');
        var sheetName = 'GRAX_DATA_'+ objectName;
        var dataSheet = ss.getSheetByName(sheetName);
        if (dataSheet==null){
          dataSheet = ss.insertSheet();
          dataSheet.setName('GRAX_DATA_'+ objectName);
          SpreadsheetApp.getActive().moveActiveSheet(SpreadsheetApp.getActive().getNumSheets());
        }else{
          dataSheet.activate();
        }
        dataSheet.clear();
        var dataSheet = ss.getSheetByName(sheetName);
        updateSheet(dataSheet,getSearchCsv(latestSearchId,fields));

        listSheet = makeListActive();
        ss.getRange('I' + (i+1)).setValue(getDate());
      }
    }
    var dataSheet = ss.getSheetByName(GRAXDemoName);
    dataSheet.activate();
    displayToastAlert('Successfully Refreshed All GRAX Data.'); 
  }else{
      displayToastAlert('Please run from ' + GRAXListName + ' Tab! ' + ss.getName()); 
  }
}

function updateSheet(sheet,data){
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

function writeDataToSheet(data) {
  var ss = SpreadsheetApp.getActive();
  sheet = ss.insertSheet();
  updateSheet(sheet,data);
  return sheet.getName();
}

