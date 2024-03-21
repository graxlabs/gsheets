const APIName = 'https://EXAMPLE.secure.grax.io/';
const APIToken = '<REDACTED>';

const SnapshotTabName = 'GRAX_Snapshots';
const SearchTabName = 'GRAX_Searches';
const GRAXSegmentKeyName = 'Snapshot Date';

var Object_Counter = 0;

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

// getSearchStatus('MeAyKMuMmM9PY8zffFEC6C');
function getSearchStatus(searchId){
  var obj = getAPIData('api/v2/searches/' + searchId, 'get');
  return obj['status'];
}

// executeSearch("Case","createdAt","2024-03-01T04:00:00Z","2024-03-14T05:00:00Z","live");
function executeSearch(objectName,dateField,timeFrom,timeTo,status,filter){
  var options = getHeader('post');
  var payload = {
    'object': objectName,
    'status': status,
    'timeField': dateField,
    'timeFieldMax': timeTo,
    'timeFieldMin': timeFrom
  };

  if (filter != null){
    payload.filters={
      "mode":"and",
      "fields" : [
        filter
      ]
    }
  }

  options.payload = JSON.stringify(payload);
  Logger.log(options.payload);
  displayAlert('Retrieving GRAX History for ' + objectName + ' ' + timeFrom + ' to ' + timeTo);
  var response = UrlFetchApp.fetch(APIName + 'api/v2/searches',options);
  var data = JSON.parse(response.getContentText());
  return data.id; 
}

function executeSearchAndWait(objectName,dateField,timeFrom,timeTo,status,filter){
  var searchId = executeSearch(objectName,dateField,timeFrom,timeTo,status,filter);
  var searchStatus = getSearchStatus(searchId);
  Utilities.sleep(1000);
  while(searchStatus!='success'){
    Utilities.sleep(1000);
    searchStatus = getSearchStatus(searchId);
  }
  return searchId;
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

function downloadSearchData(searchId){
  var response = null;
  try{
    response = UrlFetchApp.fetch(APIName + 'api/v2/searches/' + searchId + '/download?latest=true&fields=Id', getHeader('get'));
  }catch(exception){
    displayAlert('Exception: ' + exception.toString());
  }
  return response;
}

// ----------------------------------------------------------------
function getFilter(filterField,filterType,filterValue){
  var filter = 
    {
      "field" : filterField.toString(),
      "filterType" : filterType.toString(),
      "not":false,
      "value" : filterValue.toString()
    };

  if (filterField!=null && filterField!='')  
   return filter;
  else
    return null;
}

// refreshAllSnapshots();
function refreshAllSnapshots() {
  var listSheet = makeSheetActive(SnapshotTabName);
  if (listSheet.getName() == SnapshotTabName){
    var data = listSheet.getDataRange().getValues();
    for(var i = 0; i < data.length; i++){
      if (i!=0){
        var objectName = data[i][0];
        var fields = data[i][1];
        var dateField = data[i][2];
        var segmentStart = data[i][3];
        var numberofSegments = data[i][4];
        var sheetName = data[i][5];
        var filter = getFilter(data[i][6],data[i][7],data[i][8]);
        var cumulative = data[i][9];
        var searchStartDate = data[i][10];
        if (searchStartDate==null || searchStartDate==''){
          searchStartDate=segmentStart;
        }
        displayAlert('Starting Snapshot:' + objectName + ' segmentStart: ' + segmentStart + ' Sheet Name: ' + sheetName + ' searchStartDate: ' + searchStartDate);
        runSnapShot(objectName,fields,dateField,segmentStart,numberofSegments,sheetName,filter,cumulative,searchStartDate);
      }
    }
    displayAlert('Successfully Refreshed Snapshot Data.'); 
  }
}

function runSnapShot(objectName,fields,dateField,segmentStart,numberofSegments,sheetName,filter,isCumulative,searchStartDate){
  var startDate = new Date(segmentStart);
  var endDate = null;
  var firstDayOfSearch = null;
  var headers = null;
  var aggregatedData=null;
  for(var i=0; i < numberofSegments; i++){
    if (isCumulative.toString().toLowerCase()=="true"){
      firstDayOfSearch=new Date(searchStartDate);
    }else{
      firstDayOfSearch = new Date(segmentStart.getFullYear(), startDate.getMonth() + i);
    }
    endDate = getLastDayOfMonth(new Date(segmentStart.getFullYear(), startDate.getMonth() + i));
    // Skip Searching History > Last Day of Current Month (FUTURE)
    if (getLastDayOfMonth(Date.now()) >= getLastDayOfMonth(endDate)){
      var segmentKey = endDate.toLocaleString('default',{ month: 'long' }) + ' ' + endDate.getFullYear().toString();
      latestSearchId = executeSearchAndWait(objectName,dateField,firstDayOfSearch.toISOString(),endDate.toISOString(),"live",filter);
      csvSearchhData=getSearchCsv(latestSearchId,fields);
      const columnsToRemove = [0, 1, 2, 3, 4, 5, 6, 7, 8];
      csvSearchhData = removeColumns(csvSearchhData, columnsToRemove);
      headers = csvSearchhData[0];
      csvSearchhData.splice(0,1);
      addColumn(csvSearchhData,headers.length+1,segmentKey);
      if (aggregatedData==null){
        headers[headers.length]=GRAXSegmentKeyName;
        aggregatedData = new Array(headers);
      }
      aggregatedData = [].concat(aggregatedData,csvSearchhData);
    }
  }
  var dataSheet = getNewSheet(sheetName);
  updateSheet(dataSheet,aggregatedData);
}

// ----------------------------------------------------------------
// GRAX Search Processing
// refreshAllSearchData();
function refreshAllSearchData() {
  var listSheet = makeSheetActive(SearchTabName);
  if (listSheet.getName()==SearchTabName){
    var data = listSheet.getDataRange().getValues();
    for(var i = 0; i < data.length; i++){
      listSheet = makeSheetActive(SearchTabName);
      if (i!=0){
        var objectName = data[i][0];
        var fields = data[i][1];
        var dateMin = data[i][2];
        var dateMax = data[i][3];
        var dateField = data[i][4];
        var status = data[i][5];
        var latestSearchId = data[i][6].toString().trim();
        latestSearchId = executeSearchAndWait(objectName,dateField,dateMin,dateMax,status);
        listSheet.getRange('G' + (i+1)).setValue(latestSearchId);
        listSheet.getRange('H' + (i+1)).setValue(getDate());
        var sheetName = 'GRAX_DATA_'+ objectName + '_' + Object_Counter;
        var dataSheet = fillSheet(sheetName,latestSearchId,fields);
        listSheet = makeSheetActive(SearchTabName);
        listSheet.getRange('I' + (i+1)).setValue(getDate());
        Object_Counter++;
      }
    }
    makeSheetActive(SearchTabName);
    displayAlert('Successfully Refreshed Seach Data.'); 
  } else {
      displayAlert('Please run from ' + SearchTabName + ' Tab! ' + ss.getName()); 
  }
}

// refreshAllGRAXData();
function refreshAllGRAXData(){
  refreshAllSnapshots();
  refreshAllSearchData();
}

// ------------------------------------------------
// Google Sheets Helper Functions

// Add GRAX Menu Options
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("GRAX Data 🚀")
    .addItem("Refresh All GRAX Data", "refreshAllGRAXData")
    .addItem("Refresh GRAX Snapshots", "refreshAllSnapshots")
    .addItem("Refresh GRAX Object Searches", "refreshAllSearchData")
    .addItem("GRAX Documentation", "documentationPopUp")
    .addItem("GRAX.com", "websitePopUp")
    .addToUi();
}

function displayAlert(message) {
  Logger.log(message);
  SpreadsheetApp.getActive().toast(message, "GRAX Alert"); 
}

function makeSheetActive(sheetName){
  var ss = SpreadsheetApp.getActive();
  var listSheet = ss.getSheetByName(sheetName);
  listSheet.activate();
  return listSheet;
}

function getNewSheet(sheetName){
  var ss = SpreadsheetApp.getActive();
  var dataSheet = ss.getSheetByName(sheetName);
  if (dataSheet==null){
    dataSheet = ss.insertSheet();
    dataSheet.setName(sheetName);
    SpreadsheetApp.getActive().moveActiveSheet(SpreadsheetApp.getActive().getNumSheets());
  }else{
    dataSheet.activate();
    dataSheet.clear();
  }
  return dataSheet;
}

function fillSheet(sheetName,latestSearchId,fields){
  var dataSheet = getNewSheet(sheetName);
  updateSheet(dataSheet,getSearchCsv(latestSearchId,fields));
  return dataSheet;
}

function updateSheet(sheet,data){
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}
// ------------------------------------------------
// Utility Functions
function getDate(){
  return Utilities.formatDate(new Date(), timezone, 'M/d/yyyy HH:mm');
}

function getLastDayOfMonth(date){
  var currentDate = new Date(date);
  var lastDayOfMonth =new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0, 23,59);
  return lastDayOfMonth;
}

 // Function to remove specified columns
function removeColumns(array, columnsToRemove) {
  return array.map(row =>
    row.filter((_, index) => !columnsToRemove.includes(index))
  );
}

function addColumn(array, columnIndex, value) {
  array.forEach(row => {
    row.splice(columnIndex, 0, value);
  });
}

function openURL(url){
  // var url = 'https://documentation.grax.com/docs/introduction';
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function() {google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410) </script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening GRAX. Make sure the pop up is not blocked." );

}

function documentationPopUp(){
  openURL('https://documentation.grax.com/docs/introduction');
}

function websitePopUp(){
  openURL('https://www.grax.com');
}


