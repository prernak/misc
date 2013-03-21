/**
* Script for Spreadsheet @
* https://docs.google.com/spreadsheet/pub?key=0AtNxSwigiymodHk3aEtfQUNQVjhJSG9STFNNNTlSMXc&output=html
*/


function getAddress(name,index,sheet) {
  var YOUR_KEY="";
  var url ="https://maps.googleapis.com/maps/api/place/"
  + "textsearch/json?query="
  + name
  + "&"
  + "sensor=true&key="+YOUR_KEY;
  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText();
  var data = JSON.parse(json);
  Logger.log("Results : "+json);
  if(data.status == "OVER_QUERY_LIMIT") {
    Browser.msgBox("OVER_QUERY_LIMIT");
    Logger.log("OVER_QUERY_LIMIT");
    return data.status;
  }    
  var arr = new Array();
  for(i =0 ; i < data["results"].length; i++ ) {    
    var str =     data["results"][i]["formatted_address"];
    Logger.log(name+" : ");
    Logger.log(str);
    if(str.indexOf("MA") != -1)
      arr.push(str);       
  }
  sheet.getRange("F"+(index+1)).setValue(arr.join("\n"));
  return "DONE";
};

/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  for (var i = 1; i <= numRows - 1; i++) {
    var name = values[i][1].toString();
    Logger.log(name+" : Company");
    if(name.length > 0) {
      var val = getAddress(name,(i+1),sheet);
      if(val== "OVER_QUERY_LIMIT")
        return;
    }
    SpreadsheetApp.flush();
    Utilities.sleep(1000);
  }
};


/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Read Data",
    functionName : "readRows"
  }];
  sheet.addMenu("Get Company Address", entries);
};
