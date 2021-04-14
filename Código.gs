var userProperties = PropertiesService.getUserProperties();

var BASEAPIURL = 'https://api.bamboohr.com/api/gateway.php/ubiwhere/v1/'
var EMPLOYEESAPIURL = BASEAPIURL + 'employees/directory'

var HolidaysSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('holidays');

function onOpen() {   
  ui = SpreadsheetApp.getUi()

  var message = 'Dont forget to setup your API Key, start and end dates';
  var title = 'Welcome to BambooHR holidays!';
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);

  ui
  .createMenu('BambooHR')
  .addItem('Set API key', 'setAPIKey')
  .addItem('Delete API key', 'resetKey')
  .addItem('Set Start Date', 'setStartDate')
  .addItem('Set End Date', 'setEndDate')
  .addItem('Get Holidays from BambooHR', 'writeHolidaysOnSheet')
  .addSeparator()
  .addToUi();
}

function checkSetup(){
  try{
    var startDate = userProperties.getProperty('STARTDATE')
    var endDate = userProperties.getProperty('ENDDATE')
    var apiKey = userProperties.getProperty('APIKEY')
  }
  catch(e){
      var message = 'Dont forget to setup ' + e.message.split(' ')[0];
      var title = 'Setup ' + e.message.split(' ')[0];
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      return false
  }

  if (startDate==null || startDate==""){
      var message = 'Dont forget to setup' + e.message.split(' ')[0];
      var title = 'Setup' + + e.message.split(' ')[0];
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      return false
    }
    else if (endDate==null || endDate==""){
      var message = 'Dont forget to setup your end date';
      var title = 'Setup your end date!';
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      return false
    }
    else if (apiKey==null || apiKey==""){
      var message = 'Dont forget to setup your API Key';
      var title = 'Setup your API KEY!';
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      return false
    }
  return true
}

function setStartDate(){
  ui = SpreadsheetApp.getUi()
  var startDate = ui.prompt('Please the start date in the format YYYY-MM-DD.' , ui.ButtonSet.OK);
  userProperties.setProperty('STARTDATE', startDate.getResponseText());
}

function setEndDate(){
  ui = SpreadsheetApp.getUi()
  var endDate = ui.prompt('Please the end date in the format YYYY-MM-DD.' , ui.ButtonSet.OK);
  userProperties.setProperty('ENDDATE', endDate.getResponseText());
}

function setAPIKey(){
  ui = SpreadsheetApp.getUi()
  var scriptValue = ui.prompt('Please provide your API key.' , ui.ButtonSet.OK);
  userProperties.setProperty('APIKEY', scriptValue.getResponseText());
}

function resetKey(){
  userProperties.deleteProperty(APIKEY);
}

function writeHolidaysOnSheet(){
  var headers = {
    'Accept': 'application/json',
    'Authorization': 'Basic ' + Utilities.base64Encode(userProperties.getProperty('APIKEY') + ":" + '')
  }
  
  if (checkSetup()==true){
    HolidaysSheet.clear();
    HolidaysSheet.getRange("A1:H1").setValues([
      ['user','Name', 'From', 'To', 'Type', 'Amount', 'Status', 'Tag']
    ]);


    var holidays = getHolidays(headers)
    var employees_dict = getEmployees(headers)
    var row = 0
    for (i = 0; i < holidays.length; i++) { 
      var user = employees_dict[holidays[i].employeeId]
      if (user != "" && user != null){
        var name = holidays[i].name;
        var start = holidays[i].start;
        var end = holidays[i].end;
        var type = holidays[i].type.name;
        var amount = holidays[i].amount.amount + ' ' + holidays[i].amount.unit;
        var status = holidays[i].status.status;
        var tag = start + " - " + end

        HolidaysSheet.getRange(row+2, 1, 1, 8).setValues([[user, name, start, end, type, amount, status, tag]])
        row += 1
      }
    }
  }
}

function getEmployees(headers){
  var response = callAPIwithGet(EMPLOYEESAPIURL, false, headers);
  var users = JSON.parse(response);
  var employees = users.employees
  var employees_dict = {}
  for (var i = 0; i < employees.length; i++){
    try{
      var employee_id = employees[i].id
      var employee = employees[i].workEmail.split('@')[0]
      employees_dict[employee_id] = employee
    }
    catch(e){
      continue
    } 
  }
  return employees_dict
}

function getHolidays(headers){
  startDate = userProperties.getProperty('STARTDATE')
  endDate = userProperties.getProperty('ENDDATE')
  var holidaysAPIUrl = BASEAPIURL + 'time_off/requests/?start='+ startDate + '&end='+ endDate
  var response = callAPIwithGet(holidaysAPIUrl, false, headers);
  var holidays = JSON.parse(response.getContentText());
  return holidays
}

function callAPIwithGet(url, muteHttpExceptions, headers) {
  var options = {
            'method': 'get',
            'headers': headers ,
            'muteHttpExceptions': muteHttpExceptions
  };
  Logger.log(options)
  var response = UrlFetchApp.fetch(url, options); 
  if(!response)
  Logger.log("API request failed: " + url); 
  return response;
}