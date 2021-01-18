//PROGRAM FOR EXPORTING JSON FL3XX FOR ROCKET-ROUTE
//MATYAS RIHA 2021


//Creating GUI menu and options
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('FL3XX')
      .addItem('Export JSON', 'Export')
      .addSeparator()
      .addItem('Import JSON', 'Import')
      .addToUi();

}

//Function that stores main loop and makeJSON
function getValues()
{
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  //Declaring variables rows, columns and text for creating JSON 
  var row = sheet.getLastRow();
  const column = 1;
  //Main loop for reading each A column and creating JSON output
  var tailRegistration = [] //Array storing different tail registrations
  for (var i = row; i > 1; i--){ 
    tailRegistration.unshift(sheet.getRange(i, column).getValue());
}
  return tailRegistration; //Returning tailRegistration array to function caller, caller is in the function Export
}

function getValuesB()
{
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  //Declaring variables rows, columns and text for creating JSON 
  var row = sheet.getLastRow();
  const column = 2;
  //Main loop for reading each A column and creating JSON output
  var rocketRoute = [] //Arrray storing different Rocket Route ID's
  for (var i = row; i > 1; i--){
    rocketRoute.unshift(sheet.getRange(i, column).getValue());
}
  return rocketRoute; //Returning rocketRoute array to function caller, caller is in the function Export
}


//Function for JSON exporting and GUI 
function Export()
{
  //Acessing active sheet
  var ss = SpreadsheetApp.getActive();
  var z = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  //Declaring row variable and variables that stores data from getValues functions
  var row = sheet.getLastRow();
  var finalTextA = getValues(); //Variable calling function getValues and returning values from getValues function
  var finalTextB = getValuesB(); //Variable calling function getValuesB and returning values from getValuesB function
  var convert = new Array();
  if (finalTextA.length === finalTextB.length)
  {
    var b = 0;
    var n = 1;
    for (var i = row; i > 1; i--)
    {
      convert[n] = ('"' + finalTextA[b] + '"' + ':' + ' ' + finalTextB[b] + ',' + '    ');
      b++;
      n++;
    }
  }
  else
  {
    const error = 'Error-tailRegistartion and Rocket Route ID lenght is not matching'
    displayText_(error);
  }
  var lastBracket = row + 1;
  convert[0] = '{    '
  convert[lastBracket] = '}'
  var json = makeJSON(convert);
  Logger.log(json);
  displayText_(json);
  //Logger.log(finalTextA);
}

function makeJSON(convertText)
{
  var finalString = convertText.join('');
  return finalString;
}

function displayText_(text) {
  var output = HtmlService.createHtmlOutput("<textarea style='width:100%;' rows='20'>" + text + "</textarea>");
  output.setWidth(400)
  output.setHeight(300);
  SpreadsheetApp.getUi()
      .showModalDialog(output, 'Exported JSON');
}

