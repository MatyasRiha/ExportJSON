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
  Logger.log(finalTextA); //372
  var lastItem = finalTextA.length; //372, because i need to write at index 372, because N = 1, so 0 is clear for {
  var rocketRoute = lastItem - 1; //371, because last index of returning arrays is 372, we are not chcecking for row number 1 in foor loop
  var convert = new Array();
  if (finalTextA.length === finalTextB.length)
  {
    const loop = lastItem - 1; //371 How many times it should loop, it shouldnt loop last item, because of its difference
    var b = 0;
    var n = 1;
    for (var i = loop; i != 0; i--)
    {
      convert[n] = ('"' + finalTextA[b] + '"' + ':' + ' ' + finalTextB[b] + ',' + '    '); //After loop N = 372
      b++;
      n++;
    }
    convert[lastItem] = ('"' + finalTextA[rocketRoute] + '"' + ':' + ' ' + finalTextB[rocketRoute]);
  }
  else
  {
    const error = 'Error-tailRegistartion and Rocket Route ID length is not matching'
    displayText_(error);
  }
  var lastBracket = lastItem + 1; //373 the last bracket have index of 1 + the last text
  convert[0] = '{    ' //Index 0 is clear for bracket
  convert[lastBracket] = ' }'
  //Logger.log(convert);
  var json = makeJSON(convert); //Calling make json function with argument convert-array
  displayText_(json);
}

//Function making from a array (convert) to a string 
function makeJSON(convertText)
{
  var finalString = convertText.join(''); //Making array
  return finalString; //Returning final string 
}

//Function dipslying text
function displayText_(text) {
  var output = HtmlService.createHtmlOutput("<textarea style='width:100%;' rows='20'>" + text + "</textarea>");
  output.setWidth(400)
  output.setHeight(300);
  SpreadsheetApp.getUi()
      .showModalDialog(output, 'Exported JSON');
}
