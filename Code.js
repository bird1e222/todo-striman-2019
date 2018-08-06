//-----------------------------------------------------------------------------------------------------------------
function generateUID() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = spreadsheet.getSheetByName("Data");   
  var columnHeaderValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var UID_COLUMN = columnHeaderValues[0].indexOf("UID") + 1;
  var lastRow     = sheet.getLastRow();
  var uidRange    = sheet.getRange(2, UID_COLUMN, lastRow-1);
  var uidValues   = uidRange.getValues();
  var uidCounterRange = spreadsheet.getRangeByName('UID_Counter');
  var nextCount = uidCounterRange.getValue();

  for (var row in uidValues) {
    if (uidValues[row][0] == 0) {
      uidValues[row][0] = nextCount++;
    }
  }
  
  uidCounterRange.setValue(nextCount);
  uidRange.setValues(uidValues).setHorizontalAlignment("center");
}


//-----------------------------------------------------------------------------------------------------------------
function updateFields(e){

  var sheet = SpreadsheetApp.getActiveSheet();
  var columnHeaderValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var TIMESTAMP_COLUMN = columnHeaderValues[0].indexOf("Modifiée") + 1;
  var STATUS_COLUMN = columnHeaderValues[0].indexOf("Statut") + 1;
  var STARTED_COLUMN = columnHeaderValues[0].indexOf("Débutée") + 1;
  var ENDED_COLUMN = columnHeaderValues[0].indexOf("Terminée") + 1;
  var ACHIEVEMENT_COLUMN = columnHeaderValues[0].indexOf("Achèvement") + 1;
  var ACCOUNTABLE_COLUMN = columnHeaderValues[0].indexOf("Responsable") + 1;
  var activeSheet = e.source.getActiveSheet();
  var activeSheetName = activeSheet.getSheetName();
  var activeRange = e.range;
  var currentCell = sheet.getCurrentCell();
  var firstRow = activeRange.getRow();
  var firstColumn = activeRange.getColumn();
  var lastRow = activeRange.getLastRow();
  var lastColumn = activeRange.getLastColumn();
  var numRows = activeRange.getNumRows();
  var timestampArray = [];
  var timestampFormatArray = [];
  var currentDate = new Date();

  if ((firstRow == 1) || (activeSheetName != "Data") || ((firstColumn == TIMESTAMP_COLUMN) & (lastColumn == TIMESTAMP_COLUMN))) {return};
  
  if (((firstColumn == STATUS_COLUMN) & (lastColumn == STATUS_COLUMN)) & (numRows == 1)) {
    if (currentCell.getValue() == "En cours") {
      if (activeSheet.getRange(firstRow, STARTED_COLUMN).getValue() == "") {
        activeSheet.getRange(firstRow, STARTED_COLUMN).setValue(currentDate);
      }
      activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).setValue(0);
    }
    if (currentCell.getValue() == "Assignée") {
      if (activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).getValue() == ""){
        activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).setValue(0);
      }
    }
    if ((currentCell.getValue() == "En cours") || (currentCell.getValue() == "Assignée")  || (currentCell.getValue() == "Suspendue")) {
      activeSheet.getRange(firstRow, ENDED_COLUMN).setValue("");
    }
    if ((currentCell.getValue() == "Terminée") || (currentCell.getValue() == "Annulée")){
        activeSheet.getRange(firstRow, ENDED_COLUMN).setValue(currentDate);
        activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).setValue(1);
    }
  }

  if (((firstColumn == ACCOUNTABLE_COLUMN) & (lastColumn == ACCOUNTABLE_COLUMN)) & (numRows == 1)) {
    if ((currentCell.getValue() != "") & (activeSheet.getRange(firstRow, STATUS_COLUMN).getValue() == "")) {
      activeSheet.getRange(firstRow, STATUS_COLUMN).setValue("Assignée");
      if (activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).getValue() == ""){
        activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).setValue(0);
      }
    }
  }

  
  for (var i = 0; i < numRows; i++) {
    timestampArray.push([]);
    timestampFormatArray.push([]);
    timestampArray[i][0] = currentDate;
    timestampFormatArray[i][0] = "yyyy-mm-dd  \\[hh:mm\\]";
  }

  sheet.getRange(firstRow, TIMESTAMP_COLUMN, numRows).setNumberFormats(timestampFormatArray).setValues(timestampArray);

}



//-----------------------------------------------------------------------------------------------------------------
function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu('S:Triman')
  .addItem('Afficher Catégories', 'showDialog')
  .addToUi();
}


//-----------------------------------------------------------------------------------------------------------------
function showDialog() {  
  var html = HtmlService.createTemplateFromFile('Page').evaluate();
  SpreadsheetApp.getUi()
  .showSidebar(html);
}
var valid = function(){
  try{
    return SpreadsheetApp.getActiveRange().getDataValidation().getCriteriaValues()[0].getValues();
  }catch(e){
    return null
  }
}

var currentSelection = function(){
  try{
    var arrayOfValues = [{}];
    var arrayOfValues = SpreadsheetApp.getActiveRange().getValue().split("\n");    
    return arrayOfValues;
  }catch(e){
    return null
  }
}



//-----------------------------------------------------------------------------------------------------------------
function fillCell(e){
  var s = [];
  for(var i in e){
    if(i.substr(0, 2) == 'ch') s.push(e[i]);
  }
  if(s.length) SpreadsheetApp.getActiveRange().setValue(s.join('\n'));
}
