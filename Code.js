/* eslint no-var: 0 */

/**
 * Generate unique identifiers (UID) in the sheet.
 *
 */
function generateUID() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('Data');
    var columnHeaderValues = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues();
    var UID_COLUMN = columnHeaderValues[0].indexOf('UID') + 1;
    var lastRow = sheet.getLastRow();
    var uidRange = sheet.getRange(2, UID_COLUMN, lastRow);
    var uidValues = uidRange.getValues();
    var uidCounterRange = spreadsheet.getRangeByName('UID_Counter');
    var nextCount = uidCounterRange.getValue();
    var rangeLength = uidRange.getHeight();

    for (var i = 0; i < rangeLength-1; i++) {
      if (uidValues[i][0] == 0) {
        uidValues[i][0] = nextCount++;
      }
    }
    uidCounterRange.setValue(nextCount);
    uidRange.setValues(uidValues).setHorizontalAlignment('center');
  } catch (error) {
    Logger.log('%s : %s', error.name, error.message);
  }
}

/* exported updateFields */
/**
 *
 *
 * @param {*} e
 */
function updateFields(e) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var columnHeaderValues = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues();
  var TIMESTAMP_COLUMN = columnHeaderValues[0].indexOf('Modifiée') + 1;
  var STATUS_COLUMN = columnHeaderValues[0].indexOf('Statut') + 1;
  var STARTED_COLUMN = columnHeaderValues[0].indexOf('Débutée') + 1;
  var ENDED_COLUMN = columnHeaderValues[0].indexOf('Terminée') + 1;
  var ACHIEVEMENT_COLUMN = columnHeaderValues[0].indexOf('Achèvement') + 1;
  var ACCOUNTABLE_COLUMN = columnHeaderValues[0].indexOf('Responsable') + 1;
  var activeSheet = e.source.getActiveSheet();
  var activeSheetName = activeSheet.getSheetName();
  var activeRange = e.range;
  var currentCell = sheet.getCurrentCell();
  var firstRow = activeRange.getRow();
  var firstColumn = activeRange.getColumn();
  //  var lastRow = activeRange.getLastRow();
  var lastColumn = activeRange.getLastColumn();
  var numRows = activeRange.getNumRows();
  var timestampArray = [];
  var timestampFormatArray = [];
  var currentDate = new Date();

  if ((firstRow == 1) || (activeSheetName != 'Data') ||
    ((firstColumn == TIMESTAMP_COLUMN) &
      (lastColumn == TIMESTAMP_COLUMN))) {
    return;
  }

  if (((firstColumn == STATUS_COLUMN) &
      (lastColumn == STATUS_COLUMN)) & (numRows == 1)) {
    if (currentCell.getValue() == 'En cours') {
      if (activeSheet.getRange(firstRow, STARTED_COLUMN).getValue() == '') {
        activeSheet.getRange(firstRow, STARTED_COLUMN).setValue(currentDate);
      }
      activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).setValue(0);
    }
    if (currentCell.getValue() == 'Assignée') {
      if (activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).getValue() == '') {
        activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).setValue(0);
      }
    }
    if ((currentCell.getValue() == 'En cours') ||
      (currentCell.getValue() == 'Assignée') ||
      (currentCell.getValue() == 'Suspendue')) {
      activeSheet.getRange(firstRow, ENDED_COLUMN).setValue('');
    }
    if ((currentCell.getValue() == 'Terminée') ||
      (currentCell.getValue() == 'Annulée')) {
      activeSheet.getRange(firstRow, ENDED_COLUMN).setValue(currentDate);
      activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).setValue(1);
    }
  }

  if (((firstColumn == ACCOUNTABLE_COLUMN) &
      (lastColumn == ACCOUNTABLE_COLUMN)) & (numRows == 1)) {
    if ((currentCell.getValue() != '') &
      (activeSheet.getRange(firstRow, STATUS_COLUMN).getValue() == '')) {
      activeSheet.getRange(firstRow, STATUS_COLUMN).setValue('Assignée');
      if (activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).getValue() == '') {
        activeSheet.getRange(firstRow, ACHIEVEMENT_COLUMN).setValue(0);
      }
    }
  }


  for (var i = 0; i < numRows; i++) {
    timestampArray.push([]);
    timestampFormatArray.push([]);
    timestampArray[i][0] = currentDate;
    timestampFormatArray[i][0] = 'yyyy-mm-dd  \\[hh:mm\\]';
  }

  sheet.getRange(firstRow, TIMESTAMP_COLUMN, numRows)
    .setNumberFormats(timestampFormatArray).setValues(timestampArray);
}

/* exported addStrimanMenu */
/**
 * Create a custom menu for this S:Triman spreadsheet.
 *
 */
function addStrimanMenu() {
  SpreadsheetApp.getUi()
    .createMenu('S:Triman')
    .addItem('Afficher Catégories', 'showDialog')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Options avancées')
      .addItem('Générer UID', 'generateUID')
      .addItem('Insérer NewLine dans Catégorie', 'insertNewLineInCategory')
      .addItem('Supprimer la cache \'Catégories\'', 'removeCachedCategories'))
    .addToUi();

  generateUID();
}

/* exported showDialog */
/**
 * Show a sidebar with a list of check box items.
 * This list comes from a data validation range.
 * This function is called from 'Page.html'.
 *
 */
function showDialog() {
  var html = HtmlService.createTemplateFromFile('Page').evaluate()
    .setTitle('Liste des catégories');
  SpreadsheetApp.getUi().showSidebar(html);
}

/* exported valid */
/**
 * Returns an array of categories, as defined by a data validation range.
 *
 * @return {array}
 */
var valid = function() {
  try {
    return getCategories();
  } catch (e) {
    return null;
  }
};

/* exported currentSelection */
/**
 * Returns the strings contained in the current cell, in an array and
 * separated by 'new line'.
 *
 * @return {array}
 */
var currentSelection = function() {
  try {
    var arrayOfValues = SpreadsheetApp.getActiveRange().getValue().split('\n');
    return arrayOfValues;
  } catch (e) {
    return null;
  }
};

/* exported fillCell */
/**
 * Sets the value of the current cell with user's sidebar selection.
 *
 * @param {*} e
 */
function fillCell(e) {
  // First, check that the selected cell's column is valid
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnHeaderValues = activeSheet.getRange(1, 1, 1, activeSheet
    .getLastColumn()).getValues();
  var CATEGORY_COLUMN = columnHeaderValues[0].indexOf('Catégorie') + 1;

  // If not, refresh sidebar. Otherwise fill in selected cell with user choice.
  if (activeSheet.getCurrentCell().getColumn() != CATEGORY_COLUMN) {
    showDialog();
  } else {
  var s = [];
  for (var i in e) {
    if (i.substr(0, 2) == 'ch') s.push(e[i]);
  }
  if (s.length) SpreadsheetApp.getActiveRange().setValue(s.join('\n'));
  }
}

/**
 * Returns an array of categories from cached data if available.
 * If not, returns an array of categories from 'Catégorie' named range and
 * put the data in cache for faster future calls.
 *
 * @return {array}
 */
function getCategories() {
  // if (SpreadsheetApp.getActiveRange().getDataValidation() == null)
  //  return null;
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var columnHeaderValues = activeSheet.getRange(1, 1, 1, activeSheet
    .getLastColumn()).getValues();
  var CATEGORY_COLUMN = columnHeaderValues[0].indexOf('Catégorie') + 1;

  if (SpreadsheetApp.getCurrentCell().getColumn() != CATEGORY_COLUMN) {
    return null;
  }

  var cache = CacheService.getScriptCache();
  var cached = cache.get('categories');
  if (cached != null) {
    var newArray1D = cached.split(',');
    var newArray2D = [];
    while (newArray1D.length) newArray2D.push(newArray1D.splice(0, 1));
    Logger.log('newArray2D = %s', newArray2D);
    return newArray2D;
  }
  // var categoryArray = SpreadsheetApp.getActiveRange().getDataValidation()
  // .getCriteriaValues()[0].getValues();
  var categoryArray = SpreadsheetApp.getActiveSpreadsheet()
    .getRangeByName('Catégories').getValues();
  cache.put('categories', categoryArray, 1500);
  Logger.log('categoryArray = %s', categoryArray);
  return categoryArray;
}

/* exported removeCachedCategories */
/**
 * Remove categories from cache.
 *
 */
function removeCachedCategories() {
  var cache = CacheService.getScriptCache();
  cache.remove('categories');
}

/* exported showCachedCategories */
/**
 *
 *
 */
function showCachedCategories() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('categories');
  if (cached != null) {
    Logger.log('Catégories = %s', cached);
  } else {
    Logger.log('No categories in cache');
  }
}

/* exported insertNewLineInCategory */
/**
 * Insert 'new line' after each string of the cells of a column 'Catégorie'..
 *
 */
function insertNewLineInCategory() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var columnHeaderValues = activeSheet.getRange(1, 1, 1, activeSheet
    .getLastColumn()).getValues();
  var CATEGORY_COLUMN = columnHeaderValues[0].indexOf('Catégorie') + 1;
  var lastRow = activeSheet.getLastRow();
  var categoryRange = activeSheet.getRange(2, CATEGORY_COLUMN, lastRow - 1);
  var categoryValues = categoryRange.getValues();

  for (var i = 0; i < categoryValues.length; i++) {
    categoryValues[i][0] = categoryValues[i][0].toString()
      .replace(/, /g, '\n');
  }
  categoryRange.setValues(categoryValues);
}
