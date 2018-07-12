// Collection of useful sheet tools (Shools)
// See Menu.js for menu
// See Messages.js for Alert Boxes
  
var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
function cellCount() {
  
//  Set variables
  var countTab = "Cell Count";
  var names = [["Sheet Name"], ["Tab Name"], ["Cells Used"], ["Cells Max"]];
  var totalUsed = 0;
  var totalMax = 0;
  
//  Get sheet and load tabs into array
// var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = sheet.getSheets();
  
  if (sheet.getSheetByName(countTab) == null) {
      sheet.insertSheet(countTab); }

  var cTab = sheet.getSheetByName(countTab);
  cTab.clear();
  cTab.getRange(2, 1, names.length)
      .setValues(names)
      .setBackground("#cfe2f3");

  cTab.getRange(2, 2, 1, tabs.length + 1)
      .merge()
      .setValue(sheet.getName() + " - " + sheet.getId())
      .setBackground("#cfe2f3");
  
  for(var i = 0; i < tabs.length; i++) {
                    
    cellsUsed = (tabs[i].getLastColumn() * tabs[i].getLastRow());
    cellsMax = (tabs[i].getMaxColumns() * tabs[i].getMaxRows());
    
    totalUsed = cellsUsed + totalUsed;
    totalMax = cellsMax + totalMax;
    
     cTab.getRange(3, 2 + i, 3).setValues([[tabs[i].getName()], [cellsUsed], [cellsMax]]);  
  }
  
  if (totalUsed > 1500000) { var usedColor = "red"; }
  else { var usedColor = "green"; }
  
  if (totalMax > 1500000) { var maxColor = "red"; }
  else { var maxColor = "green"; }
  
  cTab.getRange(3, tabs.length + 2, 3).setValues([["Totals"], [totalUsed], [totalMax]])
      .setFontWeight("bold")
      .setHorizontalAlignment("right")
      .setBackgrounds([["#cfe2f3"], ["white"], ["white"]])
      .setFontColors([["black"], [usedColor], [maxColor]]);
  
  if (cTab.getMaxRows() > (cTab.getLastRow() + 1)) { 
    cTab.deleteRows((cTab.getLastRow() + 1), (cTab.getMaxRows() - (cTab.getLastRow() + 1))); }
  if (cTab.getMaxColumns() > (tabs.length + 3)) {
    cTab.deleteColumns((tabs.length + 3), (cTab.getMaxColumns() - (tabs.length + 3))); }
  }

  function trimSheets() {

// Connect to spreadsheet and directory 
  
// var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetID = sheet.getId();
  var tabs = sheet.getSheets();
  var driveFile = DriveApp.getFileById(sheetID);
    var timeZone = sheet.getSpreadsheetTimeZone();
    var dateTime = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd'_'HH:mm");

// Make backup copy of spreadsheet
  driveFile.makeCopy("Backup_" + sheet.getName() + "_" + dateTime);
    
// Remove unused rows and columns
    for(var i = 0; i < tabs.length; i++) {
      var usedCols = tabs[i].getLastColumn() + 1;
      var maxCols = tabs[i].getMaxColumns() - usedCols;
      var usedRows = tabs[i].getLastRow() + 1;
      var maxRows = tabs[i].getMaxRows() - usedRows;
      
      if(maxCols > usedCols){
       tabs[i].deleteColumns(tabs[i].getLastColumn() + 1, tabs[i].getMaxColumns() - (tabs[i].getLastColumn() + 1));
        }
      if(maxRows > usedRows){
       tabs[i].deleteRows(tabs[i].getLastRow() + 1, tabs[i].getMaxRows() - (tabs[i].getLastRow() + 1));
        }      
    }
 }

//  NEED TO WORK ON THIS FUNCTION
 function trimActive() {
  
  // var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var tab = sheet.getActiveSheet();
       
  // Remove unused rows and columns
        var usedCols = tab.getLastColumn() + 1;
        var maxCols = tab.getMaxColumns() - usedCols;
        var usedRows = tab.getLastRow() + 1;
        var maxRows = tab.getMaxRows() - usedRows;
   
   Logger.log(usedCols);
   Logger.log(maxCols);
        
        if(maxCols > 0){
         tab.deleteColumns(usedCols, maxCols);
          }
        if(maxRows > usedRows){
         tab.deleteRows(tab.getLastRow() + 1, tab.getMaxRows() - (tab.getLastRow() + 1));
          }      
   }   

function hideAll() {
  
  // Get sheet and tabs
// var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = sheet.getSheets();
  // Get active sheet name
  var activeTab = sheet.getActiveSheet().getName();
  
  // Check tab name against active tab and hide if false 
  for(var i=0; i < tabs.length; i++) {
    if(tabs[i].getName() != activeTab) {
      tabs[i].hideSheet();
    }
  }
}

function unHide() {
  
  // Get sheet and tabs
// var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = sheet.getSheets();

  // Unhide all sheets
  for(var i=0; i < tabs.length; i++) {
    tabs[i].showSheet();
  }
}

function tabsAndRange() {

// Set variables
  var dataTab = "Tab Names and Ranges";
  var names = [["Sheet Name"], ["Tab Name"], ["Active Range"]];
  var tabs = sheet.getSheets();

    if (sheet.getSheetByName(dataTab) == null) {
      sheet.insertSheet(dataTab);  }

  var dTab = sheet.getSheetByName(dataTab);
  dTab.clear();
  
  dTab.getRange(2, 1, names.length)
      .setValues(names)
      .setBackground("#cfe2f3");

  dTab.getRange(2, 2, 1, tabs.length + 1)
      .merge()
      .setValue(sheet.getName() + " - " + sheet.getId())
      .setBackground("#cfe2f3");
  
  for(var i = 0; i < tabs.length; i++) {
    
    dTab.getRange(3, 2 + i, 2).setValues([[tabs[i].getName()], [tabs[i].getDataRange().getA1Notation()]]);
  }
 
  if (dTab.getMaxRows() > (dTab.getLastRow() + 1)) { 
    dTab.deleteRows((dTab.getLastRow() + 1), (dTab.getMaxRows() - (dTab.getLastRow() + 1))); }
  if (dTab.getMaxColumns() > (tabs.length + 3)) {
    dTab.deleteColumns((tabs.length + 3), (dTab.getMaxColumns() - (tabs.length + 3))); }
  
  
}

function listToTabs() {

  var sheet = SpreadsheetApp.getActive();
  var list = sheet.getActiveRange().getValues();
  for(var i = 0; i < list.length; i++) {
    sheet.insertSheet(list[i][0]);
   }
}