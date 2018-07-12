function countMsg() {
  // Display count cells alert box
    var ui = SpreadsheetApp.getUi();
    var alertBox = ui.alert("Please confirm", "This function will create a new tab called 'Cell Count'. \n If the tab exist it will be cleared and updated.", ui.ButtonSet.OK_CANCEL);
    
    if (alertBox == ui.Button.OK) {
      cellCount(); }
    else {
      return; }
  }
  
  function trimMsg() {
  // Display trim sheets alert box
    var ui = SpreadsheetApp.getUi();
    var alertBox = ui.alert("Please confirm", "This function will remove unused rows and columns from ALL sheets in the spreadsheet. \n This is potentially destructive, a copy of the spreadsheet will be created as a backup before the sheets are trimmed \n The backup file can be delete after you are sure no data was lost.", ui.ButtonSet.OK_CANCEL);
  
    if (alertBox == ui.Button.OK) {
      trimSheets(); }
      else {
        return; }
  }

  function tabsMsg() {
  // Displays tab names and range alert box
    var ui = SpreadsheetApp.getUi();
    var alertBox = ui.alert("Please confirm", "This function will create a new tab called 'Tab Names and Ranges'. \n If the tab exist it will be cleared and updated.", ui.ButtonSet.OK_CANCEL);

    if (alertBox == ui.Button.OK) {
      tabsAndRange(); }
      else {
        return; }
  }

  function trimActiveMsg() {
  // Displays trim current tasb alert box
  var ui = SpreadsheetApp.getUi();
  var alertBox = ui.alert("Please confirm", "This function will remove unused rows and columns (Leaving a 1 row and column buffer) from the current tab. Be Careful, this is potentially destructive, UNDO (Ctrl+z) can be used twice to undo this function", ui.ButtonSet.OK_CANCEL);

  if (alertBox == ui.Button.OK) {
    trimActive(); }
    else {
      return; }     
  }