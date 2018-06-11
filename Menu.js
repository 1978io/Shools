// Menu script for Shools

function onOpen() {
    // Create Menu
    SpreadsheetApp.getUi()
                  .createMenu("Shools")
                  .addItem("Hide All But Current", "hideAll")
                  .addItem("Unhide All", "unHide")
                  .addSeparator()
                  .addItem("Count Cells", "countMsg")
                  .addItem("Trim Sheet(s)", "trimMsg")
                  .addSeparator()
                  .addItem("Tab Name(s) and Active Range(s)", "tabsMsg")
                  .addToUi();
    }