// Menu script for Shools

function onOpen() {
    // Create Menu
    SpreadsheetApp.getUi()
                  .createMenu("Shools")
                  .addItem("Hide All But Current", "hideAll")
                  .addItem("Unhide All", "unHide")
                  .addSeparator()
                  .addItem("Count Cells", "countMsg")
                  .addSeparator()
                  .addItem("Trim All Tab(s)", "trimMsg")
                  .addItem("Trim Current Tab", "trimActiveMsg")
                  .addSeparator()
                  .addItem("Tab Name(s) and Active Range(s)", "tabsMsg")
                  .addSeparator()
                  .addItem("List to Tabs", "listToTabs")
                  .addToUi();
    }