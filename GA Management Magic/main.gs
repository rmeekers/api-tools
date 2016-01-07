/* Management Magic for Google Analytics
*   Adds a menu item to manage Google Analytics Properties
*
* Copyright ©2015 Pedro Avila (pdro@google.com)
**************************************************************************/


/**************************************************************************
* Main function runs on application open, setting the menu of commands
*/
function onOpen(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // create the addon menu
  try {
    var menu = ui.createAddonMenu();
    if (e && e.authMode == ScriptApp.AuthMode.NONE) {
      // Add a normal menu item (works in all authorization modes).
      menu.addItem('List filters', 'requestFilterList')
      .addItem('Update filters', 'requestFilterUpdate')
      .addSeparator()
      .addItem('List custom dimensions', 'requestCDList')
      .addItem('Update custom dimensions', 'requestCDUpdate')
      .addSeparator()
      .addItem('List custom metrics', 'requestMetricList')
      .addItem('Update custom metrics', 'requestMetricUpdate')
      .addSeparator()
      .addItem('About this Add-on','about');
    } else {
      menu.addItem('List filters', 'requestFilterList')
      .addItem('Update filters', 'requestFilterUpdate')
      .addSeparator()
      .addItem('List custom dimensions', 'requestCDList')
      .addItem('Update custom dimensions', 'requestCDUpdate')
      .addSeparator()
      .addItem('List custom metrics', 'requestMetricList')
      .addItem('Update custom metrics', 'requestMetricUpdate')
      .addSeparator()
      .addItem('About this Add-on','about');
    }
    menu.addToUi();
     
  } catch (e) {
    Browser.msgBox(e.message);
  }
}


/**************************************************************************
* Install function runs when the application is installed
*/
function onInstall(e) {
  onOpen(e);
}

/**
* Shows the side bar populated with the content from the instructions page
*/
function about() {
  var html = HtmlService.createHtmlOutputFromFile('about')
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setTitle('About')
  .setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

// http://stackoverflow.com/a/2117523/1027723
function generateUUID_(){
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random()*16|0, v = c == 'x' ? r : (r&0x3|0x8);
    return v.toString(16);
  });
}
