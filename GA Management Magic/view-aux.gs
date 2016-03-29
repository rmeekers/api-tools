/* Management Magic for Google Analytics
*    Auxiliary functions for View Management
*    https://developers.google.com/analytics/devguides/config/mgmt/v3/mgmtReference/management/profiles#resource
*
* Copyright Rutger Meekers (rutger@meekers.eu)
***************************************************************************/


/**************************************************************************
* Adds a formatted sheet to the spreadsheet to faciliate data management.
*/
function formatViewsSheet(createNew) {
  // Get common values
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var date = new Date();
  var sheetName = "Views@"+ date.getTime();
  
  // Normalize/format the values of the parameters
  createNew = (createNew === undefined) ? false : createNew;
  
  // Insert a new sheet or warn the user that formatting will erase data on the current sheet
  try {
    if (createNew) {
      sheet = ss.insertSheet(sheetName, 0);
    } else if (!createNew) {
      // Show warning to user and ask to proceed
      var response = ui.alert("WARNING: This will erase all data on the current sheet", "Would you like to proceed?", ui.ButtonSet.YES_NO);
      if (response == ui.Button.YES) {
        sheet.setName(sheetName);
        Logger.log('The user clicked YES.');
      } else if (response == ui.Button.NO) {
        ui.alert('Format cancelled.');
        Logger.log('The user clicked NO.');
        return sheet;
      } else {
        Logger.log('The user clicked the close button in the dialog\'s title bar.');
        return sheet;
      }
    }
  } catch (error) {
    Browser.msgBox(error.message);
    return sheet;
  }
  
  // set local vars
  var cols = 14;
  var numRows = sheet.getMaxRows();
  var numCols = sheet.getMaxColumns();
  var deltaCols = numCols - cols;
  
  // set the number of columns
  try {
    if (deltaCols > 0) {
      sheet.deleteColumns(cols, deltaCols);
    } else if (deltaCols < 0) {
      sheet.insertColumnsAfter(numCols, -deltaCols);
    }
  } catch (e) {
    return "failed to set the number of columns\n"+ e.message;
  }
  
  var includeCol = sheet.getRange("A2:A");
  var botFilteringEnabledCol = sheet.getRange("D2:D");
  var currencyCol = sheet.getRange("E2:E");
  var enhancedECommerceTrackingCol = sheet.getRange("F2:F");
  var stripSiteSearchCategoryParametersCol = sheet.getRange("J2:J");
  var stripSiteSearchQueryParametersCol = sheet.getRange("K2:K");
  var timezoneCol = sheet.getRange("L2:L");
  var typeCol = sheet.getRange("M2:M");
  
  // set header values and formatting
  try {
    var headerRange = sheet.getRange(1,1,1,sheet.getMaxColumns());
    ss.setNamedRange("header_row", headerRange);
    sheet.getRange("A1").setValue("Include");
    sheet.getRange("B1").setValue("webPropertyId");
    sheet.getRange("C1").setValue("name");
    sheet.getRange("D1").setValue("botFilteringEnabled");
    sheet.getRange("E1").setValue("currency");
    sheet.getRange("F1").setValue("enhancedECommerceTracking");
    sheet.getRange("G1").setValue("excludeQueryParameters");
    sheet.getRange("H1").setValue("siteSearchCategoryParameters");
    sheet.getRange("I1").setValue("siteSearchQueryParameters");
    sheet.getRange("J1").setValue("stripSiteSearchCategoryParameters");
    sheet.getRange("K1").setValue("stripSiteSearchQueryParameters");
    sheet.getRange("L1").setValue("timezone");
    sheet.getRange("M1").setValue("type");
    sheet.getRange("N1").setValue("websiteUrl");
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4285F4");
    headerRange.setFontColor("#FFFFFF");
    
    // Include Column: modify data validation values
    var includeValues = ['✓'];
    var includeRule = SpreadsheetApp.newDataValidation().requireValueInList(includeValues, true).build();
    includeCol.setDataValidation(includeRule);

    // currency Column: modify data validation values
    var currencyValues = ['USD','AED','ARS','AUD','BGN','BOB','BRL','CAD','CHF','CLP','CNY','COP','CZK','DKK','EGP','EUR','FRF','GBP','HKD','HRK','HUF','IDR','ILS','INR','JPY','KRW','LTL','MAD','MXN','MYR','NOK','NZD','PEN','PHP','PKR','PLN','RON','RSD','RUB','SAR','SEK','SGD','THB','TRY','TWD','UAH','VEF','VND','ZAR'];
    var currencyRule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(currencyValues, true).build();
    currencyCol.setDataValidation(currencyRule);

    // timezone Column: modify data validation values
    // timezone database values https://en.wikipedia.org/wiki/List_of_tz_database_time_zones
    var timezoneValues = ['AD','AE','AF','AG','AI','AL','AM','AO','AQ','AR','AS','AT','AU','AW','AX','AZ','BA','BB','BD','BE','BF','BG','BH','BI','BJ','BL','BM','BN','BO','BQ','BR','BS','BT','BW','BY','BZ','CA','CC','CD','CF','CG','CH','CI','CK','CL','CM','CN','CO','CR','CU','CV','CW','CX','CY','CZ','DE','DJ','DK','DM','DO','DZ','EC','EE','EG','EH','ER','ES','ET','FI','FJ','FK','FM','FO','FR','GA','GB','GD','GE','GF','GG','GH','GI','GL','GM','GN','GP','GQ','GR','GS','GT','GU','GW','GY','HK','HN','HR','HT','HU','ID','IE','IL','IM','IN','IO','IQ','IR','IS','IT','JE','JM','JO','JP','KE','KG','KH','KI','KM','KN','KP','KR','KW','KY','KZ','LA','LB','LC','LI','LK','LR','LS','LT','LU','LV','LY','MA','MC','MD','ME','MF','MG','MH','MK','ML','MM','MN','MO','MP','MQ','MR','MS','MT','MU','MV','MW','MX','MY','MZ','NA','NC','NE','NF','NG','NI','NL','NO','NP','NR','NU','NZ','OM','PA','PE','PF','PG','PH','PK','PL','PM','PN','PR','PS','PT','PW','PY','QA','RE','RO','RS','RU','RW','SA','SB','SC','SD','SE','SG','SH','SI','SJ','SK','SL','SM','SN','SO','SR','SS','ST','SV','SX','SY','SZ','TC','TD','TF','TG','TH','TJ','TK','TL','TM','TN','TO','TR','TT','TV','TW','TZ','UA','UG','UM','US','UY','UZ','VA','VC','VE','VG','VI','VN','VU','WF','WS','YE','YT','ZA','ZM','ZW'];
    var timezoneRule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(timezoneValues, true).build();
    timezoneCol.setDataValidation(timezoneRule);

    // type Column: modify data validation values
    var typeValues = ['WEB','APP'];
    var typeRule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(typeValues, true).build();
    typeCol.setDataValidation(typeRule);
    
    // Boolean Columns: modify data validation values
    var booleanValues = ['TRUE','FALSE'];
    var booleanRule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(booleanValues, true).build();
    botFilteringEnabledCol.setDataValidation(booleanRule);
    enhancedECommerceTrackingCol.setDataValidation(booleanRule);
    stripSiteSearchCategoryParametersCol.setDataValidation(booleanRule);
    stripSiteSearchQueryParametersCol.setDataValidation(booleanRule);

  } catch (e) {
    return "failed to set the header values and format ranges\n"+ e.message;
  }
  
  return sheet;
}