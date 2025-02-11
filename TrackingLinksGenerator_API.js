// Function to call getAccountApps, getLinkDomains, and setupAccountDetails sequentially
function initializeAccountDetails() {
  deleteAllTriggers();  // Ensure all existing triggers are deleted
  getAccountApps();
  getLinkDomains();
  Utilities.sleep(2000); // Adding a delay to ensure each function completes before the next starts
  setupAccountDetails();
  setupDynamicHeaders();
  addOnEditTrigger();  // Add the onEdit trigger
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

// Function to delete all triggers
function deleteAllTriggers() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

// Function to ensure the onEdit trigger is only added once
function addOnEditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = triggers.some(function(trigger) {
    return trigger.getHandlerFunction() === 'onEdit' && trigger.getEventType() === ScriptApp.EventType.ON_EDIT;
  });

  if (!triggerExists) {
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

// Function to fetch account apps from the API and write to "Account Data" sheet starting from the second row
function getAccountApps() {
  try {
    // Get the API key from "Account Details" sheet
    var accountSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account Details");
    var apiKey = accountSheet.getRange("E1").getValue();
    
    // Construct the URL with the API key
    var apiUrl = "https://api.singular.net/api/v1/singular_links/apps?api_key=" + apiKey;
    
    // Fetch data from the API
    var response = UrlFetchApp.fetch(apiUrl, {
      method: "get"
    });
    
    // Parse the JSON response
    var responseData = JSON.parse(response.getContentText());
    
    // Get the "Account Data" sheet
    var accountDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account Data");
    
    // Clear any existing data in the sheet before writing new data
    accountDataSheet.clear();
    
    // Define the sequence of headers
    var headers = ["app", "app_id", "app_platform", "app_longname", "app_site_id", "site_public_id", "store_url"];
    
    // Set the headers in the first row
    accountDataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format the headers
    var headerRange = accountDataSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1434A4"); // Set background color to blue
    headerRange.setFontColor("#FFFFFF"); // Set font color to white
    headerRange.setFontWeight("bold"); // Set font to bold

    // Iterate over each app and print respective values under headers
    var startRow = 2; // Start writing data from the second row
    responseData.available_apps.forEach(function(app) {
      var appData = [];
      headers.forEach(function(header) {
        appData.push(app[header] || ""); // Push value or empty string if not present
      });
      accountDataSheet.getRange(startRow, 1, 1, appData.length).setValues([appData]);
      startRow++; // Move to the next row for the next app
    });
    
  } catch (error) {
    // Handle errors
    Logger.log("Error: " + error.message);
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

// Function to fetch link domains from the API and write to "Account Data" sheet starting from column I
function getLinkDomains() {
  try {
    // Get the API key from "Account Details" sheet
    var accountSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account Details");
    var apiKey = accountSheet.getRange("E1").getValue();
    
    // Construct the URL with the API key
    var apiUrl = "https://api.singular.net/api/v1/singular_links/domains?api_key=" + apiKey;
    
    // Fetch data from the API
    var response = UrlFetchApp.fetch(apiUrl, {
      method: "get"
    });
    
    // Parse the JSON response
    var responseData = JSON.parse(response.getContentText());
    
    // Get the "Account Data" sheet
    var accountDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account Data");
    
    // Clear any existing data in columns I and J before writing new data
    accountDataSheet.getRange("I:J").clear();
    
    // Define the sequence of headers
    var headers = ["subdomain", "dns_zone"];
    
    // Set the headers in the first row of columns I and J
    accountDataSheet.getRange(1, 9, 1, headers.length).setValues([headers]);
    
    // Format the headers
    var headerRange = accountDataSheet.getRange(1, 9, 1, headers.length);
    headerRange.setBackground("#1434A4"); // Set background color to blue
    headerRange.setFontColor("#FFFFFF"); // Set font color to white
    headerRange.setFontWeight("bold"); // Set font to bold

    // Iterate over each domain and print respective values under headers
    var startRow = 2; // Start writing data from the second row
    responseData.available_domains.forEach(function(domain) {
      var domainData = [];
      headers.forEach(function(header) {
        domainData.push(domain[header] || ""); // Push value or empty string if not present
      });
      accountDataSheet.getRange(startRow, 9, 1, domainData.length).setValues([domainData]);
      startRow++; // Move to the next row for the next domain
    });
    
  } catch (error) {
    // Handle errors
    Logger.log("Error: " + error.message);
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

// Function to set up the Account Details sheet with drop-down menus and clear specific cells
function setupAccountDetails() {
  // Set up the drop-down menus for App and subdomain
  createDropdowns();

  // Clear specific cells E5, F5, E7, and A11:E12
  clearManualSelectionCells();
  clearOldData(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Account Details'));

  // Add the onEdit trigger to handle automatic data population when cells are edited
  ScriptApp.newTrigger("onEditHandler")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

// Function to create drop-down menus for App and subdomain
function createDropdowns() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var appDetailsSheet = spreadsheet.getSheetByName('Account Data');
  var accountDetailsSheet = spreadsheet.getSheetByName('Account Details');
  
  // Get unique list of Apps from Account Data sheet
  var appData = appDetailsSheet.getRange("A2:A").getValues().flat().filter(String);
  
  // Create a data validation (drop-down) rule for Apps
  var appValidation = SpreadsheetApp.newDataValidation().requireValueInList(appData, true).build();
  
  // Apply the drop-down rule to cell E7 in the Account Details sheet
  accountDetailsSheet.getRange("E7").setDataValidation(appValidation);

  // Get unique list of subdomains from Account Data sheet
  var accountDataSheet = spreadsheet.getSheetByName('Account Data');
  var subdomainData = accountDataSheet.getRange("I2:I").getValues().flat().filter(String);
  
  // Create a data validation (drop-down) rule for subdomains
  var subdomainValidation = SpreadsheetApp.newDataValidation().requireValueInList(subdomainData, true).build();
  
  // Apply the drop-down rule to cell E5 in the Account Details sheet
  accountDetailsSheet.getRange("E5").setDataValidation(subdomainValidation);
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Handles the onEdit event for the "Account Details" sheet.
 * Updates and validates data based on edits in specific cells.
 *
 * @param {Object} e - The event object containing information about the edit.
 */
function onEditHandler(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var sheetName = sheet.getName();

  // Check if the edited sheet is Account Details
  if (sheetName === "Account Details") {
    // Handle edits in cell E7 (App selection)
    if (range.getA1Notation() === "E7") {
      // Clear old data from the relevant cells
      clearOldData(sheet);

      // Get the selected App
      var app = sheet.getRange("E7").getValue();
      
      // If an App is selected, populate relevant data for that App
      if (app) {
        var appDetailsSheet = e.source.getSheetByName("Account Data");
        var appDetailsData = appDetailsSheet.getDataRange().getValues();
        
        var platforms = [];
        var androidBundleIds = [];
        var iosBundleIds = [];
        var iosRow = false;
        var androidRow = false;
        
        // Iterate through the Account Data data to find matching records
        for (var i = 1; i < appDetailsData.length; i++) {
          if (appDetailsData[i][0] === app) {
            // Collect platform information
            if (!platforms.includes(appDetailsData[i][2])) {
              platforms.push(appDetailsData[i][2]);
            }
            // Handle Android platform
            if (appDetailsData[i][2] === "android") {
              androidRow = true;
              androidBundleIds.push(appDetailsData[i][3]);
              sheet.getRange("A11").setValue("android");
              sheet.getRange("B11").setValue(appDetailsData[i][3]); // Bundle ID
              sheet.getRange("C11").setValue(appDetailsData[i][1]); // App ID
              sheet.getRange("D11").setValue(appDetailsData[i][4]); // App Site ID
              sheet.getRange("E11").setValue(appDetailsData[i][6]); // Store URL
            } 
            // Handle iOS platform
            else if (appDetailsData[i][2] === "ios") {
              iosRow = true;
              iosBundleIds.push(appDetailsData[i][3]);
              sheet.getRange("A12").setValue("ios");
              sheet.getRange("B12").setValue(appDetailsData[i][3]); // Bundle ID
              sheet.getRange("C12").setValue(appDetailsData[i][1]); // App ID
              sheet.getRange("D12").setValue(appDetailsData[i][4]); // App Site ID
              sheet.getRange("E12").setValue(appDetailsData[i][6]); // Store URL
            }
          }
        }

        // Create data validation (drop-down) rules for bundle IDs
        var androidBundleIdValidation = SpreadsheetApp.newDataValidation().requireValueInList(androidBundleIds, true).build();
        var iosBundleIdValidation = SpreadsheetApp.newDataValidation().requireValueInList(iosBundleIds, true).build();

        // Apply drop-down rule to Android row if available
        if (androidRow) {
          sheet.getRange("B11").setDataValidation(androidBundleIdValidation);
        } else {
          sheet.getRange("A11:E11").clearContent();
        }

        // Apply drop-down rule to iOS row if available
        if (iosRow) {
          sheet.getRange("B12").setDataValidation(iosBundleIdValidation);
        } else {
          sheet.getRange("A12:E12").clearContent();
        }
      }
    } 
    // Handle edits in cells B11 or B12 (Bundle ID selection)
    else if (range.getA1Notation() === "B11" || range.getA1Notation() === "B12") {
      var app = sheet.getRange("E7").getValue();
      var platform = range.getRow() === 11 ? "android" : "ios";
      var bundleId = range.getValue();
      
      // If App, platform, and bundle ID are selected, populate relevant data
      if (app && platform && bundleId) {
        var appDetailsSheet = e.source.getSheetByName("Account Data");
        var appDetailsData = appDetailsSheet.getDataRange().getValues();
        
        for (var i = 1; i < appDetailsData.length; i++) {
          if (appDetailsData[i][0] === app && appDetailsData[i][2] === platform && appDetailsData[i][3] === bundleId) {
            var row = platform === "android" ? 11 : 12;
            sheet.getRange("C" + row).setValue(appDetailsData[i][1]); // App ID
            sheet.getRange("D" + row).setValue(appDetailsData[i][4]); // App Site ID
            sheet.getRange("E" + row).setValue(appDetailsData[i][6]); // Store URL
            break;
          }
        }
      }
    } 
    // Handle edits in cell E5 (subdomain selection)
    else if (range.getA1Notation() === "E5") {
      var subdomain = sheet.getRange("E5").getValue();
      
      // If a subdomain is selected, populate the DNS Zone
      if (subdomain) {
        var accountDataSheet = e.source.getSheetByName("Account Data");
        var linkDomainsData = accountDataSheet.getDataRange().getValues();
        
        for (var j = 1; j < linkDomainsData.length; j++) {
          if (linkDomainsData[j][8] === subdomain) {
            sheet.getRange("F5").setValue(linkDomainsData[j][9]); // DNS Zone
            break;
          }
        }
      }
    }
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

// Function to clear old data from specific cells
function clearOldData(sheet) {
  sheet.getRange("A11:E12").clearContent(); // Clear content from A11 to E12
  sheet.getRange("B15:B15").clearContent(); // Clear content for B15
  sheet.getRange("B16:B16").clearContent(); // Clear content for B16
  sheet.getRange("B19:B19").clearContent(); // Clear content for B19
  sheet.getRange("B11").clearDataValidations(); // Clear data validations from B11
  sheet.getRange("B12").clearDataValidations(); // Clear data validations from B12
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

// Function to clear manual selection cells (E5, F5, E7)
function clearManualSelectionCells() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Account Details');
  sheet.getRange("E5").clearContent(); // Clear content from E5
  sheet.getRange("F5").clearContent(); // Clear content from F5
  sheet.getRange("E7").clearContent(); // Clear content from E7
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Sets up dynamic headers in the "Create Links" sheet based on checkbox selections.
 * 
 * This function performs the following actions:
 * 1. Retrieves the "Create Links" sheet and initializes the column index for headers.
 * 2. Clears the existing content and formatting from the header row and data columns below.
 * 3. Iterates over a predefined list of headers associated with checkboxes, and sets the headers in row 7 based on the checkbox values.
 * 4. Applies formatting to the dynamic headers including background color, font color, font weight, horizontal and vertical alignment, and text wrap.
 * 5. Adds static headers after the dynamic headers and applies similar formatting.
 * 
 * The headers and corresponding checkboxes are defined as:
 * - A2: "Deep Link (Android & iOS)"
 * - A3: "Deferred Deep Link (Android & iOS)"
 * - A4: "Deep Link (Android)"
 * - A5: "Deferred Deep Link (Android)"
 * - C2: "Deep Link (iOS)"
 * - C3: "Deferred Deep Link (iOS)"
 * - C4: "Campaign Name"
 * - C5: "Campaign ID"
 * - E2: "Sub Campaign Name"
 * - E3: "Sub Campaign ID"
 * - E4: "Fallback URL Web"
 * - E5: "Passthrough"
 * 
 * Static headers are:
 * - "Response Code"
 * - "Result"
 * - "Tracking_Link Name"
 * - "Short link"
 * - "Long Link"
 * - "Response data"
 */
// Function to add headers dynamically based on checkbox selections and add static headers
function setupDynamicHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create Links");
  const headerRow = 7;
  const startColumn = 3; // Column C

  const headers = [
    { checkbox: "A2", header: "Deep Link (Android & iOS)" },
    { checkbox: "A3", header: "Deferred Deep Link (Android & iOS)" },
    { checkbox: "A4", header: "Deep Link (Android)" },
    { checkbox: "A5", header: "Deferred Deep Link (Android)" },
    { checkbox: "C2", header: "Deep Link (iOS)" },
    { checkbox: "C3", header: "Deferred Deep Link (iOS)" },
    { checkbox: "C4", header: "Campaign Name" },
    { checkbox: "C5", header: "Campaign ID" },
    { checkbox: "E2", header: "Sub Campaign Name" },
    { checkbox: "E3", header: "Sub Campaign ID" },
    { checkbox: "E4", header: "Fallback URL Web" },
    { checkbox: "E5", header: "Passthrough" }
  ];

  const staticHeaders = ["Response Code", "Result", "Tracking_Link Name", "Short link", "Long Link", "Response data", "QR Code URL"];

  // Initialize the column index for headers
  let columnIndex = startColumn;

  // Clear the entire header row and data columns below
  sheet.getRange(headerRow, startColumn, 1, sheet.getMaxColumns() - startColumn + 1).clearContent().clearFormat();
  sheet.getRange(headerRow + 1, startColumn, sheet.getMaxRows() - headerRow, sheet.getMaxColumns() - startColumn + 1).clearContent();

  // Iterate over the headers array and set headers based on checkbox selections
  headers.forEach(headerObj => {
    const checkboxValue = sheet.getRange(headerObj.checkbox).getValue();
    if (checkboxValue) {
      // Set the header in row 7
      const headerCell = sheet.getRange(headerRow, columnIndex);
      headerCell.setValue(headerObj.header);
      // Apply formatting for dynamic headers
      headerCell.setBackground("#cfe2f3").setFontColor("#000000").setFontWeight("bold");
      // Apply alignment and text wrap
      headerCell.setHorizontalAlignment("center").setVerticalAlignment("middle").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      columnIndex++;
    }
  });

  // Add static headers after the dynamic headers
  staticHeaders.forEach(staticHeader => {
    const staticHeaderCell = sheet.getRange(headerRow, columnIndex);
    staticHeaderCell.setValue(staticHeader);
    // Apply formatting for static headers
    staticHeaderCell.setBackground("#d9d2e9").setFontColor("#000000").setFontWeight("bold");
    // Apply alignment and text wrap
    staticHeaderCell.setHorizontalAlignment("center").setVerticalAlignment("middle").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    columnIndex++;
  });
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Handles the onEdit event for the "Create Links" sheet.
 * 
 * This function performs the following actions when a checkbox is edited:
 * 1. Checks if the edited cell is within the specified checkbox range.
 * 2. If a checkbox is deselected (value is false), clears the corresponding column data and formatting.
 * 3. Calls `setupDynamicHeaders()` to re-setup the dynamic headers based on the changed checkboxes.
 * 
 * @param {Object} e - The event object containing information about the edit.
 * @property {Object} e.range - The range that was edited.
 * @property {Object} e.range.getSheet() - The sheet that contains the edited range.
 * @property {String} e.range.getA1Notation() - The A1 notation of the edited cell.
 * @property {any} e.range.getValue() - The new value of the edited cell.
 */
// Function to handle changes in checkboxes and adjust headers dynamically
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const headerRow = 7;
  const startColumn = 3; // Column C

  // Check if the edited range is within the checkboxes
  if (sheet.getName() === "Create Links") {
    const headers = [
      "A2", "A3", "A4", "A5", "C2", "C3", "C4", "C5", "E2", "E3", "E4", "E5"
    ];
    if (headers.includes(range.getA1Notation())) {
      // Clear the column data if the checkbox is deselected
      if (!range.getValue()) {
        const headerIndex = headers.indexOf(range.getA1Notation());
        const columnToClear = startColumn + headerIndex;
        sheet.getRange(headerRow, columnToClear).clearContent().clearFormat(); // Clear the header and its formatting
        sheet.getRange(headerRow + 1, columnToClear, sheet.getMaxRows() - headerRow, 1).clearContent(); // Clear the column data
      }
      
      // Re-setup the dynamic headers based on the changed checkboxes
      setupDynamicHeaders();
    }
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Creates Singular links based on the data in the "Create Links" sheet and other related sheets.
 * 
 * This function performs the following steps:
 * 1. Retrieves necessary sheets and configuration values.
 * 2. Clears any existing response headers in the "Create Links" sheet.
 * 3. Processes rows of data in batches, as defined by the batch size.
 * 4. Validates row data for required fields and formats.
 * 5. If validation errors are found, logs them in the "Result" column of the "Create Links" sheet.
 * 6. If no validation errors are present, processes the row to create Singular links.
 * 7. Pauses between batches to avoid hitting API rate limits.
 * 
 * @throws {Error} Throws an error if there is an issue during processing.
 */
function createSingularLinks() {
  try {
    // Get the active spreadsheet and the required sheets
    const createLinksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create Links");
    const accountDetailsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account Details");

    // Define the header row and the start index for processing (row 8)
    const headerRow = 7; // Headers are in row 7
    const startIndex = headerRow + 1; // Start processing from row 8

    // Get the batch size from cell C27 of the Account Details sheet
    const batchSize = parseInt(accountDetailsSheet.getRange("C27").getValue()); // BATCH_SIZE from cell C27

    // Get the authorization value from cell E1 of the Account Details sheet
    const authorization = accountDetailsSheet.getRange("E1").getValue(); // Authorization from cell E1

    // Clear any existing response headers in the Create Links sheet before starting
    clearResponseHeaders(createLinksSheet, headerRow);

    // Get the data range starting from row 8 and including all columns
    const dataRange = createLinksSheet.getRange(startIndex, 1, createLinksSheet.getLastRow() - headerRow, createLinksSheet.getLastColumn());

    // Retrieve all the values from the data range
    const values = dataRange.getValues();

    // Initialize an array to hold valid rows for processing
    const validRows = [];

    // Validate each row and collect valid rows
    for (let i = 0; i < values.length; i++) {
      const rowData = values[i];
      const rowIndex = i + startIndex;

      // Validate row data and collect errors
      let validationErrors = validateRowData(rowData, rowIndex, createLinksSheet, accountDetailsSheet);

      // If no validation errors, add row data to validRows array
      if (!validationErrors) {
        validRows.push({ rowData, rowIndex });
      } else {
        // Write validation errors to the "Result" column
        createLinksSheet.getRange(rowIndex, getHeaderIndex(createLinksSheet, "Result")).setValue(validationErrors.trim());
      }
    }

    // Process the valid rows in batches
    let currentIndex = 0;
    while (currentIndex < validRows.length) {
      for (let i = currentIndex; i < currentIndex + batchSize && i < validRows.length; i++) {
        const { rowData, rowIndex } = validRows[i];
        processRow(rowData, rowIndex, createLinksSheet, authorization);
      }

      // Pause for 1 minute before continuing to avoid hitting API limits
      if (currentIndex + batchSize < validRows.length) {
        Utilities.sleep(60000); // Wait for one minute before processing the next batch
      }

      // Move the current index forward by the batch size
      currentIndex += batchSize;
    }
  } catch (error) {
    // Log any errors encountered during the process
    Logger.log("Error: " + error.message);
  }
}

function validateRowData(rowData, rowIndex, createLinksSheet, accountDetailsSheet) {
  // Implement your validation logic here
  // Return validation error messages as a string or return null if there are no errors
  let validationErrors = "";

  const appId = getCellValue(accountDetailsSheet, "C11") || getCellValue(accountDetailsSheet, "C12") || "data not available";
  const sourceName = rowData[0];
  const trackingLinkName = rowData[1];
  const linkSubdomain = getCellValue(accountDetailsSheet, "E5");
  const fallbackUrl = getFallbackUrl(createLinksSheet, accountDetailsSheet, rowIndex);
  const clickDeterministicWindow = getCellValue(accountDetailsSheet, "C22");
  const clickProbabilisticWindow = getCellValue(accountDetailsSheet, "C23");
  const viewDeterministicWindow = getCellValue(accountDetailsSheet, "C24");
  const viewProbabilisticWindow = getCellValue(accountDetailsSheet, "C25");
  const clickDeterministicReEngWindow = getCellValue(accountDetailsSheet, "C26");

  // Validate App ID
  if (appId === "data not available") {
    validationErrors += "App ID is missing. ";
  }

  // Validate Source Name against acceptable values
  if (!["social", "email", "sms", "crosspromo"].includes(sourceName.toLowerCase())) {
    validationErrors += 'Source Name is incorrect (must be one of "Social", "Email", "SMS", "Crosspromo"). ';
  }

  // Validate Tracking Link Name
  if (!trackingLinkName) {
    validationErrors += "Tracking Link Name is missing. ";
  }

  // Validate Link Subdomain
  if (!linkSubdomain) {
    validationErrors += "Link Subdomain is missing. ";
  }

  // Validate Fallback URL to ensure it starts with http or https
  if (!fallbackUrl || !(fallbackUrl.startsWith("http://") || fallbackUrl.startsWith("https://"))) {
    validationErrors += "Fallback URL is missing or incorrect should start with http or https. ";
  }

  // Validate Click Deterministic Window to ensure it's a number within the range 0-30
  if (!clickDeterministicWindow || isNaN(clickDeterministicWindow) || clickDeterministicWindow < 0 || clickDeterministicWindow > 30) {
    validationErrors += "Click Deterministic Window is missing, not a number, or out of the valid range (0-30). ";
  }

  // Validate Click Probabilistic Window to ensure it's a number within the range 0-24
  if (!clickProbabilisticWindow || isNaN(clickProbabilisticWindow) || clickProbabilisticWindow < 0 || clickProbabilisticWindow > 24) {
    validationErrors += "Click Probabilistic Window is missing, not a number, or out of the valid range (0-24). ";
  }

  // Validate View Deterministic Window to ensure it's a number within the range 0-24
  if (!viewDeterministicWindow || isNaN(viewDeterministicWindow) || viewDeterministicWindow < 0 || viewDeterministicWindow > 24) {
    validationErrors += "View Deterministic Window is missing, not a number, or out of the valid range (0-24). ";
  }

  // Validate View Probabilistic Window to ensure it's a number within the range 0-24
  if (!viewProbabilisticWindow || isNaN(viewProbabilisticWindow) || viewProbabilisticWindow < 0 || viewProbabilisticWindow > 24) {
    validationErrors += "View Probabilistic Window is missing, not a number, or out of the valid range (0-24). ";
  }

  // Validate Click Deterministic Re-Engagement Window to ensure it's a number within the range 0-30
  if (!clickDeterministicReEngWindow || isNaN(clickDeterministicReEngWindow) || clickDeterministicReEngWindow < 0 || clickDeterministicReEngWindow > 30) {
    validationErrors += "Click Deterministic ReEngagement Window is missing, not a number, or out of the valid range (0-30). ";
  }

  return validationErrors || null;
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Clears the content of specific response header columns in the sheet.
 * 
 * This function is used to remove the content from columns that are designated for storing 
 * response data from API calls. It ensures that any previous data is cleared before new data 
 * is processed and added to the sheet.
 * 
 * @param {Sheet} sheet - The sheet where the headers are located and where the content needs to be cleared.
 * @param {number} headerRow - The row number where the headers are located.
 */
function clearResponseHeaders(sheet, headerRow) {
  // Define the headers that need to be cleared
  const headers = ["Response Code", "Result", "Tracking_Link Name", "Short link", "Long Link", "Response data", "QR Code URL"];
  
  // Iterate over each header in the headers array
  headers.forEach(header => {
    // Get the column index of the current header
    const columnIndex = getHeaderIndex(sheet, header);
    
    // If the header exists in the sheet (columnIndex is not -1)
    if (columnIndex !== -1) {
      // Get the range starting from the row below the headerRow to the last row in that column
      const range = sheet.getRange(headerRow + 1, columnIndex + 1, sheet.getMaxRows() - headerRow);
      
      // Clear the content of the selected range
      range.clearContent();
    }
  });
}




// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//


function identifyInvalidRows() {
  try {
    //Logger.log("Starting identifyInvalidRows function.");

    const createLinksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create Links");
    if (!createLinksSheet) {
      Logger.log("Error: 'Create Links' sheet not found.");
      return;
    }

    const headerRow = 7;
    const startIndex = headerRow + 1;
    const lastRow = createLinksSheet.getLastRow();
    //Logger.log(`lastRow ${lastRow}.`);
    const lastColumn = createLinksSheet.getLastColumn();

    //Logger.log(`Sheet has ${lastRow - headerRow} rows and ${lastColumn} columns.`);

    const dataRange = createLinksSheet.getRange(startIndex, 1, lastRow - headerRow, lastColumn);
    const values = dataRange.getValues();

    const responseCodeIndex = getHeaderIndex(createLinksSheet, "Response Code");
    if (responseCodeIndex === -1) {
      Logger.log("Error: 'Response Code' column not found.");
      return;
    }

    //Logger.log(`'Response Code' column found at index ${responseCodeIndex + 1}.`);

    let invalidRows = [];

    for (let i = 0; i < values.length; i++) {
      const rowData = values[i];
      const responseCode = rowData[responseCodeIndex];
      const rowIndex = i + startIndex;

      //Logger.log(`Row ${rowIndex}: Response Code = ${responseCode}`);

      if (responseCode != 200 && responseCode != 201) {
      //if (responseCode != 200 && responseCode != 201 && responseCode !== "" && responseCode !== null) {
       // Logger.log(`Row ${rowIndex}: Invalid response code.`);
        invalidRows.push(rowIndex);
      }
    }

    Logger.log(`Invalid Rows: ${invalidRows.length}`);

    // Display the list of invalid rows in a dialog box before proceeding
    if (invalidRows.length > 0) {
      // const ui = SpreadsheetApp.getUi();
      // ui.alert('Invalid Rows Detected', `Rows to Reprocess: ${invalidRows.join(', ')}`, ui.ButtonSet.OK);

      // After the alert is shown, proceed to clear response headers and reprocess invalid rows
      clearResponse(createLinksSheet, headerRow, invalidRows);
      reprocessInvalidRows(invalidRows);
    } else {
      //Logger.log('No invalid rows found.');
    }

  } catch (error) {
    Logger.log("Error in identifyInvalidRows: " + error.message);
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

function clearResponse(sheet, headerRow, invalidRows) {
  // Define the headers that need to be cleared
  const headers = ["Response Code", "Result", "Tracking_Link Name", "Short link", "Long Link", "Response data", "QR Code URL"];
  
  // Iterate over each header in the headers array
  headers.forEach(header => {
    // Get the column index of the current header
    const columnIndex = getHeaderIndex(sheet, header);
    
    // If the header exists in the sheet (columnIndex is not -1)
    if (columnIndex !== -1) {
      // Iterate through the invalid rows and clear the relevant columns
      invalidRows.forEach(row => {
       // Logger.log(`Clearing content in column ${columnIndex + 1} for row ${row}`);
        sheet.getRange(row, columnIndex + 1).clearContent();
      });
    }
  });
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

function getHeaderIndex(sheet, headerName) {
  try {
    const headers = sheet.getRange(7, 1, 1, sheet.getLastColumn()).getValues()[0];
    const index = headers.indexOf(headerName);
    //Logger.log(`Header '${headerName}' found at index ${index + 1}`);
    return index;
  } catch (error) {
    Logger.log(`Error in getHeaderIndex: ${error.message}`);
    return -1;
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

function reprocessInvalidRows(invalidRows) {
  try {
    //Logger.log("Starting reprocessInvalidRows function.");

    // Get the active spreadsheet and the required sheets
    const createLinksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create Links");
    const accountDetailsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account Details");

    if (!createLinksSheet || !accountDetailsSheet) {
      Logger.log("Error: Required sheets not found.");
      return;
    }

    // Define the header row and start index for processing
    const headerRow = 7; // Headers are in row 7
    const responseCodeIndex = getHeaderIndex(createLinksSheet, "Response Code");

    // Validate row data for all invalid rows
    let rowsToBatchProcess = [];
    
    invalidRows.forEach(rowIndex => {
      const rowData = createLinksSheet.getRange(rowIndex, 1, 1, createLinksSheet.getLastColumn()).getValues()[0];

      // Validate row data and collect errors
      let validationErrors = validateRowData(rowData, rowIndex, createLinksSheet, accountDetailsSheet);

      // If no validation errors, check if the responseCode is still empty
      if (!validationErrors) {
        const responseCode = createLinksSheet.getRange(rowIndex, responseCodeIndex + 1).getValue();
        if (!responseCode) {
          rowsToBatchProcess.push(rowIndex);
        }
      } else {
        // Write validation errors to the "Result" column
        createLinksSheet.getRange(rowIndex, getHeaderIndex(createLinksSheet, "Result")).setValue(validationErrors.trim());
      }
    });

    // Process the rows with empty response codes in batches
    if (rowsToBatchProcess.length > 0) {
      batchProcessRows(rowsToBatchProcess, createLinksSheet, accountDetailsSheet);
    }

  } catch (error) {
    Logger.log("Error in reprocessInvalidRows: " + error.message);
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

function batchProcessRows(rowsToBatchProcess, createLinksSheet, accountDetailsSheet) {
  try {
    const batchSize = parseInt(accountDetailsSheet.getRange("C27").getValue()); // BATCH_SIZE from cell C27
    const authorization = accountDetailsSheet.getRange("E1").getValue(); // Authorization from cell E1

    let currentIndex = 0;
    while (currentIndex < rowsToBatchProcess.length) {
      for (let i = currentIndex; i < currentIndex + batchSize && i < rowsToBatchProcess.length; i++) {
        const rowIndex = rowsToBatchProcess[i];
        const rowData = createLinksSheet.getRange(rowIndex, 1, 1, createLinksSheet.getLastColumn()).getValues()[0];

        processRow(rowData, rowIndex, createLinksSheet, authorization);
      }

      // Pause for 1 minute before continuing to avoid hitting API limits
      if (currentIndex + batchSize < rowsToBatchProcess.length) {
        Utilities.sleep(60000); // Wait for one minute before processing the next batch
      }

      currentIndex += batchSize;
    }

  } catch (error) {
    Logger.log("Error in batchProcessRows: " + error.message);
  }
}


// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Processes a single row of data from the sheet, constructs the request body, and sends it to the Singular API.
 * 
 * @param {Array} rowData - An array of values representing the row data from the sheet.
 * @param {number} rowIndex - The index of the row being processed.
 * @param {Sheet} sheet - The sheet object where the row data is located.
 * @param {string} authorization - The authorization token to be used in the API request headers.
 */
function processRow(rowData, rowIndex, sheet, authorization) {
  try {
    // Retrieve the "Account Details" and "Create Links" sheets from the active spreadsheet
    const accountDetailsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account Details");
    const createLinksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create Links");

    // Construct the request body using data from the row, "Account Details", and "Create Links" sheets
    const requestBody = constructRequestBody(rowData, accountDetailsSheet, createLinksSheet, rowIndex);

    // Display the constructed request body in a dialog box for debugging or verification
    //showRequestBody(requestBody);

    // Define the headers for the API request, including authorization and content type
    const headers = {
      "Authorization": authorization,
      "Content-Type": "application/json"
    };

    try {
      // Define the API endpoint URL for creating Singular links
      const url = "https://api.singular.net/api/v1/singular_links/links";

      // Send a POST request to the Singular API with the constructed request body and headers
      const response = UrlFetchApp.fetch(url, {
        "method": "post",
        "headers": headers,
        "payload": JSON.stringify(requestBody)
      });

      // Handle the API response, logging or processing results as necessary
      handleApiResponse(response, rowIndex, rowData, sheet);
    } catch (error) {
      // If an error occurs during the API request, handle it accordingly (e.g., log error, update sheet)
      handleApiError(error, rowIndex, sheet);
    }
  } catch (error) {
    // Log any errors that occur during the processing of the row
    Logger.log("Error: " + error.message);
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Constructs the request body for the Singular API call based on the row data and details from the "Account Details" and "Create Links" sheets.
 *
 * @param {Array} rowData - An array of values representing the row data from the sheet.
 * @param {Sheet} accountDetailsSheet - The sheet object representing "Account Details".
 * @param {Sheet} createLinksSheet - The sheet object representing "Create Links".
 * @param {number} rowIndex - The index of the row being processed.
 * @returns {Object} The constructed request body for the API call.
 */
function constructRequestBody(rowData, accountDetailsSheet, createLinksSheet, rowIndex) {
  // Initialize the request body object with relevant key-value pairs
  let requestBody = {
    // Get the App ID from either cell C11 or C12 in the "Account Details" sheet, or set a default value if not available
    "app_id": getCellValue(accountDetailsSheet, "C11") || getCellValue(accountDetailsSheet, "C12") || "data not available",
    "partner_id": "",  // Set partner ID (can be updated later based on specific needs)
    "link_type": "custom",  // Define the link type as 'custom'
    "source_name": rowData[0],  // Get the source name from the first column of rowData
    // Determine whether re-engagement is enabled, based on the value in cell C28 of "Account Details"
    "enable_reengagement": getReengagementValue(getCellValue(accountDetailsSheet, "C28")),
    "tracking_link_name": rowData[1],  // Get the tracking link name from the second column of rowData
    "link_subdomain": getCellValue(accountDetailsSheet, "E5"),  // Get the link subdomain from cell E5 in "Account Details"
    "link_dns_zone": "sng.link",  // Set the DNS zone to 'sng.link'
    // Get the fallback URL, which is a combination of data from "Create Links" and "Account Details" sheets
    "destination_fallback_url": getFallbackUrl(createLinksSheet, accountDetailsSheet, rowIndex),
    "click_deterministic_window": getCellValue(accountDetailsSheet, "C22"),  // Get the click deterministic window from cell C22
    "click_probabilistic_window": getCellValue(accountDetailsSheet, "C23"),  // Get the click probabilistic window from cell C23
    "View_deterministic_window": getCellValue(accountDetailsSheet, "C24"),  // Get the view deterministic window from cell C24
    "view_probabilistic_window": getCellValue(accountDetailsSheet, "C25"),  // Get the view probabilistic window from cell C25
    "click_reengagement_window": getCellValue(accountDetailsSheet, "C26"),  // Get the click re-engagement window from cell C26
    "enable_ctv": false,  // Set the CTV (Connected TV) option to false by default
    // Construct link parameters using the provided rowData and "Create Links" sheet
    "link_parameter": constructLinkParameter(createLinksSheet, rowData),
    // Construct redirection details for Android and iOS using the respective platform, rowData, and sheets
    "android_redirection": constructRedirection(accountDetailsSheet, createLinksSheet, "android", rowData),
    "ios_redirection": constructRedirection(accountDetailsSheet, createLinksSheet, "ios", rowData)
  };

  // Remove Android redirection from the request body if app_site_id is not available
  if (requestBody.android_redirection && !requestBody.android_redirection.app_site_id) {
    delete requestBody.android_redirection;
  }

  // Remove iOS redirection from the request body if app_site_id is not available
  if (requestBody.ios_redirection && !requestBody.ios_redirection.app_site_id) {
    delete requestBody.ios_redirection;
  }

  // Return the constructed request body object
  return requestBody;
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Determines the value for the 'enable_reengagement' field based on the content of a specific cell.
 *
 * @param {Object} cellValue - The value of the cell from which the re-engagement setting is determined.
 * @returns {boolean} Returns true if the cell value is 'true' (case insensitive), otherwise false.
 */
function getReengagementValue(cellValue) {
  return cellValue !== null && cellValue.toString().toLowerCase() === "true";
}


// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Handles the API response and updates the relevant row in the Google Sheet.
 * 
 * This function checks the response code from an API call and updates specific columns in the 
 * sheet based on whether the API request was successful or not.
 * 
 * If the response is successful (HTTP status 200 or 201), it extracts relevant data from the 
 * JSON response, updates the corresponding columns in the sheet, and appends the success data 
 * to the "Links History" sheet.
 * 
 * If the response indicates an error, it logs the error code and response data in the appropriate 
 * columns of the sheet.
 * 
 * @param {HTTPResponse} response - The response object from the API call.
 * @param {number} rowIndex - The index of the current row being processed in the sheet.
 * @param {Array} rowData - The data from the current row being processed.
 * @param {Sheet} sheet - The sheet where the data and results are being updated.
 */
function handleApiResponse(response, rowIndex, rowData, sheet) {
  // Get the header row values from row 7
  const headerRow = sheet.getRange(7, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find the column indices for each of the needed headers
  const responseCodeIndex = headerRow.indexOf("Response Code") + 1;
  const resultIndex = headerRow.indexOf("Result") + 1;
  const trackingLinkNameIndex = headerRow.indexOf("Tracking_Link Name") + 1;
  const shortLinkIndex = headerRow.indexOf("Short link") + 1;
  const longLinkIndex = headerRow.indexOf("Long Link") + 1;
  const responseDataIndex = headerRow.indexOf("Response data") + 1;

  // Get the response code from the API response
  const responseCode = response.getResponseCode();
  // Set the row number to match the data row
  const rowNumber = rowIndex;

  if (responseCode === 200 || responseCode === 201) {
    // Parse the JSON response data
    const responseData = JSON.parse(response.getContentText());
    const shortLink = responseData.short_link;
    const clickTrackingLink = responseData.click_tracking_link;
    const trackingLinkName = responseData.tracking_link_name;

    // Set the values in the respective columns for the row
    sheet.getRange(rowNumber, responseCodeIndex).setValue(responseCode);
    sheet.getRange(rowNumber, resultIndex).setValue("success");
    sheet.getRange(rowNumber, trackingLinkNameIndex).setValue(trackingLinkName);
    sheet.getRange(rowNumber, shortLinkIndex).setValue(shortLink);
    sheet.getRange(rowNumber, longLinkIndex).setValue(clickTrackingLink);
    sheet.getRange(rowNumber, responseDataIndex).setValue(response.getContentText());

    // Append the successful response data to the "Links History" sheet
    appendToLinksHistory(trackingLinkName, shortLink, clickTrackingLink, response.getContentText());
  } else {
    // If the response code is not 200 or 201, log the error
    sheet.getRange(rowNumber, responseCodeIndex).setValue(responseCode);
    sheet.getRange(rowNumber, resultIndex).setValue("Error");
    sheet.getRange(rowNumber, responseDataIndex).setValue(response.getContentText());
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Appends a new record to the "Links History" sheet with details about the tracking link.
 *
 * @param {string} trackingLinkName - The name of the tracking link.
 * @param {string} shortLink - The shortened version of the link.
 * @param {string} clickTrackingLink - The full click tracking link.
 * @param {string} responseData - The response data from the API call.
 */
function appendToLinksHistory(trackingLinkName, shortLink, clickTrackingLink, responseData) {
  // Get the "Links History" sheet from the active spreadsheet
  const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Links History");
  
  // Get the last row with data in the "Links History" sheet to determine where to append the new data
  const lastRow = historySheet.getLastRow();
  // Calculate the row number for the new data
  const newRow = lastRow + 1;

  // Get the current date to record when the link was added
  const today = new Date();

  // Set the values in the respective columns of the new row in the "Links History" sheet
  historySheet.getRange(newRow, 1).setValue(today);              // Column 1: Date
  historySheet.getRange(newRow, 2).setValue(trackingLinkName);    // Column 2: Tracking Link Name
  historySheet.getRange(newRow, 3).setValue(shortLink);           // Column 3: Short Link
  historySheet.getRange(newRow, 4).setValue(clickTrackingLink);   // Column 4: Long Link
  historySheet.getRange(newRow, 5).setValue(responseData);        // Column 5: Response Data
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Handles errors that occur during the API request, logging the error and updating the sheet with error details.
 *
 * @param {Error} error - The error object thrown during the API request.
 * @param {number} rowIndex - The index of the row being processed when the error occurred.
 * @param {Sheet} sheet - The sheet where the row data is located.
 */
function handleApiError(error, rowIndex, sheet) {
  // Retrieve the header row (assumed to be in the 7th row) to find the indices of specific columns
  const headerRow = sheet.getRange(7, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find the index of the "Response Code" column (adjusted for 1-based indexing)
  const responseCodeIndex = headerRow.indexOf("Response Code") + 1;
  // Find the index of the "Result" column (adjusted for 1-based indexing)
  const resultIndex = headerRow.indexOf("Result") + 1;
  // Find the index of the "Response data" column (adjusted for 1-based indexing)
  const responseDataIndex = headerRow.indexOf("Response data") + 1;

  // Extract the numeric error code from the error message (if available)
  const errorCode = error.message.match(/\d+/);
  // Set the row number to match the row where the error occurred
  const rowNumber = rowIndex;

  // Update the sheet with the error details
  sheet.getRange(rowNumber, responseCodeIndex).setValue(errorCode);  // Set the error code in the "Response Code" column
  sheet.getRange(rowNumber, resultIndex).setValue("Error");          // Mark the "Result" column with "Error"
  sheet.getRange(rowNumber, responseDataIndex).setValue(error.message); // Record the full error message in the "Response data" column

  // Log the error message to the Google Apps Script Logger for debugging
  Logger.log("Error during API request: " + error.message);
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Function to get the value of a specified cell in a given sheet.
 * 
 * @param {Sheet} sheet - The Google Sheets object.
 * @param {string} cell - The cell reference (e.g., "B19").
 * @return {any} The value of the specified cell or null if not found.
 */
function getCellValue(sheet, cell) {
  // Get the value of the specified cell
  const value = sheet.getRange(cell).getValue();
  
  // Return the cell value or null if the value is not present
  return value ? value : null;
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Retrieves the fallback URL to be used in the API request payload.
 * 
 * This function first checks the "Create Links" sheet for a "Fallback URL Web" header and attempts
 * to retrieve the corresponding URL from the specified row.
 * 
 * If the "Fallback URL Web" header is not found or the cell is empty, it falls back to checking 
 * cell B19 in the "Account Details" sheet for a default fallback URL.
 * 
 * The function returns the found URL or a default message "data not available" if no URL is provided.
 * 
 * @param {Sheet} createLinksSheet - The sheet containing the link creation data.
 * @param {Sheet} accountDetailsSheet - The sheet containing account-specific details.
 * @param {number} rowIndex - The index of the current row being processed.
 * @returns {string} The fallback URL to be used, or a default message if none is found.
 */
function getFallbackUrl(createLinksSheet, accountDetailsSheet, rowIndex) {
  // Retrieve the column index of the header "Fallback URL Web" in the "Create Links" sheet
  const fallbackHeaderIndex = getHeaderIndex(createLinksSheet, "Fallback URL Web");
  
  // If the "Fallback URL Web" header is found
  if (fallbackHeaderIndex !== -1) {
    // Get the value from the specific cell in the current row and the column index of "Fallback URL Web"
    const fallbackUrlFromSheet = createLinksSheet.getRange(rowIndex, fallbackHeaderIndex + 1).getValue();
    
    // Check if the retrieved value is empty
    if (fallbackUrlFromSheet) {
      // Return the fallback URL found in the "Create Links" sheet
      return fallbackUrlFromSheet;
    }
  }

  // If the "Fallback URL Web" header is not found or the cell is empty,
  // get the value from the "Account Details" sheet cell B19
  const fallbackUrlFromAccountDetails = getCellValue(accountDetailsSheet, "B19");

  // Return the fallback URL found in the "Account Details" sheet or a default message if not found
  return fallbackUrlFromAccountDetails || "data not available";
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Constructs the `linkParameter` object to be used in the API request payload.
 * 
 * This function retrieves specific campaign-related data from the `createLinksSheet`,
 * mapping it to corresponding parameter keys as defined in `headerMapping`.
 * 
 * Additionally, it checks the value in cell `C29` of the "Account Details" sheet to conditionally
 * add an `_smtype` parameter to the `linkParameter` object.
 * 
 * The function iterates over predefined headers, checking their presence in the sheet, 
 * and adds their associated values to the `linkParameter` object if they are not empty.
 * 
 * Finally, it logs the constructed `linkParameter` object for debugging and returns it.
 * 
 * @param {Sheet} createLinksSheet - The sheet containing the link creation data.
 * @param {Array} rowData - The data for the current row being processed.
 * @returns {Object} The constructed `linkParameter` object.
 */

function constructLinkParameter(createLinksSheet, rowData) {
  // Get the header row data from row 7 of the createLinksSheet
  const headerRow = getRowData(createLinksSheet, "7:7");
  
  // Initialize an empty object to store the link parameters
  let linkParameter = {};

  // Log the header row and row data for debugging purposes
  // Logger.log("Header Row: " + JSON.stringify(headerRow));
  // Logger.log("Row Data: " + JSON.stringify(rowData));

  // Define a mapping between the header names in the sheet and the corresponding link parameter keys
  const headerMapping = {
    "Campaign Name": "pcn",
    "Campaign ID": "pcid",
    "Sub Campaign Name": "pscn",
    "Sub Campaign ID": "pscid",
    "Passthrough": "_p"
  };

  // Iterate over each header defined in the headerMapping object
  for (let header in headerMapping) {
    const param = headerMapping[header];
    const index = headerRow.indexOf(header);

    if (index !== -1) {
      const value = rowData[index].toString().trim();

      // Add to linkParameter only if the value is not empty
      if (value) {
        linkParameter[param] = value;
        //Logger.log(`Set linkParameter.${param} to: ${linkParameter[param]} from header: ${header} at index: ${index}`);
      } else {
        //Logger.log(`Skipped empty value for header: ${header}`);
      }
    } else {
      Logger.log(`Header not found: ${header}`);
    }
  }

  // Check the value in cell C29 of "Account Details" sheet
  const accountDetailsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Account Details");
  const smtypeValue = getCellValue(accountDetailsSheet, "C29");

  // If the value is "true" (or empty/null), add "_smtype": "3" to linkParameter
  if (smtypeValue === null || smtypeValue.toString().toLowerCase() === "true" || smtypeValue === "") {
    linkParameter["_smtype"] = "3";
    //Logger.log('Added "_smtype": "3" to linkParameter based on C29 value.');
  }

  // Log the final constructed linkParameter object for debugging
  //Logger.log("Constructed linkParameter: " + JSON.stringify(linkParameter));
  
  // Return the constructed linkParameter object
  return linkParameter;
}

//-----------------------------------------------------------------------------------------------------------------//
//-----------------------------------------------------------------------------------------------------------------//

/**
 * Constructs a redirection object for either Android or iOS platforms based on the data
 * provided in the Google Sheets.
 *
 * This function builds the redirection object, which includes app site IDs, destination URLs,
 * deep link URLs, and deferred deep link URLs for a specific platform (Android or iOS). It 
 * checks for the presence of data in specific cells and headers in the "Account Details" and 
 * "Create Links" sheets. If the necessary data is not found, the function ensures that the 
 * final object reflects these conditions.
 *
 * @param {Sheet} accountDetailsSheet - The sheet object representing the "Account Details" sheet.
 * @param {Sheet} createLinksSheet - The sheet object representing the "Create Links" sheet.
 * @param {string} platform - The platform type, either "android" or "ios", used to determine 
 *                            which cells and headers to use.
 * @param {Array} rowData - The array of data for the current row being processed.
 * @returns {Object} redirection - The constructed redirection object with relevant parameters
 *                                 for the specified platform.
 */
function constructRedirection(accountDetailsSheet, createLinksSheet, platform, rowData) {
  let redirection = {}; // Initialize an empty object to store the redirection details.

  // Determine the cell references based on the platform (Android or iOS).
  const appSiteIdCell = platform === "android" ? "D11" : "D12";
  const destinationUrlCell = platform === "android" ? "E11" : "E12";
  
  // Set the appropriate header names based on the platform.
  const deepLinkHeader = platform === "android" ? "Deep Link (Android)" : "Deep Link (iOS)";
  const deferredDeepLinkHeader = platform === "android" ? "Deferred Deep Link (Android)" : "Deferred Deep Link (iOS)";
  
  // Headers that apply to both platforms.
  const deepLinkAltHeader = "Deep Link (Android & iOS)";
  const deferredDeepLinkAltHeader = "Deferred Deep Link (Android & iOS)";
  
  // Cells that may contain alternate deep link URLs.
  const deeplinkCell = platform === "android" ? "B15" : "B16";

  // Retrieve and set the app_site_id from the account details sheet
  redirection.app_site_id = getCellValue(accountDetailsSheet, appSiteIdCell);
  
  // Retrieve and set the destination_url from the account details sheet
  redirection.destination_url = getCellValue(accountDetailsSheet, destinationUrlCell);
  
  // Set the destination_deeplink_url with the priority:
  // 1. Platform-specific deep link from createLinksSheet
  // 2. Alternate deep link for both Android & iOS from createLinksSheet
  // 3. Fallback to deep link scheme from accountDetailsSheet if others are unavailable
  redirection.destination_deeplink_url = 
    getHeaderData(createLinksSheet, deepLinkHeader, rowData) || 
    getHeaderData(createLinksSheet, deepLinkAltHeader, rowData) || 
    getCellValue(accountDetailsSheet, deeplinkCell);

  // Set the destination_deferred_deeplink_url with the priority:
  // 1. Platform-specific deferred deep link from createLinksSheet
  // 2. Alternate deferred deep link for both Android & iOS from createLinksSheet
  // 3. Platform-specific deep link from createLinksSheet
  // 4. Alternate deep link for both Android & iOS from createLinksSheet  
  // 5. Fallback to deep link scheme from accountDetailsSheet if others are unavailable
  redirection.destination_deferred_deeplink_url = 
    getHeaderData(createLinksSheet, deferredDeepLinkHeader, rowData) || 
    getHeaderData(createLinksSheet, deferredDeepLinkAltHeader, rowData) || 
    getHeaderData(createLinksSheet, deepLinkHeader, rowData) || 
    getHeaderData(createLinksSheet, deepLinkAltHeader, rowData) || 
    getCellValue(accountDetailsSheet, deeplinkCell);

  // Return the constructed redirection object
  return redirection;
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Function to get the header index of a specified header name in a given sheet.
 * 
 * @param {Sheet} sheet - The Google Sheets object.
 * @param {string} headerName - The name of the header to find the index for.
 * @return {number} The index of the header or -1 if not found.
 */
// function getHeaderIndex(sheet, headerName) {
//   // Get the header row values (assumed to be in row 7)
//   const headerRow = sheet.getRange(7, 1, 1, sheet.getLastColumn()).getValues()[0];
  
//   // Iterate through the header row to find the index of the specified header name
//   for (let i = 0; i < headerRow.length; i++) {
//     if (headerRow[i] === headerName) {
//       return i; // Return the index if the header is found
//     }
//   }
//   return -1; // Return -1 if the header is not found
// }

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Retrieves the data from a specific column in a row based on the header name.
 *
 * @param {Sheet} sheet - The Google Sheets sheet object where the headers are located.
 * @param {string} headerName - The name of the header to search for.
 * @param {Array} rowData - An array representing the data in a single row.
 * @returns {any} - The value from the row corresponding to the specified header name, or null if the header is not found.
 */
function getHeaderData(sheet, headerName, rowData) {
  // Get the index (column number) of the specified header in the sheet using the getHeaderIndex function
  const headerIndex = getHeaderIndex(sheet, headerName);
  
  // Check if the header exists (headerIndex is not -1)
  if (headerIndex !== -1) {
    // If the header exists, return the data from the corresponding column in the rowData array
    return rowData[headerIndex];
  }
  
  // If the header does not exist, return null to indicate that the data is not available
  return null;
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Function to get the row data from a specified range in the sheet.
 * 
 * @param {Sheet} sheet - The Google Sheets object.
 * @param {string} range - The range reference (e.g., "7:7" for the header row).
 * @return {Array} The row data as an array.
 */
function getRowData(sheet, range) {
  return sheet.getRange(range).getValues()[0];
}


// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

/**
 * Displays the constructed request body in a dialog box within Google Sheets UI.
 *
 * @param {Object} requestBody - The JSON object representing the request body to be sent in an API request.
 */
function showRequestBody(requestBody) {
  // Get the Google Sheets UI object to interact with the user interface
  var ui = SpreadsheetApp.getUi();

  // Convert the requestBody object to a formatted JSON string with indentation for readability
  var formattedRequestBody = JSON.stringify(requestBody, null, 2); // Pretty-print JSON
  
  // Display an alert dialog box showing the formatted JSON request body
  ui.alert('Request Body', formattedRequestBody, ui.ButtonSet.OK);
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

function generateQRCodeURLs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create Links");
  var lastRow = sheet.getLastRow();
  var headerRow = 7;

  // Get the size from a specific cell, e.g., cell A1
  var sizeInput = sheet.getRange("H5").getValue();
  
  // Validate the size input
  var size = (sizeInput >= 100 && sizeInput <= 1000) ? sizeInput : 250;

  // Find the column index of "Long Link" in row 7
  var headerRowValues = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var longLinkColumnIndex = headerRowValues.indexOf("Long Link") + 1; // Add 1 because indexOf is 0-based

  if (longLinkColumnIndex === 0) {
    Logger.log('Column "Long Link" not found.');
    return;
  }

  // Determine the output column index (third column after "Long Link")
  var outputColumnIndex = longLinkColumnIndex + 2;

  // Loop through each row, starting from the row below the header
  for (var i = headerRow + 1; i <= lastRow; i++) {
    var link = sheet.getRange(i, longLinkColumnIndex).getValue();
    
    if (link) {
      // Generate QR code URL using goqr.me API with the specified size
      var qrCodeUrl = "https://api.qrserver.com/v1/create-qr-code/?size=" + size + "x" + size + "&data=" + encodeURIComponent(link);
      
      // Place the QR code URL into the output column
      sheet.getRange(i, outputColumnIndex).setValue(qrCodeUrl);
    }
  }
}

// ---------------------------------------------------------------------------------------//
// ---------------------------------------------------------------------------------------//

function downloadQRCodesToDrive() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create Links");
  var lastRow = sheet.getLastRow();
  var headerRow = 7;

  // Get the size from a specific cell, e.g., cell A1
  var sizeInput = sheet.getRange("H5").getValue();
  
  // Validate the size input
  var size = (sizeInput >= 100 && sizeInput <= 1000) ? sizeInput : 250;

  // Find the column index of "Long Link" in row 7
  var headerRowValues = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var longLinkColumnIndex = headerRowValues.indexOf("Long Link") + 1; // Add 1 because indexOf is 0-based

  if (longLinkColumnIndex === 0) {
    Logger.log('Column "Long Link" not found.');
    return;
  }

  // Determine the output column index (third column after "Long Link")
  var outputColumnIndex = longLinkColumnIndex + 2;

  // Get or create the "Tracking Links QR Images" folder
  var folderName = "Tracking Links QR Images";
  var folders = DriveApp.getFoldersByName(folderName);
  var folder;
  
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  // Loop through each row, starting from the row below the header
  for (var i = headerRow + 1; i <= lastRow; i++) {
    var link = sheet.getRange(i, longLinkColumnIndex).getValue();
    var imageName = sheet.getRange(i, 2).getValue(); // Get image name from column B
    
    if (link && imageName) {
      // Generate QR code URL using goqr.me API with the specified size
      var qrCodeUrl = "https://api.qrserver.com/v1/create-qr-code/?size=" + size + "x" + size + "&data=" + encodeURIComponent(link);

      // Fetch the image as a blob
      var imageBlob = UrlFetchApp.fetch(qrCodeUrl).getBlob();
      
      // Get the current date and time
      var now = new Date();
      var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
      
      // Set the file name with the image name from column B, date, and time
      var fileName = imageName + "_" + formattedDate + ".png";
      imageBlob.setName(fileName);
      
      // Save the image to the "Tracking Links QR Images" folder
      folder.createFile(imageBlob);
      
      // Optionally, place a confirmation message in the output column
      sheet.getRange(i, outputColumnIndex).setValue("Image Saved: " + fileName);
    }
  }
}












