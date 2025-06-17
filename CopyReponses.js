// --- CONFIGURATION (BEST PRACTICE: Use Project Properties for IDs) ---
// Go to Script Editor > Project settings (the gear icon on the left) > Script properties.
// Add the following properties:
// Key: DIRECTORY_SPREADSHEET_ID
// Value: 1fDf21MXhPuCrw5hHP_yrrYZNf108cISAETKTolz9LiU (Replace with your actual Directory Spreadsheet ID)
//
// Key: TARGET_ATTENDANCE_SPREADSHEET_ID
// Value: YOUR_ATTENDANCE_TRACKER_SPREADSHEET_ID_HERE (Replace with the ID of your 'attendance tracker' Google Sheet)

// This script needs to be authorized to access external spreadsheets.
// When you run it for the first first time or set up a trigger, you'll be prompted to grant permissions.

// --- MAIN TRIGGER: For a SINGLE Event Attendance Form ---
/**
 * Master handler for a specific Google Form submission.
 * This function should be set as the ONLY 'On form submit' trigger for *your primary Event Attendance form*.
 * It first processes the submission to the central 'Event Attendance' sheet (where the script lives),
 * assigns an ID there, and then copies the complete row to the 'Event Attendance' sheet in the
 * TARGET_ATTENDANCE_SPREADSHEET.
 *
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e The event object from the form submission.
 */
function masterEventFormSubmit(e) {
  Logger.log("MASTER EVENT FORM SUBMIT: masterEventFormSubmit triggered.");

  if (!e || !e.values || !e.range || !e.source) {
    Logger.log("Error: Event object is incomplete for masterEventFormSubmit. Skipping processing.");
    SpreadsheetApp.getUi().alert("Error", "Form submission event data is incomplete. Processing skipped.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // --- Get references to both central and target sheets ---
  const currentSs = SpreadsheetApp.getActiveSpreadsheet();
  const centralEventAttendanceSheet = getCaseInsensitiveSheetByName(currentSs, "event attendance");
  if (!centralEventAttendanceSheet) {
    Logger.log("Error: 'event attendance' sheet not found in the CURRENT spreadsheet for masterEventFormSubmit.");
    SpreadsheetApp.getUi().alert("Error", "'event attendance' sheet not found in current spreadsheet. Cannot process form submission.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const targetAttendanceSpreadsheetId = PropertiesService.getScriptProperties().getProperty('TARGET_ATTENDANCE_SPREADSHEET_ID');
  if (!targetAttendanceSpreadsheetId) {
    Logger.log("Error: TARGET_ATTENDANCE_SPREADSHEET_ID not set in Project Properties.");
    SpreadsheetApp.getUi().alert("Configuration Error", "TARGET_ATTENDANCE_SPREADSHEET_ID is not set. Cannot process form submission.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  let targetAttendanceSs;
  try {
    targetAttendanceSs = SpreadsheetApp.openById(targetAttendanceSpreadsheetId);
    Logger.log(`Successfully opened target attendance spreadsheet with ID: ${targetAttendanceSpreadsheetId}`);
  } catch (error) {
    Logger.log(`Error opening target attendance spreadsheet by ID ${targetAttendanceSpreadsheetId}: ${error.message}. Please check ID and permissions.`);
    SpreadsheetApp.getUi().alert("Target Sheet Access Error", `Failed to open target attendance spreadsheet. Check ID and permissions: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const targetEventAttendanceSheet = getCaseInsensitiveSheetByName(targetAttendanceSs, "event attendance");
  if (!targetEventAttendanceSheet) {
    Logger.log("Error: 'event attendance' sheet not found (case-insensitively) in the target attendance spreadsheet for masterEventFormSubmit.");
    SpreadsheetApp.getUi().alert("Error", "'event attendance' sheet not found in target spreadsheet. Cannot process form submission.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  let centralRowIndex;
  try {
    Logger.log("Step 1: Calling processSingleFormSubmission(e) to add initial data to CENTRAL 'event attendance' sheet.");
    // This function appends the raw form data to the central sheet
    processSingleFormSubmission(e, centralEventAttendanceSheet);
    centralRowIndex = centralEventAttendanceSheet.getLastRow();
    Logger.log(`Step 1 Complete: Initial data added to central 'event attendance' sheet at row ${centralRowIndex}.`);
  } catch (error) {
    Logger.log(`Error in processSingleFormSubmission (central append): ${error.message} Stack: ${error.stack}`);
    SpreadsheetApp.getUi().alert("Error", `Error in initial form data processing to central sheet: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return; // Stop if the first step fails
  }

  try {
    Logger.log("Step 2: Calling assignEventAttendanceIds for the new row on the CENTRAL sheet.");
    // Assign ID to the row that was just added to the central sheet
    assignEventAttendanceIds(centralEventAttendanceSheet, centralRowIndex, centralRowIndex);
    Logger.log(`Step 2 Complete: IDs assigned for new entry in central 'event attendance' sheet at row ${centralRowIndex}.`);
  } catch (error) {
    Logger.log(`Error in assignEventAttendanceIds (master trigger for central sheet): ${error.message} Stack: ${error.stack}`);
    SpreadsheetApp.getUi().alert("Error", `Error assigning Event Attendance ID on central sheet: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    // Continue if ID assignment isn't critical to stop the whole process, or return if it is.
  }

  try {
    Logger.log(`Step 3: Copying processed row from central sheet (row ${centralRowIndex}) to target 'event attendance' sheet.`);
    // Read the now fully processed row from the central sheet
    const processedRow = centralEventAttendanceSheet.getRange(centralRowIndex, 1, 1, centralEventAttendanceSheet.getLastColumn()).getValues()[0];
    
    // Append this processed row to the target sheet
    targetEventAttendanceSheet.appendRow(processedRow);
    Logger.log(`Step 3 Complete: Processed row copied to target 'event attendance' sheet at row ${targetEventAttendanceSheet.getLastRow()}.`);
  } catch (error) {
    Logger.log(`Error copying processed row to target sheet: ${error.message} Stack: ${error.stack}`);
    SpreadsheetApp.getUi().alert("Error", `Error copying data to target attendance tracker: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }

  Logger.log("MASTER EVENT FORM SUBMIT: All event attendance form processing complete.");
}

// --- NEW FUNCTION: Syncs data from multiple designated Form Responses sheets ---
/**
 * Monitors and syncs new form responses from multiple specified "Form Responses" sheets
 * into the central "event attendance" sheet. After processing (including ID assignment)
 * in the central sheet, the data is copied to the "event attendance" sheet in the
 * TARGET_ATTENDANCE_SPREADSHEET.
 *
 * This function should be set as a time-driven trigger (e.g., every hour, daily).
 * It uses a "last synced row" property for each source sheet to avoid re-importing old data.
 */
function syncAllFormResponsesToEventAttendance() {
  Logger.log("SYNC ALL FORM RESPONSES: Sync process started.");

  // --- Get references to both central and target sheets ---
  const currentSs = SpreadsheetApp.getActiveSpreadsheet();
  const centralEventAttendanceSheet = getCaseInsensitiveSheetByName(currentSs, "event attendance");
  if (!centralEventAttendanceSheet) {
    Logger.log("Error: 'event attendance' sheet not found in the CURRENT spreadsheet for syncAllFormResponsesToEventAttendance.");
    SpreadsheetApp.getUi().alert("Error", "'event attendance' sheet not found in current spreadsheet. Cannot perform sync.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const targetAttendanceSpreadsheetId = PropertiesService.getScriptProperties().getProperty('TARGET_ATTENDANCE_SPREADSHEET_ID');
  if (!targetAttendanceSpreadsheetId) {
    Logger.log("Error: TARGET_ATTENDANCE_SPREADSHEET_ID not set in Project Properties. Exiting sync.");
    SpreadsheetApp.getUi().alert("Configuration Error", "TARGET_ATTENDANCE_SPREADSHEET_ID is not set. Cannot perform sync.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  let targetAttendanceSs;
  try {
    targetAttendanceSs = SpreadsheetApp.openById(targetAttendanceSpreadsheetId);
    Logger.log(`Successfully opened target attendance spreadsheet with ID: ${targetAttendanceSpreadsheetId}`);
  } catch (error) {
    Logger.log(`Error opening target attendance spreadsheet by ID ${targetAttendanceSpreadsheetId}: ${error.message}. Please check ID and permissions.`);
    SpreadsheetApp.getUi().alert("Target Sheet Access Error", `Failed to open target attendance spreadsheet. Check ID and permissions: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const targetEventAttendanceSheet = getCaseInsensitiveSheetByName(targetAttendanceSs, "event attendance");
  if (!targetEventAttendanceSheet) {
    Logger.log("Error: 'event attendance' sheet not found (case-insensitively) in the target attendance spreadsheet. Exiting sync.");
    SpreadsheetApp.getUi().alert("Error", "'event attendance' sheet not found in target spreadsheet. Cannot perform sync.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Define the source form response sheets to monitor.
  // IMPORTANT:
  // - Each entry should point to a *specific* Form Response sheet.
  // - The 'formDataMap' specifies how columns in *that specific source sheet* map to
  //   the Event Attendance sheet's expected columns (Event Name, Event ID, Last Name, etc.).
  // - Adjust 'spreadsheetId' and 'sheetName' for each form you want to sync.
  // - Adjust the 'colIndex' values based on the *0-indexed* columns of the *source form's responses*.
  const sourceFormsConfig = [
    {
      spreadsheetId: currentSs.getId(), // Use current spreadsheet ID if form is in the same file
      sheetName: 'Form Responses 1', // E.g., 'Form Responses 1'
      formDataMap: {
        timestampCol: 0, // Column A
        eventNameCol: 1, // Column B
        eventIdCol: 2,   // Column C
        lastNameCol: 3,  // Column D
        firstNameCol: 4, // Column E
        roleCol: 5,      // Column F
        emailCol: 6,     // Column G
        phoneCol: 7      // Column H
      }
    },
    {
      // Example for another form in a different spreadsheet:
      // You need to know the Spreadsheet ID and the name of its Form Responses sheet.
      spreadsheetId: 'ANOTHER_SPREADSHEET_ID_HERE', // <-- REPLACE THIS with the actual ID
      sheetName: 'Another Form Responses', // <-- REPLACE THIS with the actual sheet name
      formDataMap: {
        timestampCol: 0, // Assuming Timestamp is first column (index 0)
        eventNameCol: 4, // Maybe Event Name is 5th column (index 4) in this form
        eventIdCol: 5,
        lastNameCol: 1,
        firstNameCol: 2,
        roleCol: 6,
        emailCol: 7,
        phoneCol: 8
      }
    }
    // Add more configurations for other form response sheets as needed
  ];

  const scriptProperties = PropertiesService.getScriptProperties();
  let centralRowsAppendedIndexes = []; // To track rows in the central sheet for ID assignment and copying

  for (const config of sourceFormsConfig) {
    let sourceSs;
    try {
      sourceSs = SpreadsheetApp.openById(config.spreadsheetId);
    } catch (e) {
      Logger.log(`Error opening source spreadsheet ID ${config.spreadsheetId}: ${e.message}. Skipping this source.`);
      SpreadsheetApp.getUi().alert("Sync Error", `Failed to open source spreadsheet ${config.spreadsheetId}. Check ID and permissions.`, SpreadsheetApp.getUi().ButtonSet.OK);
      continue;
    }

    // Get the source sheet using case-insensitive lookup
    const sourceSheet = getCaseInsensitiveSheetByName(sourceSs, config.sheetName);
    if (!sourceSheet) {
      Logger.log(`Warning: Source sheet "${config.sheetName}" not found (case-insensitively) in spreadsheet ID ${config.spreadsheetId}. Skipping.`);
      continue;
    }

    const lastRowInSource = sourceSheet.getLastRow();
    const propertyKey = `lastSyncedRow_${config.spreadsheetId}_${sourceSheet.getName().replace(/\s/g, '_')}`; // Use actual sheet name for property key
    let lastSyncedRow = parseInt(scriptProperties.getProperty(propertyKey) || '1'); // Start from 1 (headers) if no property

    // If there's only a header row, or we've synced everything, skip.
    if (lastRowInSource <= lastSyncedRow) {
      Logger.log(`No new entries in ${sourceSheet.getName()} (ID: ${config.spreadsheetId}). Last row: ${lastRowInSource}, Last synced: ${lastSyncedRow}.`);
      continue;
    }

    Logger.log(`Syncing new entries from ${sourceSheet.getName()} (ID: ${config.spreadsheetId}) from row ${lastSyncedRow + 1} to ${lastRowInSource}.`);

    // Get new data from source sheet (from the row after last synced up to the last row)
    const newDataRange = sourceSheet.getRange(lastSyncedRow + 1, 1, lastRowInSource - lastSyncedRow, sourceSheet.getLastColumn());
    const newData = newDataRange.getValues();

    if (newData.length === 0) {
      Logger.log(`No new actual data found in ${sourceSheet.getName()} (ID: ${config.spreadsheetId}).`);
      scriptProperties.setProperty(propertyKey, String(lastRowInSource)); // Update to prevent re-checking empty range
      continue;
    }

    let appendedCount = 0;
    for (let i = 0; i < newData.length; i++) {
      const rowData = newData[i];
      const formSheetName = sourceSheet.getName(); // Use actual sheet name for record
      const dataMap = config.formDataMap;

      // Extract and map data using the specific config for this source form
      const timestampRaw = new Date(); // Current processing timestamp
      const eventDate = Utilities.formatDate(timestampRaw, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd'); // Format for Event Date
      const eventName = rowData[dataMap.eventNameCol] || '';
      const eventId = rowData[dataMap.eventIdCol] || '';
      const lastName = rowData[dataMap.lastNameCol] || '';
      const firstName = rowData[dataMap.firstNameCol] || '';
      const role = rowData[dataMap.roleCol] || 'Unknown';
      const email = rowData[dataMap.emailCol] || '';
      const phone = rowData[dataMap.phoneCol] || '';

      const fullName = `${String(firstName).trim()} ${String(lastName).trim()}`;

      // Prepare the data row for the "Event Attendance" sheet
      // Column Order: Person ID, Full Name, Event Name, Event ID, First Name, Last Name, Email, Phone Number, Form Sheet, Role, Event Date, First Time?, Needs Follow-up?, Timestamp
      const dataRowForEventAttendance = [
        '',           // Column A: Person ID (leave blank for now, assignEventAttendanceIds will fill)
        fullName,     // Column B: Full Name
        eventName,    // Column C: Event Name
        eventId,      // Column D: Event ID
        firstName,    // Column E: First Name
        lastName,     // Column F: Last Name
        email,        // Column G: Email
        phone,        // Column H: Phone Number
        formSheetName,// Column I: Form Sheet
        role,         // Column J: Role
        eventDate,    // Column K: Event Date (Date part of processing timestamp)
        '',           // Column L: First Time? (Skipped as per request)
        '',           // Column M: Needs Follow-up? (Skipped as per request)
        timestampRaw  // Column N: Timestamp (Full processing timestamp)
      ];

      try {
        centralEventAttendanceSheet.appendRow(dataRowForEventAttendance);
        const appendedRowIndex = centralEventAttendanceSheet.getLastRow();
        centralRowsAppendedIndexes.push(appendedRowIndex); // Add this row to the list for ID assignment
        appendedCount++;
      } catch (error) {
        Logger.log(`Error appending row from ${formSheetName} to CENTRAL Event Attendance: ${error.message}`);
      }
    }

    scriptProperties.setProperty(propertyKey, String(lastRowInSource)); // Update last synced row
    Logger.log(`Successfully synced ${appendedCount} new entries from ${sourceSheet.getName()} (ID: ${config.spreadsheetId}) to central sheet.`);
  }

  Logger.log(`Total new entries synced to CENTRAL Event Attendance: ${centralRowsAppendedIndexes.length}`);

  // After all data is appended to the CENTRAL sheet, assign IDs in a batch
  if (centralRowsAppendedIndexes.length > 0) {
    Logger.log(`Assigning IDs for ${centralRowsAppendedIndexes.length} new entries in CENTRAL sheet.`);
    // Get the min and max row numbers from the collected list in the central sheet
    const minRowCentral = Math.min(...centralRowsAppendedIndexes);
    const maxRowCentral = Math.max(...centralRowsAppendedIndexes);
    try {
      assignEventAttendanceIds(centralEventAttendanceSheet, minRowCentral, maxRowCentral);
      Logger.log("Batch ID assignment complete for new entries in CENTRAL sheet.");
    } catch (error) {
      Logger.log(`Error during batch ID assignment in CENTRAL sheet: ${error.message} Stack: ${error.stack}`);
      SpreadsheetApp.getUi().alert("Error", `Error during batch ID assignment in central sheet: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }

    // Now, copy the fully processed rows from the CENTRAL sheet to the TARGET sheet
    Logger.log(`Copying processed data from CENTRAL sheet (rows ${minRowCentral}-${maxRowCentral}) to TARGET sheet.`);
    try {
      const processedDataToCopy = centralEventAttendanceSheet.getRange(minRowCentral, 1, centralRowsAppendedIndexes.length, centralEventAttendanceSheet.getLastColumn()).getValues();
      targetEventAttendanceSheet.appendRows(processedDataToCopy); // Use appendRows for batch append
      Logger.log(`Successfully copied ${processedDataToCopy.length} rows to TARGET sheet.`);
    } catch (error) {
      Logger.log(`Error copying processed data from central to target sheet: ${error.message} Stack: ${error.stack}`);
      SpreadsheetApp.getUi().alert("Error", `Error copying processed data to target attendance tracker: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } else {
    Logger.log("No new entries to assign IDs to or copy.");
  }

  Logger.log("SYNC ALL FORM RESPONSES: Sync process finished.");
}


// --- UPDATED processSingleFormSubmission (now receives central sheet) ---
/**
 * Processes a single form submission event and logs it to the "event attendance" sheet
 * in the CURRENT (central) SPREADSHEET.
 * New entries are added to the next available blank row.
 *
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e The event object passed by the form submission trigger.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} centralEventAttendanceSheet The Event Attendance sheet object in the central spreadsheet.
 */
function processSingleFormSubmission(e, centralEventAttendanceSheet) {
  Logger.log("processSingleFormSubmission function started.");

  if (!e || !e.values || !e.range || !e.source) {
    Logger.log("Error: Event object is undefined or incomplete. Exiting function.");
    return;
  }

  const sheet = e.source.getActiveSheet(); // This is the form response sheet from the original form's spreadsheet
  const sheetName = sheet.getName(); // The name of the form response sheet (e.g., "Form Responses 1")
  const formData = e.values;     // Get the form response data (array of values from the submitted row)

  // As per user's instruction: Do NOT write headers. Assume they are pre-existing in both sheets.
  // The script will now always append data to the next available row.

  // --- IMPORTANT: Ensure these indices match your PRIMARY Event Attendance Form structure ---
  // formData indices are 0-based, corresponding to the order of questions in your form.
  // Assuming form: Timestamp(0), Event Name(1), Event ID(2), Last Name(3), First Name(4), Role(5), Email(6), Phone(7)
  const eventName = formData[1] || '';   // Assuming Event Name is the 2nd form question
  const eventId = formData[2] || '';     // Assuming Event ID is the 3rd form question
  const lastName = formData[3] || '';    // Assuming Last Name is the 4th form question
  const firstName = formData[4] || '';   // Assuming First Name is the 5th form question
  const role = formData[5] || 'Unknown'; // Assuming Role is the 6th form question
  const email = formData[6] || '';       // Assuming Email is the 7th form question
  const phone = formData[7] || '';       // Assuming Phone Number is the 8th form question

  const fullName = `${String(firstName).trim()} ${String(lastName).trim()}`;
  const formSheetNameValue = sheetName || ''; // Form Sheet name

  const timestampRaw = new Date(); // Current processing timestamp
  const eventDate = Utilities.formatDate(timestampRaw, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd'); // Format for Event Date

  // Prepare the data row for the "Event Attendance" sheet (14 columns)
  // Column Order: Person ID, Full Name, Event Name, Event ID, First Name, Last Name, Email, Phone Number, Form Sheet, Role, Event Date, First Time?, Needs Follow-up?, Timestamp
  const dataRowForEventAttendance = [
    '',           // Column A: Person ID (leave blank for now, assignEventAttendanceIds will fill)
    fullName,     // Column B: Full Name
    eventName,    // Column C: Event Name
    eventId,      // Column D: Event ID
    firstName,    // Column E: First Name
    lastName,     // Column F: Last Name
    email,        // Column G: Email
    phone,        // Column H: Phone Number
    formSheetNameValue, // Column I: Form Sheet
    role,         // Column J: Role
    eventDate,    // Column K: Event Date (Date part of processing timestamp)
    '',           // Column L: First Time? (Skipped as per request)
    '',           // Column M: Needs Follow-up? (Skipped as per request)
    timestampRaw  // Column N: Timestamp (Full processing timestamp)
  ];

  // Append the data row to the CENTRAL Event Attendance sheet
  centralEventAttendanceSheet.appendRow(dataRowForEventAttendance);
  Logger.log("Data successfully appended to central Event Attendance sheet via single form submission.");
}


// --- assignEventAttendanceIds (receives the sheet it needs to operate on) ---
/**
 * Processes Event Attendance entries to assign unique, purely numeric Personal ID codes.
 * This version is modified to process a specific range of rows, allowing for single-row
 * or batch processing. It operates on the sheet passed to it (which will be the central sheet).
 *
 * It fetches comprehensive ID maps and the highest existing ID from all relevant sheets
 * to ensure consistency and correct ID generation.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} eventAttendanceSheet The Event Attendance sheet object (the central sheet for ID assignment).
 * @param {number} startRow The starting row (1-indexed) to process in Event Attendance.
 * @param {number} endRow The ending row (1-indexed) to process in Event Attendance.
 */
function assignEventAttendanceIds(eventAttendanceSheet, startRow, endRow) {
  Logger.log(`Event Attendance ID script started for rows ${startRow} to ${endRow} on sheet "${eventAttendanceSheet.getName()}".`);

  // --- Configuration ---
  const directorySpreadsheetId = PropertiesService.getScriptProperties().getProperty('DIRECTORY_SPREADSHEET_ID');
  if (!directorySpreadsheetId) {
    Logger.log('Error: DIRECTORY_SPREADSHEET_ID not set in Project Properties. Cannot assign IDs.');
    SpreadsheetApp.getUi().alert("Configuration Error", "DIRECTORY_SPREADSHEET_ID is not set in Script Properties. Cannot assign IDs.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  Logger.log(`Attempting to access directory spreadsheet with ID: ${directorySpreadsheetId}`);

  // Name of the main Event Attendance sheet that THIS FUNCTION is operating on (e.g., the central sheet).
  // Its column structure is used for lookup and update within this function.
  const eventAttendanceTabName = eventAttendanceSheet.getName(); // Use actual name of the sheet passed in
  const eventAttendanceNameColumn = 2; // Column B (Full Name) in the sheet this function is operating on
  const eventAttendanceIdColumn = 1;   // Column A (Person ID) in the sheet this function is operating on

  // Configuration for external sheets (these are in the directory spreadsheet)
  const externalSheetIdColumn = 1;     // Column A in external sheets
  const externalSheetNameColumn = 2;   // Column B in external sheets

  const directoryTabName = 'directory';
  const newMemberFormTabName = 'new member form';
  const sundayServiceAttendTabName = 'Sunday Service Attend';
  const directoryEventAttendanceLogTabName = 'event attendance'; // "event attendance" tab located within your DIRECTORY SPREADSHEET


  // --- Get spreadsheet and sheet references ---
  let directorySs;
  try {
    directorySs = SpreadsheetApp.openById(directorySpreadsheetId);
    Logger.log(`Successfully opened directory spreadsheet with ID: ${directorySpreadsheetId}`);
  } catch (e) {
    Logger.log(`Error opening directory spreadsheet by ID ${directorySpreadsheetId}: ${e.toString()}. Please check ID and permissions.`);
    SpreadsheetApp.getUi().alert("Directory Access Error", `Failed to open directory spreadsheet. Check ID and permissions: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Helper to get sheets by name (case-insensitive)
  const getCaseInsensitiveSheetByName = (spreadsheet, targetName) => {
    if (!spreadsheet) return null;
    const sheets = spreadsheet.getSheets();
    const lowerCaseTargetName = targetName.toLowerCase();
    for (const sheet of sheets) {
      if (sheet.getName().toLowerCase() === lowerCaseTargetName) {
        return sheet;
      }
    }
    return null; // Sheet not found
  };

  // Helper to get sheets safely, using case-insensitive lookup
  const getSheetSafely = (spreadsheet, sheetName, type = "directory") => {
    try {
      if (!spreadsheet) {
        Logger.log(`Error: Spreadsheet object is null for getSheetSafely (sheet: "${sheetName}", type: "${type}").`);
        return null;
      }
      const sheet = getCaseInsensitiveSheetByName(spreadsheet, sheetName);

      if (!sheet) {
        Logger.log(`Warning: Sheet named "${sheetName}" not found (case-insensitively) in the ${type} spreadsheet. Proceeding without it.`);
        return null;
      }
      Logger.log(`Successfully accessed "${sheet.getName()}" (originally sought as "${sheetName}") sheet in the ${type} spreadsheet.`);
      return sheet;
    } catch (e) {
      Logger.log(`Error accessing "${sheetName}" sheet in the ${type} spreadsheet: ${e.toString()}`);
      return null;
    }
  };

  const directorySheet = getSheetSafely(directorySs, directoryTabName, "directory");
  const newMemberFormExternalSheet = getSheetSafely(directorySs, newMemberFormTabName, "directory");
  const sundayServiceAttendExternalSheet = getSheetSafely(directorySs, sundayServiceAttendTabName, "directory");
  const directoryEventAttendanceLogSheet = getSheetSafely(directorySs, directoryEventAttendanceLogTabName, "directory");

  if (!directorySheet) { // Directory is critical for ID lookup
    Logger.log(`Critical Error: Sheet named "${directoryTabName}" not found (case-insensitively) in the directory spreadsheet. Script cannot continue.`);
    SpreadsheetApp.getUi().alert("Critical Error", `"${directoryTabName}" sheet not found in the directory spreadsheet. Cannot assign IDs.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // --- Read ALL data from relevant sheets for building comprehensive lookup maps and highest ID ---
  const getDataFromSheet = (sheet, sheetNameForLogging) => { // Added sheetNameForLogging to avoid confusion
    if (!sheet) return [];
    try {
      if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
        Logger.log(`Sheet "${sheetNameForLogging}" is empty or has no columns. Returning empty data.`);
        return [];
      }
      const data = sheet.getDataRange().getValues();
      Logger.log(`Read full ${data.length} rows from "${sheetNameForLogging}".`);
      return data;
    } catch (e) {
      Logger.log(`Error reading data from "${sheetNameForLogging}": ${e.toString()}`);
      return [];
    }
  };

  const directoryDataFull = getDataFromSheet(directorySheet, directoryTabName);
  const eventAttendanceDataFull = getDataFromSheet(eventAttendanceSheet, eventAttendanceTabName); // Data from the sheet this function is operating on (e.g., central sheet)
  const newMemberFormDataFull = getDataFromSheet(newMemberFormExternalSheet, newMemberFormTabName);
  const sundayServiceAttendDataFull = getDataFromSheet(sundayServiceAttendExternalSheet, sundayServiceAttendTabName);
  const directoryEventAttendanceLogDataFull = getDataFromSheet(directoryEventAttendanceLogSheet, directoryEventAttendanceLogTabName);


  // --- Calculate the highest existing numeric ID across ALL relevant sheets ---
  let highestExistingNumber = 0;
  const sheetsForHighestNum = [
    { data: directoryDataFull, name: directoryTabName, idCol: externalSheetIdColumn - 1 },
    { data: newMemberFormDataFull, name: newMemberFormTabName, idCol: externalSheetIdColumn - 1 },
    { data: sundayServiceAttendDataFull, name: sundayServiceAttendTabName, idCol: externalSheetIdColumn - 1 },
    { data: eventAttendanceDataFull, name: eventAttendanceTabName, idCol: eventAttendanceIdColumn - 1 }, // From the sheet this function is operating on
    { data: directoryEventAttendanceLogDataFull, name: directoryEventAttendanceLogTabName, idCol: externalSheetIdColumn - 1 } // From DIRECTORY spreadsheet
  ];

  for (const sheetInfo of sheetsForHighestNum) {
    if (!sheetInfo.data || sheetInfo.data.length < 2) continue; // Skip if no header or no data rows
    let currentHighestInSheet = 0;
    // Iterate from row 1 (index 1) to skip header row
    for (let i = 1; i < sheetInfo.data.length; i++) {
      const id = sheetInfo.data[i][sheetInfo.idCol];
      if (id) {
        const number = extractNumberFromId(id);
        if (!isNaN(number)) {
          currentHighestInSheet = Math.max(currentHighestInSheet, number);
        }
      }
    }
    highestExistingNumber = Math.max(highestExistingNumber, currentHighestInSheet);
    Logger.log(`Highest num in "${sheetInfo.name}": ${currentHighestInSheet}. Overall highest: ${highestExistingNumber}`);
  }
  Logger.log(`Final highest number across all sheets before processing: ${highestExistingNumber}`);

  // --- Build Lookup Maps (Name -> ID from external sheets) ---
  const buildLookupMap = (data, idCol, nameCol, mapName) => {
    const map = new Map();
    if (!data || data.length < 2) return map; // No data rows (only header or empty)
    for (let i = 1; i < data.length; i++) { // Skip header row
      const row = data[i];
      const id = row[idCol];
      const name = row[nameCol];
      if (name && id) { // Ensure both name and ID exist
        map.set(String(name).trim().toUpperCase(), String(id).trim());
      }
    }
    Logger.log(`Built ${mapName} map with ${map.size} entries.`);
    return map;
  };

  const directoryMap = buildLookupMap(directoryDataFull, externalSheetIdColumn - 1, externalSheetNameColumn - 1, directoryTabName);
  const newMemberFormExternalMap = buildLookupMap(newMemberFormDataFull, externalSheetIdColumn - 1, externalSheetNameColumn - 1, newMemberFormTabName);
  const sundayServiceAttendExternalMap = buildLookupMap(sundayServiceAttendDataFull, externalSheetIdColumn - 1, externalSheetNameColumn - 1, sundayServiceAttendTabName);
  const directoryEventAttendanceLogMap = buildLookupMap(directoryEventAttendanceLogDataFull, externalSheetIdColumn - 1, externalSheetNameColumn - 1, directoryEventAttendanceLogTabName);


  // --- Process the specified range of rows in 'Event Attendance' (the sheet passed to this function) ---
  if (endRow < startRow || startRow < 1 || endRow < 1) {
    Logger.log(`Invalid row range specified: startRow=${startRow}, endRow=${endRow}. Skipping ID assignment.`);
    return;
  }
  
  const actualLastRowOnTargetSheet = eventAttendanceSheet.getLastRow();
  const rangeEndRow = Math.min(endRow, actualLastRowOnTargetSheet);
  const numRowsToProcess = rangeEndRow - startRow + 1;

  if (numRowsToProcess <= 0) {
    Logger.log(`No rows to process in the specified range (${startRow}-${endRow}) within actual data range (up to ${actualLastRowOnTargetSheet}).`);
    return;
  }

  // Get the full name and existing ID for the batch of rows from the sheet this function is operating on
  const rangeToProcess = eventAttendanceSheet.getRange(startRow, eventAttendanceIdColumn, numRowsToProcess, eventAttendanceNameColumn);
  const valuesToProcess = rangeToProcess.getValues(); // This will get values from Col A and Col B

  const idsToUpdate = []; // Array to store IDs that need to be written back (e.g., [[id1], [id2], ...])
  const eventAttendanceProcessedNamesAndIdsThisRun = new Map(); // Cache for names processed in this run

  for (let i = 0; i < valuesToProcess.length; i++) {
    const currentRowData = valuesToProcess[i];
    const currentRowGlobalIndex = startRow + i; // Global 1-indexed row number in the sheet

    const existingIdInEventCell = currentRowData[eventAttendanceIdColumn - 1] ? String(currentRowData[eventAttendanceIdColumn - 1]).trim() : "";
    const eventAttendanceNameValue = currentRowData[eventAttendanceNameColumn - 1]; // Full Name from Column B

    const nameHasContent = (eventAttendanceNameValue && String(eventAttendanceNameValue).trim() !== '');

    Logger.log(`Processing row ${currentRowGlobalIndex}: Name: "${eventAttendanceNameValue}", Current ID in Cell: "${existingIdInEventCell}". Name has content: ${nameHasContent}`);

    let finalNumericIdToAssign = "";
    let idSourceDescription = "Not Assigned";

    if (nameHasContent) {
      const formattedName = String(eventAttendanceNameValue).trim().toUpperCase();

      // Priority 1: Check cache from this run (for consistent IDs within a batch)
      if (eventAttendanceProcessedNamesAndIdsThisRun.has(formattedName)) {
        finalNumericIdToAssign = eventAttendanceProcessedNamesAndIdsThisRun.get(formattedName);
        idSourceDescription = "this run's cache";
        Logger.log(`Row ${currentRowGlobalIndex}: Name "${formattedName}" found in this run's cache. Using ID: "${finalNumericIdToAssign}".`);
      } else {
        // Priority 2: External Sheets (these are still formatted to 5-digits)
        const getNumericIdFromExternal = (nameKey, map, mapName, rn) => {
          if (map && map.has(nameKey)) {
            const idFromSource = map.get(nameKey);
            const numericPart = extractNumberFromId(idFromSource);
            if (!isNaN(numericPart)) {
              const determinedId = String(numericPart).padStart(5, '0'); // External IDs are padded
              Logger.log(`Row ${rn}: For name "${nameKey}" in map "${mapName}", original ID: "${idFromSource}", processed to numeric: "${determinedId}".`);
              return determinedId;
            } else {
              Logger.log(`Row ${rn}: For name "${nameKey}" in map "${mapName}", ID "${idFromSource}" was unparsable.`);
            }
          }
          return "";
        };

        // Check directory sheets in order of priority
        finalNumericIdToAssign = getNumericIdFromExternal(formattedName, directoryMap, directoryTabName, currentRowGlobalIndex);
        if (finalNumericIdToAssign) {
          idSourceDescription = `external "${directoryTabName}"`;
        } else {
          finalNumericIdToAssign = getNumericIdFromExternal(formattedName, newMemberFormExternalMap, newMemberFormTabName, currentRowGlobalIndex);
          if (finalNumericIdToAssign) {
            idSourceDescription = `external "${newMemberFormTabName}"`;
          } else {
            finalNumericIdToAssign = getNumericIdFromExternal(formattedName, sundayServiceAttendExternalMap, sundayServiceAttendTabName, currentRowGlobalIndex);
            if (finalNumericIdToAssign) {
              idSourceDescription = `external "${sundayServiceAttendTabName}"`;
            } else {
              // Now check the "event attendance" tab within the DIRECTORY spreadsheet
              finalNumericIdToAssign = getNumericIdFromExternal(formattedName, directoryEventAttendanceLogMap, directoryEventAttendanceLogTabName, currentRowGlobalIndex);
              if (finalNumericIdToAssign) {
                idSourceDescription = `external "${directoryEventAttendanceLogTabName}"`;
              }
            }
          }
        }

        // Priority 3: Existing simple numeric ID in current cell (if not found externally or in cache)
        if (!finalNumericIdToAssign && existingIdInEventCell) {
            // Check if the existing cell content is purely a string of digits (e.g., "11", "007", "12345")
            if (/^\d+$/.test(existingIdInEventCell)) {
                finalNumericIdToAssign = existingIdInEventCell; // Use the exact string from the cell
                idSourceDescription = "Event Attendance (existing numeric string cell value)";
                Logger.log(`Row ${currentRowGlobalIndex}: Using existing numeric string cell value "${existingIdInEventCell}" as ID for "${formattedName}".`);
            } else {
                Logger.log(`Row ${currentRowGlobalIndex}: Existing cell ID "${existingIdInEventCell}" for "${formattedName}" is not a simple numeric string. Will proceed to other logic or generation.`);
            }
        }

        // Priority 4: Generate new ID if still not found
        if (!finalNumericIdToAssign) {
          highestExistingNumber++;
          finalNumericIdToAssign = String(highestExistingNumber).padStart(5, '0'); // New IDs are padded
          idSourceDescription = "newly generated";
          Logger.log(`Row ${currentRowGlobalIndex}: Generating new numeric ID for "${formattedName}": "${finalNumericIdToAssign}". Previous overall highest was ${highestExistingNumber - 1}.`);
        }

        // Cache the determined ID for this name FOR THIS RUN, if one was found/generated
        if (finalNumericIdToAssign) {
          eventAttendanceProcessedNamesAndIdsThisRun.set(formattedName, finalNumericIdToAssign);
          Logger.log(`Row ${currentRowGlobalIndex}: Cached ID "${finalNumericIdToAssign}" for name "${formattedName}" for this run. Source: ${idSourceDescription}.`);
        }
      }

      // Prepare for batch update: Only update if the ID is new or different from existing
      if (finalNumericIdToAssign && finalNumericIdToAssign !== existingIdInEventCell) {
        idsToUpdate.push([finalNumericIdToAssign]); // Prepare for setValues which expects [[val1], [val2], ...]
        Logger.log(`Prepared to update Row ${currentRowGlobalIndex} with ID "${finalNumericIdToAssign}". Source: ${idSourceDescription}. Old cell ID: "${existingIdInEventCell}".`);
      } else {
        idsToUpdate.push([existingIdInEventCell]); // Keep existing value to fill the array for setValues
        Logger.log(`Row ${currentRowGlobalIndex}: ID "${finalNumericIdToAssign}" for "${formattedName}" (Source: ${idSourceDescription}) matches cell value "${existingIdInEventCell}" or no update needed.`);
      }
    } else {
      idsToUpdate.push([existingIdInEventCell]); // Keep existing value if name is empty
      Logger.log(`Processing row ${currentRowGlobalIndex}: Skipping ID assignment. Name column is empty or blank. Keeping cell value "${existingIdInEventCell}".`);
    }
  }

  // Perform batch update for IDs
  if (idsToUpdate.length > 0) {
    try {
      eventAttendanceSheet.getRange(startRow, eventAttendanceIdColumn, idsToUpdate.length, 1).setValues(idsToUpdate);
      Logger.log(`Successfully batch updated IDs for ${idsToUpdate.length} rows from row ${startRow}.`);
    } catch (error) {
      Logger.log(`Error: Failed to write IDs to '${eventAttendanceSheet.getName()}' sheet at range ${startRow}:${rangeEndRow}. Error: ${error.message}`);
      SpreadsheetApp.getUi().alert("Write Error", `Failed to write batch IDs: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }

  Logger.log(`Event Attendance ID script finished for specified rows on sheet "${eventAttendanceSheet.getName()}".`);
}

/**
 * Helper function to extract the numeric suffix from an ID string.
 * Tries to handle formats like "BASE - NNNN", "BASE - NNNNN", "BASENNNNN", "BASENNNN", or just "NNNNN".
 * @param {string|number} id - The ID string or number.
 * @returns {number} The extracted numeric suffix as a number, or NaN if not found or invalid.
 */
function extractNumberFromId(id) {
  if (id === null || id === undefined || String(id).trim() === "") return NaN;
  const idStr = String(id);

  // Try matching "ANYTHING - NUMBER" (allows for various prefixes)
  let match = idStr.match(/-\s*(\d+)$/); // Allows space after dash
  if (match && match[1]) {
    return parseInt(match[1], 10);
  }

  // If not matched, try extracting numbers from the end (e.g., BASENNNNN or just NNNNN)
  // This will also match if the ID is purely numeric.
  match = idStr.match(/(\d+)$/);
  if (match && match[1]) {
    return parseInt(match[1], 10);
  }

  // If the string is purely numeric and wasn't caught by (\d+)$
  if (/^\d+$/.test(idStr)) {
    return parseInt(idStr, 10);
  }

  return NaN;
}

// Helper function from other script, included for completeness if needed.
// This function is not used in the primary flow but is good for general sheet utilities.
/**
 * Finds the true last row in a sheet that has data in any of the specified key columns.
 * Accounts for potentially sparse data.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to check.
 * @param {number[]} keyColumnNumbers An array of 1-indexed column numbers to check for data.
 * @returns {number} The 1-indexed number of the last row containing data, or 0 if no data.
 */
function getTrueLastRow(sheet, keyColumnNumbers) {
  let trueLastRow = 0;
  const headerRows = sheet.getFrozenRows() || 1; // Assume at least 1 header row

  try {
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return headerRows; // Sheet is completely empty

    for (let r = data.length - 1; r >= 0; r--) {
      let rowHasDataInKeyColumns = false;
      for (const colIndex of keyColumnNumbers) {
        // Ensure colIndex is within bounds for the current row
        if (r < data.length && (colIndex - 1) < data[r].length) {
          const cellValue = data[r][colIndex - 1];
          if (cellValue !== null && String(cellValue).trim() !== "") {
            rowHasDataInKeyColumns = true;
            break;
          }
        }
      }
      if (rowHasDataInKeyColumns) {
        trueLastRow = r + 1;
        break;
      }
    }
    return trueLastRow > 0 ? trueLastRow : headerRows; // If no data found, return header row count
  } catch (e) {
    Logger.log(`SCRIPT_ERROR: Error in getTrueLastRow for sheet "${sheet.getName()}": ${e.toString()}. Defaulting to a basic getLastRow().`);
    return sheet.getLastRow();
  }
}

// HELPER FUNCTION: To get a sheet by name regardless of its casing
/**
 * Retrieves a Google Sheet by its name in a case-insensitive manner.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The Spreadsheet object to search within.
 * @param {string} targetName The name of the sheet to find (case-insensitive).
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The Sheet object if found, or null otherwise.
 */
function getCaseInsensitiveSheetByName(spreadsheet, targetName) {
  if (!spreadsheet) {
    Logger.log("Error: Spreadsheet object is null for getCaseInsensitiveSheetByName.");
    return null;
  }
  const sheets = spreadsheet.getSheets();
  const lowerCaseTargetName = targetName.toLowerCase();
  for (const sheet of sheets) {
    if (sheet.getName().toLowerCase() === lowerCaseTargetName) {
      return sheet;
    }
  }
  return null; // Sheet not found
}
