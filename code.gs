/**
 * @OnlyCurrentDoc
 */


// --- CONFIGURATION ---
// IMPORTANT: You MUST update this URL if you create a new web app deployment that changes its link!
// Please ensure this is the exact URL of your deployed web app that has "Execute as: Me" and "Who has access: Anyone".
const WEB_APP_URL = "YOUR_WEB_APP_URL_GOES_HERE"; // PASTE YOUR NEW DEPLOYMENT URL HERE
const JD_GENERAL_FOLDER_ID = '1Sv7uvDKlzFhEiM1ljCrRGvC51KgIZJfp';
const JD_INCUMBENT_FOLDER_ID = '1ryXesBBwLs8Y1oEfLYDhDIxPQdeB_Ngx';
const CHANGE_REQUESTS_FOLDER_ID = '1XSW0ktaHt6eRkoZAuHdx8nCo1T2XCSn8';


// Defines the sequential order of approval roles
const APPROVAL_ROLES = ['Prepared By', 'Reviewed By', 'Noted By', 'Approved By'];
const MASTERLIST_EXPORT_FOLDER_ID = '1NcOH0Cx5lPRiRilGKkiO1tRiWn1d3P5q'; // <-- ADD THIS LINE
const TALENT_DATA_SPREADSHEET_ID = '1sBy8d-uuenTRu_jeT7paTtDmnxcHFOjGgn-eEG91knY'; // <-- ADD THIS LINE
const COMPETENCY_SPREADSHEET_ID = '1cj_RuroWG5Tl1OqzalyK7t4dLDck-7Ytj5O1eb-Ks5c'; // <-- ADD THIS LINE
// --- END CONFIGURATION ---


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Org Chart Tools')
    .addItem('Initialize Real-Time Change Log (Run Once)', 'initializeChangeTracking')
    .addItem('Initialize Change Log Sheet', 'initializeChangeLogSheet')
    .addSeparator()
    .addItem('Update Headcount Summary & Create Approval Records', 'takeHeadcountSnapshotWithAlert')
    .addItem('Generate Incumbency History Report', 'generateIncumbencyReport')
    .addItem('Generate Masterlist Export', 'generateMasterlistSheetWithPrompt')
    .addSeparator()
    .addItem('Debug Incumbency History', 'debugIncumbencyForPosition')
    .addItem('Clear Script Cache', 'clearScriptCache')
    .addToUi();
}


/**
 * UTILITY FUNCTION to clear the script's cache.
 */
function clearScriptCache() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirm', 'This will clear all cached data for the web app, which may cause it to load slightly slower one time. Are you sure you want to continue?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    CacheService.getScriptCache().removeAll(['incumbency_history_04-CSD-006']);
    ui.alert('Success! The script cache has been cleared. Please reload the web app for changes to take effect.');
  }
}

function processRequestAction(requestId, action, comments) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Org Chart Requests');
    if (!sheet) {
      throw new Error('"Org Chart Requests" sheet not found.');
    }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const headerMap = new Map(headers.map((h, i) => [h, i]));

    const requestIdCol = headerMap.get('RequestID');

    for (let i = 1; i < data.length; i++) {
      if (data[i][requestIdCol] === requestId) {
        // Update the status and approver details
        sheet.getRange(i + 1, headerMap.get('Status') + 1).setValue(action);
        sheet.getRange(i + 1, headerMap.get('ApproverEmail') + 1).setValue(Session.getActiveUser().getEmail());
        sheet.getRange(i + 1, headerMap.get('ApprovalTimestamp') + 1).setValue(new Date());
        sheet.getRange(i + 1, headerMap.get('ApproverComments') + 1).setValue(comments || '');
        
        return `Request ${requestId} has been successfully ${action}.`;
      }
    }
    throw new Error(`Request ID ${requestId} not found.`);
  } catch (e) {
    Logger.log('Error in processRequestAction: ' + e.message);
    throw new Error('Failed to process request action. ' + e.message);
  }
}

function logToSheet(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Debug Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Debug Log");
    logSheet.appendRow(["Timestamp", "Message"]);
  }
  const timestamp = new Date();
  logSheet.appendRow([timestamp, message]);
}

function implementApprovedChange(requestId) {
  try {
    logToSheet(`Starting implementation for request ID: ${requestId}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Org Chart Requests');
    if (!sheet) {
      throw new Error('"Org Chart Requests" sheet not found.');
    }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const headerMap = new Map(headers.map((h, i) => [h, i]));

    const requestIdCol = headerMap.get('RequestID');
    let requestFound = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][requestIdCol] === requestId) {
        requestFound = true;
        const rowData = data[i];
        logToSheet(`Found request data in row ${i + 1}: ${JSON.stringify(rowData.map(cell => cell instanceof Date ? cell.toISOString() : cell))}`);


        const requestType = rowData[headerMap.get('RequestType')];
        logToSheet(`Request type: ${requestType}`);
        let dataToSave = {};
        let mode = 'edit'; // Default to edit

        // --- IMPLEMENTATION LOGIC ---
        if (requestType.includes('Transfer') || requestType.includes('Promotion')) {
          logToSheet('Processing Transfer/Promotion logic.');
          
          let employeeId = rowData[headerMap.get('EmployeeID')];

          // --- FIX: If EmployeeID is missing, look it up from the main sheet using CurrentPositionID ---
          if (!employeeId) {
            logToSheet('EmployeeID is missing from request. Attempting lookup via CurrentPositionID.');
            if (!headerMap.has('CurrentPositionID')) {
                throw new Error("The 'Org Chart Requests' sheet is missing the required 'CurrentPositionID' column.");
            }
            const currentPositionId = rowData[headerMap.get('CurrentPositionID')]; 
            if (currentPositionId) {
              const mainSheet = ss.getSheets()[0];
              const mainData = mainSheet.getDataRange().getValues();
              const mainHeaders = mainData[0];
              const posIdIndex = mainHeaders.indexOf('Position ID');
              const empIdIndex = mainHeaders.indexOf('Employee ID');

              if (posIdIndex === -1 || empIdIndex === -1) {
                  throw new Error("Could not find 'Position ID' or 'Employee ID' columns in the main data sheet.");
              }

              for (let j = 1; j < mainData.length; j++) {
                if (mainData[j][posIdIndex] === currentPositionId) {
                  employeeId = mainData[j][empIdIndex];
                  logToSheet(`Found EmployeeID "${employeeId}" for CurrentPositionID "${currentPositionId}".`);
                  break;
                }
              }
            }
            if (!employeeId) {
              const errorMessage = `CRITICAL: Could not find a valid EmployeeID to transfer. Looked in CurrentPositionID: "${currentPositionId || 'Not Provided'}".`;
              logToSheet(errorMessage);
              throw new Error(errorMessage);
            }
          }
          // --- END FIX ---

          dataToSave = {
            positionid: rowData[headerMap.get('NewPositionID')],
            employeeid: employeeId, // Use the potentially retrieved employeeId
            employeename: rowData[headerMap.get('EmployeeName')],
            datehired: rowData[headerMap.get('DateHired')],
            dateofbirth: rowData[headerMap.get('DateOfBirth')],
            status: requestType,
            startdateinposition: rowData[headerMap.get('EffectiveDate')]
          };
        } else if (requestType.includes('replacement for vacancy')) {
          logToSheet('Processing Replacement for Vacancy logic.');
          dataToSave = {
            positionid: rowData[headerMap.get('VacantPositionID')],
            employeeid: rowData[headerMap.get('NewEmployeeID')],
            employeename: rowData[headerMap.get('NewEmployeeName')],
            status: 'FILLED VACANCY',
            startdateinposition: rowData[headerMap.get('EffectiveDate')]
          };
        } else if (requestType.includes('newly created position')) {
          logToSheet('Processing Newly Created Position logic.');
          const division = rowData[headerMap.get('Division')];
          const section = rowData[headerMap.get('Section')];
          logToSheet(`Generating new position ID for Division: ${division}, Section: ${section}`);
          const newPositionId = generateNewPositionId(division, section);
          logToSheet(`Generated new Position ID: ${newPositionId}`);
          if (newPositionId.startsWith('ERROR')) {
            throw new Error('Could not generate new Position ID: ' + newPositionId);
          }

          const reportingToId = rowData[headerMap.get('ReportingToId')];
          let reportingToName = '';

          // Get main sheet data to find the manager's name
          const mainSheet = ss.getSheets()[0];
          const mainData = mainSheet.getDataRange().getValues();
          const mainHeaders = mainData[0];
          const mainEmpIdIndex = mainHeaders.indexOf('Employee ID');
          const mainEmpNameIndex = mainHeaders.indexOf('Employee Name');

          if (mainEmpIdIndex > -1 && mainEmpNameIndex > -1) {
              for (let j = 1; j < mainData.length; j++) {
                  // Ensure case-insensitive and trim comparison
                  if (String(mainData[j][mainEmpIdIndex] || '').trim().toUpperCase() === String(reportingToId || '').trim().toUpperCase()) {
                      reportingToName = mainData[j][mainEmpNameIndex];
                      logToSheet(`Found manager name "${reportingToName}" for manager ID "${reportingToId}".`);
                      break;
                  }
              }
          }
          if (!reportingToName) {
              logToSheet(`WARNING: Could not find a name for manager with ID "${reportingToId}".`);
          }
          
          dataToSave = {
            positionid: newPositionId,
            jobtitle: rowData[headerMap.get('NewJobTitle')],
            level: rowData[headerMap.get('NewLevel')],
            division: division,
            group: rowData[headerMap.get('Group')],
            department: rowData[headerMap.get('Department')],
            section: section,
            reportingtoid: reportingToId,
            reportingto: reportingToName, // Add the manager's name
            status: 'VACANT',
            employeename: '',
            employeeid: ''
            reportingtoid: rowData[headerMap.get('ReportingToId')],
            status: 'VACANT',
            employeename: '',
            employeeid: '',
            positionstatus: rowData[headerMap.get('PositionStatus')] || 'Active'
          };
          mode = 'add';
        }
        
        logToSheet(`Data to save: ${JSON.stringify(dataToSave)}, mode: ${mode}`);

        // Call the existing save function to apply the change
        if (Object.keys(dataToSave).length > 0) {
          logToSheet(`Attempting to save data for request ${requestId}...`);
          const saveResult = saveEmployeeData(dataToSave, mode);
          logToSheet(`Save operation completed for request ${requestId}. Result: ${saveResult}`);

          // It's good practice to check the result, even if saveEmployeeData currently throws errors on failure.
          // This makes the code more robust if saveEmployeeData is changed to return a status object in the future.
          if (saveResult.includes('successfully')) {
            logToSheet('Updating implementation details in "Org Chart Requests" sheet.');
            sheet.getRange(i + 1, headerMap.get('Status') + 1).setValue('Implemented');
            sheet.getRange(i + 1, headerMap.get('ImplementerEmail') + 1).setValue(Session.getActiveUser().getEmail());
            sheet.getRange(i + 1, headerMap.get('ImplementationTimestamp') + 1).setValue(new Date());
            logToSheet('Implementation details updated.');
          } else {
             // If saveEmployeeData returns an error message instead of throwing an error.
            throw new Error(`Save operation failed for request ${requestId}: ${saveResult}`);
          }
        } else {
            logToSheet('No data to save for this request type.');
        }

        SpreadsheetApp.flush(); // Ensure all pending changes are written before returning
        return { success: true, message: `Request ${requestId} has been implemented successfully.` };
      }
    }

    if (!requestFound) {
      throw new Error(`Request ID ${requestId} not found for implementation.`);
    }

  } catch (e) {
    logToSheet('FATAL Error in implementApprovedChange: ' + e.message + ' Stack: ' + e.stack);
    return { success: false, error: 'Failed to implement request. ' + e.message };
  }
}

function getRequestCounts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Org Chart Requests');
    if (!sheet || sheet.getLastRow() < 2) {
      return { myRequestsPending: 0, myRequestsRejected: 0, approvals: 0 };
    }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusIndex = headers.indexOf('Status');
    const requestorEmailIndex = headers.indexOf('RequestorEmail');
    const approverEmailIndex = headers.indexOf('ApproverEmail');
    const userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();

    let myRequestsPending = 0;
    let myRequestsRejected = 0;
    let approvalsCount = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = (row[statusIndex] || '').toString().toLowerCase().trim();

      // Count for "My Requests" tab
      if ((row[requestorEmailIndex] || '').toString().toLowerCase().trim() === userEmail) {
        if (status === 'pending') {
          myRequestsPending++;
        } else if (status === 'rejected') {
          myRequestsRejected++;
        }
      }

      // Count for "Approvals" tab (only Pending)
      if (status === 'pending' && (row[approverEmailIndex] || '').toString().toLowerCase().trim() === userEmail) {
        approvalsCount++;
      }
    }

    return { myRequestsPending, myRequestsRejected, approvals: approvalsCount };

  } catch (e) {
    Logger.log('Error in getRequestCounts: ' + e.message);
    // In case of error, return zero counts to avoid breaking the UI
    return { myRequestsPending: 0, myRequestsRejected: 0, approvals: 0 };
  }
}

function getChangeRequests() {
  try {
    Logger.log('getChangeRequests function started.');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Org Chart Requests');
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log('Sheet "Org Chart Requests" not found or empty.');
      return { myRequests: [], approvals: [] };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Keep original headers for object keys
    const userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();

    // Create a normalized map to find column indices reliably
    const normalizedHeaderMap = new Map();
    headers.forEach((header, i) => {
      const normalizedKey = (header || '').toString().toLowerCase().replace(/\s+/g, '');
      if (normalizedKey) {
        normalizedHeaderMap.set(normalizedKey, i);
      }
    });

    // Get indices using the normalized map
    const requestorEmailIndex = normalizedHeaderMap.get('requestoremail');
    const approverEmailIndex = normalizedHeaderMap.get('approveremail');
    const statusIndex = normalizedHeaderMap.get('status');
    const supportDocIndex = normalizedHeaderMap.get('supportingdocuments');
    const submissionTimestampIndex = normalizedHeaderMap.get('submissiontimestamp');

    if (requestorEmailIndex === undefined || approverEmailIndex === undefined || statusIndex === undefined) {
      const errorMessage = "Missing required columns (RequestorEmail, ApproverEmail, or Status). Please check the 'Org Chart Requests' sheet headers.";
      Logger.log(errorMessage + " Headers found: " + headers.join(', '));
      throw new Error(errorMessage);
    }

    const requests = data.map(row => {
      const request = {};
      headers.forEach((header, i) => {
        // Use original headers for the keys in the final objects
        if (i === supportDocIndex && row[i]) {
          request[header] = `<a href="${row[i]}" target="_blank">View Documents</a>`;
        } else if (row[i] instanceof Date) {
          // Format all dates consistently
          request[header] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        } else {
          request[header] = row[i];
        }
      });
      return request;
    });

    // Use original header names for filtering, which are now guaranteed to exist
    const requestorEmailHeader = headers[requestorEmailIndex];
    const approverEmailHeader = headers[approverEmailIndex];
    const statusHeader = headers[statusIndex];
    const submissionTimestampHeader = headers[submissionTimestampIndex];

    const myRequests = requests.filter(r =>
      (r[requestorEmailHeader] || '').toString().toLowerCase().trim() === userEmail
    );

    const approvals = requests.filter(r => {
      const status = (r[statusHeader] || '').toString().toLowerCase().trim();
      return (status === 'pending' || status === 'approved' || status === 'implemented') &&
             (r[approverEmailHeader] || '').toString().toLowerCase().trim() === userEmail;
    });

    // Sort using the dynamically found timestamp header
    const sortByTimestampDesc = (a, b) => {
      const dateA = a[submissionTimestampHeader] ? new Date(a[submissionTimestampHeader]) : new Date(0);
      const dateB = b[submissionTimestampHeader] ? new Date(b[submissionTimestampHeader]) : new Date(0);
      return dateB - dateA;
    };

    myRequests.sort(sortByTimestampDesc);
    approvals.sort(sortByTimestampDesc);

    const resultObject = { myRequests, approvals };
    Logger.log('Successfully prepared data. My Requests: ' + myRequests.length + ', Approvals: ' + approvals.length);
    return resultObject;

  } catch (e) {
    Logger.log('FATAL Error in getChangeRequests: ' + e.message + ' Stack: ' + e.stack);
    // Return a safe, empty object to prevent frontend errors
    return { myRequests: [], approvals: [] };
  }
}


function getEmployeeDetails(employeeName) {
  if (!employeeName) {
    return null;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheets()[0];
    if (mainSheet.getLastRow() < 2) {
      return null;
    }
    const data = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    const empNameIndex = headers.indexOf('Employee Name');
    const empIdIndex = headers.indexOf('Employee ID');
    const dateHiredIndex = headers.indexOf('Date Hired');
    const dobIndex = headers.indexOf('Date of Birth');

    if (empNameIndex === -1 || empIdIndex === -1 || dateHiredIndex === -1 || dobIndex === -1) {
      return null;
    }

    const employeeRow = data.find(row => (row[empNameIndex] || '').toString().trim() === employeeName.trim());

    if (employeeRow) {
      const dateHired = employeeRow[dateHiredIndex] instanceof Date ? Utilities.formatDate(employeeRow[dateHiredIndex], Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
      const dateOfBirth = employeeRow[dobIndex] instanceof Date ? Utilities.formatDate(employeeRow[dobIndex], Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
      return {
        employeeId: employeeRow[empIdIndex],
        dateHired: dateHired,
        dateOfBirth: dateOfBirth
      };
    }
    return null;
  } catch (e) {
    Logger.log(`Error in getEmployeeDetails: ${e.toString()}`);
    return null;
  }
}


/**
 * DIAGNOSTIC FUNCTION
 */
function debugIncumbencyForPosition() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Debug Incumbency History',
    'Please enter the exact Position ID to debug:',
    ui.ButtonSet.OK_CANCEL);


  const button = result.getSelectedButton();
  const posId = result.getResponseText();


  if (button !== ui.Button.OK || !posId) {
    return; // User cancelled
  }


  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('change_log_sheet');
    if (!logSheet || logSheet.getLastRow() < 2) {
      ui.alert('The "change_log_sheet" is empty or not found.');
      return;
    }


    const allLogData = logSheet.getDataRange().getValues();
    const headers = allLogData.shift();


    const posIdIndex = headers.indexOf('Position ID');
    const nameIndex = headers.indexOf('Employee Name');
    const timestampIndex = headers.indexOf('Change Timestamp');
    const effectiveDateIndex = headers.indexOf('Effective Date');


    if ([posIdIndex, nameIndex, timestampIndex, effectiveDateIndex].includes(-1)) {
      throw new Error("One or more required columns (Position ID, Employee Name, Change Timestamp, Effective Date) are missing from the change_log_sheet.");
    }


    const positionEntries = allLogData
      .filter(row => row[posIdIndex] === posId && row[timestampIndex])
      .sort((a, b) => {
        const dateA = a[effectiveDateIndex] instanceof Date ? a[effectiveDateIndex] : new Date(a[timestampIndex]);
        const dateB = b[effectiveDateIndex] instanceof Date ? b[effectiveDateIndex] : new Date(b[timestampIndex]);
        return dateA - dateB;
      });


    if (positionEntries.length === 0) {
      Logger.log(`No log entries found for Position ID: ${posId}`);
      ui.alert(`No log entries were found for Position ID: "${posId}". Please check the ID and try again.`);
      return;
    }


    Logger.log(`--- DEBUG LOG FOR POSITION ID: ${posId} ---`);
    Logger.log(`Found ${positionEntries.length} entries. Sorted by true effective date:`);
    Logger.log('--------------------------------------------------');


    positionEntries.forEach((entry, index) => {
      const definitiveDate = entry[effectiveDateIndex] || entry[timestampIndex];
      const logLine = `Event #${index + 1} (Effective: ${new Date(definitiveDate).toLocaleDateString()}): ` +
        `Incumbent: "${entry[nameIndex]}", ` +
        `Effective Date (Col AA): "${entry[effectiveDateIndex]}"`;
      Logger.log(logLine);
    });


    Logger.log('--------------------------------------------------');
    Logger.log(`--- END OF DEBUG LOG ---`);


    ui.alert('Debug log created successfully. Please go to the Apps Script editor and view the logs under "Executions".');


  } catch (e) {
    Logger.log(`Error during debug: ${e.toString()}`);
    ui.alert(`An error occurred: ${e.message}`);
  }
}


function initializeChangeLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'change_log_sheet';
  if (ss.getSheetByName(sheetName)) {
    SpreadsheetApp.getUi().alert(`A sheet named "${sheetName}" already exists.`);
    return;
  }

  const sheet = ss.insertSheet(sheetName);
  const headers = [
    "Position ID", "Employee ID", "Employee Name", "Reporting to ID", "Reporting to", 
    "Job Title", "Division", "Group", "Department", "Section", "Gender", 
    "Level", "Payroll Type", "Job Level", "Contract Type", "Competency", 
    "Status", "Position Status", "Date Hired", "Date of Birth", "Contract End Date", // Added "Date of Birth"
    "Change Timestamp", "Change Type", "Transfer Note", "Effective Date",
    "Division Headcount", "Department Headcount", "Section Headcount"
  ];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert(`Successfully created the "${sheetName}" sheet.`);
}


// --- ALL OTHER FUNCTIONS ARE BELOW ---


function initializeChangeTracking() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheets()[0];
  if (mainSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Your data sheet is empty. Please add data before initializing.');
    return;
  }
  try {
    const lastCol = mainSheet.getLastColumn();
    const data = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, lastCol).getValues();


    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('lastKnownData', JSON.stringify(data));
    scriptProperties.setProperty('lastKnownColumnCount', lastCol.toString());
    scriptProperties.setProperty('incumbencyHistory', JSON.stringify({}));
    scriptProperties.setProperty('snapshotTimestamp', '');


    SpreadsheetApp.getUi().alert('Success! The real-time change log and incumbency tracking systems have been initialized.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Initialization failed. Error: ' + e.message);
  }
}


function handleSheetChange(e) {
  if (['EDIT', 'INSERT_ROW', 'REMOVE_ROW'].indexOf(e.changeType) === -1) {
    return;
  }
  const lock = LockService.getScriptLock();
  // Attempt to acquire the lock for a short period.
  // If it fails, it means another process (like implementApprovedChange) is holding it.
  // In that case, we should exit and let the main process handle the logic.
  if (lock.tryLock(500)) {
    try {
      logDataChanges();
    } finally {
      lock.releaseLock();
    }
  } else {
    Logger.log('Skipping handleSheetChange because another process is holding the lock.');
  }
}


/**
 * Invalidates the incumbency history cache for specific positions.
 * @param {string[]} positionIds - An array of Position IDs to clear from the cache.
 */
function invalidateIncumbencyCache(positionIds) {
  if (!positionIds || positionIds.length === 0) return;
  const cache = CacheService.getScriptCache();
  const cacheKeys = positionIds.map(id => `incumbency_history_${id}`);
  cache.removeAll(cacheKeys);
  Logger.log(`Invalidated incumbency cache for Position IDs: ${positionIds.join(', ')}`);
}


function logDataChanges() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  if (sheets.length === 0) {
    return;
  }

  const mainSheet = sheets[0];
  const logSheet = spreadsheet.getSheetByName('change_log_sheet');
  if (!logSheet || logSheet.getLastRow() < 1) return;

  const scriptProperties = PropertiesService.getScriptProperties();
  const lastKnownDataString = scriptProperties.getProperty('lastKnownData');
  if (!lastKnownDataString) return;

  const logSheetHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  const logHeaderMap = new Map(logSheetHeaders.map((h, i) => [h.trim(), i]));
  const mainSheetHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];

  const pendingResignationPosId = scriptProperties.getProperty('pendingResignationPosId');
  const pendingResignationDate = scriptProperties.getProperty('pendingResignationDate');
  const pendingEffectiveDatePosId = scriptProperties.getProperty('pendingEffectiveDatePosId');
  const pendingEffectiveDate = scriptProperties.getProperty('pendingEffectiveDate');
  const overrideTimestamp = scriptProperties.getProperty('overrideTimestamp');
  const isCorrection = scriptProperties.getProperty('isResignationDateCorrection');

  let timestamp = new Date();
  if (overrideTimestamp) {
    timestamp = new Date(overrideTimestamp);
    scriptProperties.deleteProperty('overrideTimestamp');
  }

  const incumbencyHistory = JSON.parse(scriptProperties.getProperty('incumbencyHistory') || '{}');
  const previousData = JSON.parse(lastKnownDataString);
  const currentData = mainSheet.getLastRow() > 1 ? mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues() : [];

  const currentDataMap = new Map(currentData.map(row => [row[0], row]));
  const previousDataMap = new Map(previousData.map(row => [row[0], row]));
  const previousEmployeeMap = new Map();
  previousData.forEach(row => {
    if (row[1]) previousEmployeeMap.set(String(row[1]).trim(), row);
  });

  const changesToLog = [];
  previousDataMap.forEach((prevRow, posId) => {
    const currentRow = currentDataMap.get(posId);
    if (!currentRow) {
      changesToLog.push(prevRow.concat([timestamp, 'Row Deleted', '']));
    } else if (JSON.stringify(prevRow) !== JSON.stringify(currentRow) || (isCorrection && posId === pendingEffectiveDatePosId)) {
      let internalTransferNote = '';
      if (currentRow[1] && currentRow[1] !== prevRow[1]) {
        const oldPositionRow = previousEmployeeMap.get(String(currentRow[1]).trim());
        if (oldPositionRow && oldPositionRow[0] !== posId) {
          internalTransferNote = `From: ${oldPositionRow[8] || 'N/A'} (${oldPositionRow[9] || 'N/A'}) - ${oldPositionRow[5] || 'N/A'}`;
        }
      }
      if (prevRow[1] && !currentRow[1] && prevRow[2]) {
        if (!incumbencyHistory[posId]) incumbencyHistory[posId] = [];
        incumbencyHistory[posId].unshift(prevRow[2]);
        incumbencyHistory[posId] = incumbencyHistory[posId].slice(0, 10);
      }
      
      if (prevRow[1] && !currentRow[1]) {
        changesToLog.push(prevRow.concat([timestamp, 'Row Modified', internalTransferNote]));
      } else {
        changesToLog.push(currentRow.concat([timestamp, 'Row Modified', internalTransferNote]));
      }
    }
  });

  currentDataMap.forEach((currentRow, posId) => {
    if (!previousDataMap.has(posId)) {
      let internalTransferNote = '';
      if (currentRow[1]) {
        const oldPositionRow = previousEmployeeMap.get(String(currentRow[1]).trim());
        if (oldPositionRow) {
          internalTransferNote = `From: ${oldPositionRow[8] || 'N/A'} (${oldPositionRow[9] || 'N/A'}) - ${oldPositionRow[5] || 'N/A'}`;
        }
      }
      changesToLog.push(currentRow.concat([timestamp, 'Row Added', internalTransferNote]));
    }
  });

  if (changesToLog.length > 0) {
    const modifiedPositionIds = [...new Set(changesToLog.map(row => row[0]).filter(String))];
    invalidateIncumbencyCache(modifiedPositionIds);

    const logData = changesToLog.map(function(changedRow) {
      const newLogRow = Array(logSheetHeaders.length).fill('');
      const changeType = changedRow[changedRow.length - 2];
      const posId = changedRow[0];
      const empId = changedRow[1];

      mainSheetHeaders.forEach((header, i) => {
        if (logHeaderMap.has(header.trim())) {
          newLogRow[logHeaderMap.get(header.trim())] = changedRow[i];
        }
      });

      const headcount = getCurrentHeadcounts(changedRow[6], changedRow[8], changedRow[9], currentData);
      if (logHeaderMap.has('Change Type')) newLogRow[logHeaderMap.get('Change Type')] = changeType;
      if (logHeaderMap.has('Transfer Note')) newLogRow[logHeaderMap.get('Transfer Note')] = changedRow[changedRow.length - 1];
      if (logHeaderMap.has('Change Timestamp')) newLogRow[logHeaderMap.get('Change Timestamp')] = changedRow[changedRow.length - 3];
      if (logHeaderMap.has('Division Headcount')) newLogRow[logHeaderMap.get('Division Headcount')] = headcount.division;
      if (logHeaderMap.has('Department Headcount')) newLogRow[logHeaderMap.get('Department Headcount')] = headcount.department;
      if (logHeaderMap.has('Section Headcount')) newLogRow[logHeaderMap.get('Section Headcount')] = headcount.section;

      const effectiveDateIndex = logHeaderMap.get('Effective Date');
      if (effectiveDateIndex !== undefined) {
        // This is the definitive vacating event from a promotion/transfer.
        // It's identified by the pending property, and we ensure it's only used once
        // by checking that the log entry still contains the employee from the `prevRow`.
        const logContainsEmployee = !!changedRow[mainSheetHeaders.indexOf('Employee ID')];

        if (pendingResignationPosId && pendingResignationDate && posId.toUpperCase() === pendingResignationPosId.toUpperCase() && logContainsEmployee) {
          newLogRow[effectiveDateIndex] = new Date(pendingResignationDate);

          // CRITICAL: Immediately delete the properties after using them once.
          // This prevents a subsequent cascading update (e.g., reporting line change)
          // from creating a second log entry for the same position and incorrectly
          // reusing the promotion date.
          scriptProperties.deleteProperty('pendingResignationPosId');
          scriptProperties.deleteProperty('pendingResignationDate');
        }
        // This handles other scenarios like a direct resignation from the form.
        else if (pendingEffectiveDatePosId && pendingEffectiveDate && posId.toUpperCase() === pendingEffectiveDatePosId.toUpperCase()) {
          newLogRow[effectiveDateIndex] = new Date(pendingEffectiveDate);
        }
      }
      return newLogRow;
    });

    if (pendingEffectiveDatePosId) {
      scriptProperties.deleteProperty('pendingEffectiveDatePosId');
      scriptProperties.deleteProperty('pendingEffectiveDate');
    }
    if (pendingResignationPosId) {
      scriptProperties.deleteProperty('pendingResignationPosId');
      scriptProperties.deleteProperty('pendingResignationDate');
    }
    
    if (isCorrection) {
      scriptProperties.deleteProperty('isResignationDateCorrection');
    }

    if (logData.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, logData.length, logData[0].length).setValues(logData);
    }
  }

  PropertiesService.getScriptProperties().setProperty('lastKnownData', JSON.stringify(currentData));
  PropertiesService.getScriptProperties().setProperty('lastKnownColumnCount', String(mainSheet.getLastColumn()));
  PropertiesService.getScriptProperties().setProperty('incumbencyHistory', JSON.stringify(incumbencyHistory));
}

function getCurrentHeadcounts(division, department, section, allData) {
  let divisionCount = 0;
  let departmentCount = 0;
  let sectionCount = 0;
  for (let i = 0; i < allData.length; i++) {
    if ((allData[i][17] || '').toString().trim().toLowerCase() === 'inactive') continue;
    if (allData[i][6] === division) {
      divisionCount++;
      if (allData[i][8] === department) {
        departmentCount++;
        if (allData[i][9] === section) {
          sectionCount++;
        }
      }
    }
  }
  return {
    division: divisionCount,
    department: departmentCount,
    section: sectionCount
  };
}


function takeHeadcountSnapshotWithAlert() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirm', 'This will update the "Previous Headcount" summary and create new approval records for all departments. Continue?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    try {
      takeHeadcountSnapshot();
      ui.alert('Success! The headcount summary has been updated and new approval records have been created for each department.');
    } catch (e) {
      ui.alert('Error: ' + e.message);
    }
  }
}


function takeHeadcountSnapshot() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheets()[0];
  let targetSheet = spreadsheet.getSheetByName('Previous Headcount');

  if (!targetSheet) {
    targetSheet = spreadsheet.insertSheet('Previous Headcount');
    targetSheet.appendRow(['Division', 'Group', 'Department', 'Section', 'Approved Plantilla']);
    targetSheet.setFrozenRows(1);
  }

  if (mainSheet.getLastRow() < 2) {
    return;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  const currentSnapshot = scriptProperties.getProperty('snapshotTimestamp');
  if (currentSnapshot) {
    scriptProperties.setProperty('previousHeadcountTimestamp', currentSnapshot);
  }

  const timestamp = new Date();
  scriptProperties.setProperty('snapshotTimestamp', timestamp.toISOString());

  const data = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 18).getValues();
  const approvalsSheet = spreadsheet.getSheetByName('Approvals');
  if (!approvalsSheet) {
    throw new Error('Sheet "Approvals" not found.');
  }

  const approversData = getApproversData();
  const uniqueDepartments = [...new Set(data.map(row => row[8]).filter(String))];
  const existingApprovalRecords = approvalsSheet.getDataRange().getValues();
  const headers = existingApprovalRecords.length > 0 ? existingApprovalRecords[0] : [];
  const snapshotColIndex = headers.indexOf('Snapshot Date');
  const deptColIndex = headers.indexOf('Department');
  const newlyCreatedRecords = [];

  uniqueDepartments.forEach(dept => {
    const recordExists = existingApprovalRecords.some((row, index) =>
      index > 0 && row[snapshotColIndex] === timestamp.toISOString() && row[deptColIndex] === dept
    );
    if (!recordExists) {
      approvalsSheet.appendRow([timestamp.toISOString(), dept, '', '', '', '', '', '', '', '']);
      newlyCreatedRecords.push(dept);
    }
  });

  newlyCreatedRecords.forEach(dept => {
    sendApprovalNotificationEmail(dept, timestamp.toISOString(), approversData, 'Prepared By');
  });

  const summary = {};
  data.forEach(function (row) {
    if ((row[17] || '').toString().trim().toLowerCase() === 'inactive') return;
    const isFilled = !!row[1];
    const division = row[6],
      group = row[7] || '', // Ensure blank values are treated as empty strings
      department = row[8] || '',
      section = row[9] || '';

    if (!division) return;
    if (!summary[division]) summary[division] = {
      filled: 0,
      vacant: 0,
      groups: {}
    };
    if (!summary[division].groups[group]) summary[division].groups[group] = {
      filled: 0,
      vacant: 0,
      departments: {}
    };
    if (!summary[division].groups[group].departments[department]) summary[division].groups[group].departments[department] = {
      filled: 0,
      vacant: 0,
      sections: {}
    };
    if (!summary[division].groups[group].departments[department].sections[section]) summary[division].groups[group].departments[department].sections[section] = {
      filled: 0,
      vacant: 0
    };

    isFilled ? summary[division].filled++ : summary[division].vacant++;
    isFilled ? summary[division].groups[group].filled++ : summary[division].groups[group].vacant++;
    isFilled ? summary[division].groups[group].departments[department].filled++ : summary[division].groups[group].departments[department].vacant++;
    isFilled ? summary[division].groups[group].departments[department].sections[section].filled++ : summary[division].groups[group].departments[department].sections[section].vacant++;
  });


  const monthHeader = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "MMM yyyy");
  const filledHeader = `${monthHeader} Filled`;
  const vacantHeader = `${monthHeader} Vacant`;

  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  let filledColIdx = targetHeaders.indexOf(filledHeader);
  let vacantColIdx = targetHeaders.indexOf(vacantHeader);
  let plantillaColIdx = targetHeaders.indexOf('Approved Plantilla');

  if (plantillaColIdx === -1) {
    targetSheet.getRange(1, 5).setValue('Approved Plantilla');
    plantillaColIdx = 4;
  }

  if (filledColIdx === -1) {
    const lastCol = targetSheet.getLastColumn();
    targetSheet.getRange(1, lastCol + 1, 1, 2).setValues([
      [filledHeader, vacantHeader]
    ]);
    filledColIdx = lastCol;
    vacantColIdx = lastCol + 1;
  }

  const existingData = targetSheet.getLastRow() > 1 ? targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).getValues() : [];
  const dataMap = new Map();
  existingData.forEach((row, index) => {
    const key = [row[0], row[1], row[2], row[3]].join('|');
    dataMap.set(key, {
      rowIndex: index + 2,
      data: row
    });
  });

  const updatedData = [];

  const processLevel = (div, group, dept, sec, counts) => {
    const key = [div, group, dept, sec].join('|');
    if (dataMap.has(key)) {
      const existingRow = dataMap.get(key);
      existingRow.data[filledColIdx] = counts.filled;
      existingRow.data[vacantColIdx] = counts.vacant;
      updatedData.push({
        range: `A${existingRow.rowIndex}`,
        values: [existingRow.data]
      });
      dataMap.delete(key);
    } else {
      const newRow = Array(targetSheet.getLastColumn()).fill('');
      newRow[0] = div;
      newRow[1] = group;
      newRow[2] = dept;
      newRow[3] = sec;
      newRow[filledColIdx] = counts.filled;
      newRow[vacantColIdx] = counts.vacant;
      targetSheet.appendRow(newRow);
    }
  };

  // --- REVISED SECTION ---
  // This revised logic filters out empty keys ('') before processing,
  // preventing the creation of blank or incomplete rows in the "Previous Headcount" sheet.
  Object.keys(summary).sort().forEach(divName => {
    processLevel(divName, '', '', '', summary[divName]); // Process Division total
    Object.keys(summary[divName].groups).sort().filter(g => g).forEach(groupName => { // Filter out empty group names
      processLevel(divName, groupName, '', '', summary[divName].groups[groupName]); // Process Group total
      Object.keys(summary[divName].groups[groupName].departments).sort().filter(d => d).forEach(deptName => { // Filter out empty dept names
        processLevel(divName, groupName, deptName, '', summary[divName].groups[groupName].departments[deptName]); // Process Dept total
        Object.keys(summary[divName].groups[groupName].departments[deptName].sections).sort().filter(s => s).forEach(secName => { // Filter out empty section names
          processLevel(divName, groupName, deptName, secName, summary[divName].groups[groupName].departments[deptName].sections[secName]); // Process Section total
        });
      });
    });
  });
  // --- END REVISED SECTION ---

  updatedData.forEach(update => {
    const range = targetSheet.getRange(update.range).offset(0, 0, 1, update.values[0].length);
    range.setValues(update.values);
  });
}


function getApproversData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const approversSheet = spreadsheet.getSheetByName('Approvers');
  const allApprovers = {};


  if (approversSheet) {
    const data = approversSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data.shift();
      data.forEach((row) => {
        const department = row[0] ? row[0].toString().trim() : '';
        const role = row[1] ? row[1].toString().trim() : '';
        const email = row[2] ? row[2].toString().trim() : '';
        if (department && role && email) {
          if (!allApprovers[department]) {
            allApprovers[department] = {};
          }
          allApprovers[department][role] = email;
        }
      });
    }
  }
  return allApprovers;
}


function sendApprovalNotificationEmail(department, snapshotTimestamp, allApproversData, roleToNotify) {
  const departmentApprovers = allApproversData[department];
  if (!departmentApprovers) {
    return;
  }
  const recipientEmail = departmentApprovers[roleToNotify];
  if (recipientEmail) {
    const subject = `Approval Required (${roleToNotify}): Org Chart Snapshot for ${department}`;
    const body = `Dear ${recipientEmail.split('@')[0].toUpperCase()},\n\nThe Organizational Chart snapshot for your department (${department}) generated on ${new Date(snapshotTimestamp).toLocaleString("en-US",{timeZone:"Asia/Manila"})} requires your signature as "${roleToNotify}".\n\nPlease visit the Organizational Chart web application to sign:\n${WEB_APP_URL}\n\nThank you,\nYour Organizational Chart Team`;
    try {
      MailApp.sendEmail(recipientEmail, subject, body);
    } catch (mailError) {
      Logger.log(`ERROR sending email to ${roleToNotify} (${recipientEmail}) for department ${department}. Error: ${mailError.message}`);
    }
  }
}


function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Interactive Organizational Chart').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function getIncumbencyHistory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const historyString = scriptProperties.getProperty('incumbencyHistory');
  return historyString ? JSON.parse(historyString) : {};
}


function getEmployeeData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const mainSheet = spreadsheet.getSheets()[0];


    const logSheet = spreadsheet.getSheetByName('change_log_sheet');
    const resignationDates = new Map();
    if (logSheet && logSheet.getLastRow() > 1) {
      const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();
      const headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
      const posIdIndex = headers.indexOf('Position ID');
      const statusIndex = headers.indexOf('Status');
      const effectiveDateIndex = headers.indexOf('Effective Date');


      if (posIdIndex > -1 && statusIndex > -1 && effectiveDateIndex > -1) {
        logData.forEach(row => {
          if (row[posIdIndex] && String(row[statusIndex]).toUpperCase() === 'RESIGNED' && row[effectiveDateIndex] instanceof Date) {
            resignationDates.set(row[posIdIndex], row[effectiveDateIndex]);
          }
        });
      }
    }


    const userPermissions = {};
    const permissionsSheet = spreadsheet.getSheetByName('Permissions');
    if (permissionsSheet) {
      const permData = permissionsSheet.getDataRange().getValues();
      if (permData.length > 0) {
        const permissionHeaders = permData.shift();
        const emailColIndex = permissionHeaders.indexOf('EMAIL');
        if (emailColIndex !== -1) {
          const userRow = permData.find(row => row[emailColIndex] && row[emailColIndex].toString().trim().toLowerCase() === userEmail);
          if (userRow) {
            permissionHeaders.forEach((header, index) => {
              if (header) {
                userPermissions[header.trim()] = userRow[index] ? userRow[index].toString().trim().toLowerCase() : '';
              }
            });
          }
        }
      }
    }
    const isFieldAuthorized = (fieldName) => (userPermissions[fieldName] === 'x' || userPermissions[fieldName] === 'all' || userPermissions[fieldName] === 'anyone');
    const isDepartmentViewable = (employeeDivision, employeeDepartment) => {
      const viewableDeptEntry = userPermissions['Viewable Department'] || '';
      if (viewableDeptEntry === 'all' || viewableDeptEntry === 'anyone') return true;
      const allowedDeptDivs = viewableDeptEntry.split(',').map(item => item.trim().toLowerCase()).filter(item => item);
      return allowedDeptDivs.includes(employeeDepartment.toLowerCase()) || allowedDeptDivs.includes(employeeDivision.toLowerCase());
    };
    const canEdit = userPermissions['Can Edit'] === 'x' || userPermissions['Can Edit'] === 'all' || userPermissions['Can Edit'] === 'anyone';
    const canApprove = userPermissions['Is Approver'] === 'x' || userPermissions['Is Approver'] === 'all' || userPermissions['Is Approver'] === 'anyone';


    if (mainSheet.getLastRow() < 2) {
      return {
        current: [], previous: {}, snapshotTimestamp: null, currentUserEmail: userEmail,
        userCanSeeAnyDepartment: false, totalApprovedPlantilla: 0, previousDateString: null, canEdit: canEdit, canApprove: canApprove
      };
    }


    const lastCol = Math.max(21, mainSheet.getLastColumn());
    const mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, lastCol).getValues();
    const employeeIdToPositionIdMap = new Map();

    // --- START OPTIMIZATION ---
    // In-line the logic from getListsForDropdowns to avoid a second sheet read and multiple loops.
    const activeEmployees = [];
    const divisions = new Set();
    const groups = new Set();
    const departments = new Set();
    const sections = new Set();

    mainData.forEach(row => {
      const employeeId = row[1] ? row[1].toString().trim() : null;
      const positionId = row[0] ? row[0].toString().trim() : null;
      if (employeeId && positionId) {
        employeeIdToPositionIdMap.set(employeeId, positionId);
      }
      // Populate dropdown list data in the same loop
      if (employeeId && (row[17] || '').toLowerCase() !== 'inactive') {
          activeEmployees.push({ id: employeeId, name: row[2] });
      }
      if(row[6]) divisions.add(row[6]);
      if(row[7]) groups.add(row[7]);
      if(row[8]) departments.add(row[8]);
      if(row[9]) sections.add(row[9]);
    });

    const refSheet = spreadsheet.getSheetByName("Reference Data");
    let staticLists = {};
    if (refSheet) {
        const refData = refSheet.getDataRange().getValues();
        const headers = refData.shift();
        headers.forEach((header, colIndex) => {
            if (header) {
                const key = header.toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-9]/gi, '');
                const values = refData.map(row => row[colIndex]).filter(String).sort();
                staticLists[key] = values;
            }
        });
    }

    const dropdownListData = {
        managers: activeEmployees.sort((a, b) => a.name.localeCompare(b.name)),
        divisions: [...divisions].sort(),
        groups: [...groups].sort(),
        departments: [...departments].sort(),
        sections: [...sections].sort(),
        ...staticLists
    };
    // --- END OPTIMIZATION ---

    const historicalNotes = getHistoricalNotes();
    const incumbencyHistory = getIncumbencyHistory();
    const employeesToShow = [];
    let hasReturnedAnyEmployee = false;
    mainData.forEach(function (row) {
      const employeeDivision = row[6] ? row[6].toString().trim() : '';
      const employeeDepartment = row[8] ? row[8].toString().trim() : '';
      if (!isDepartmentViewable(employeeDivision, employeeDepartment)) return;
      hasReturnedAnyEmployee = true;
      const posId = row[0] ? row[0].toString().trim() : null;
      if (!posId) return;


      const managerEmployeeId = row[3] ? row[3].toString().trim() : null;
      let managerPositionId = ''; 
      if (managerEmployeeId) {
        if (managerEmployeeId.includes('-')) {
          managerPositionId = managerEmployeeId;
        } else {
          managerPositionId = employeeIdToPositionIdMap.get(managerEmployeeId) || '';
        }
        if (!managerPositionId) {
            Logger.log(`Could not find manager position for employee ${row[2]} (Emp ID: ${row[1]}) who has manager ID: ${managerEmployeeId}`);
        }
      }

      const history = historicalNotes[posId] || {};
      history.lastIncumbents = incumbencyHistory[posId] || [];

      let dateHired = row[18] && row[18] instanceof Date ? Utilities.formatDate(row[18], Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
      let dateOfBirth = row[19] && row[19] instanceof Date ? Utilities.formatDate(row[19], Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
      let contractEndDate = row[20] && row[20] instanceof Date ? Utilities.formatDate(row[20], Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;

      const employeeStatus = row[16] ? row[16].toString().trim() : '';
      let resignationDate = null;
      if (employeeStatus.toUpperCase() === 'RESIGNED' && resignationDates.has(posId)) {
        resignationDate = Utilities.formatDate(resignationDates.get(posId), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      employeesToShow.push({
        positionId: posId,
        employeeId: row[1] ? row[1].toString().trim() : null,
        nodeId: posId,
        employeeName: row[2],
        managerId: managerPositionId || '',
        managerEmployeeId: managerEmployeeId || '',
        managerName: row[4],
        jobTitle: row[5],
        division: employeeDivision,
        group: row[7],
        department: employeeDepartment,
        section: row[9],
        gender: row[10] ? row[10].toString().trim() : '',
        level: row[11],
        payrollType: isFieldAuthorized('Payroll Type') ? row[12] : null,
        jobLevel: isFieldAuthorized('Job Level') ? row[13] : null,
        contractType: isFieldAuthorized('Contract Type') ? (row[14] ? row[14].toString().trim() : null) : null,
        stylingContractType: row[14] ? row[14].toString().trim() : null,
        competency: isFieldAuthorized('Competency') ? row[15] : null,
        status: employeeStatus,
        positionStatus: row[17] ? row[17].toString().trim() : 'Active',
        dateHired: dateHired,
        dateOfBirth: dateOfBirth,
        contractEndDate: contractEndDate,
        historicalNote: history,
        resignationDate: resignationDate
      });
    });


    let previousHeadcount = {};
    let totalApprovedPlantilla = 0;
    let previousDateString = null;


    try {
      const previousSheet = spreadsheet.getSheetByName('Previous Headcount');
      if (previousSheet && previousSheet.getLastRow() > 1) {
        const prevDataRange = previousSheet.getDataRange();
        const prevData = prevDataRange.getValues();
        if (prevData.length > 0) {
          const prevHeaders = prevData.shift();
          const plantillaIndex = prevHeaders.indexOf('Approved Plantilla');
          let lastFilledIndex = -1;
          for (let i = prevHeaders.length - 1; i >= 0; i--) {
            if (String(prevHeaders[i]).includes('Filled')) {
              lastFilledIndex = i;
              break;
            }
          }
          if (lastFilledIndex !== -1) {
            const lastVacantIndex = lastFilledIndex + 1;
            const dateHeader = String(prevHeaders[lastFilledIndex] || '');
            if (dateHeader) {
              previousDateString = dateHeader.replace(/ filled/i, '').trim();
            }
            prevData.forEach(function (row) {
              const division = row[0],
                group = row[1] || '',
                department = row[2] || '',
                section = row[3] || '';
              const rawPlantilla = row[plantillaIndex];
              const plantillaValue = (plantillaIndex !== -1 && rawPlantilla !== '' && !isNaN(rawPlantilla)) ? parseInt(rawPlantilla) : null;
              const filled = row[lastFilledIndex] || 0;
              const vacant = (row.length > lastVacantIndex) ? (row[lastVacantIndex] || 0) : 0;
              if (division) {
                if (!previousHeadcount[division]) {
                  previousHeadcount[division] = {
                    filled: 0,
                    vacant: 0,
                    plantilla: null,
                    groups: {}
                  };
                }
                if (!previousHeadcount[division].groups[group]) {
                  previousHeadcount[division].groups[group] = {
                    filled: 0,
                    vacant: 0,
                    plantilla: null,
                    departments: {}
                  };
                }
                if (!previousHeadcount[division].groups[group].departments[department]) {
                  previousHeadcount[division].groups[group].departments[department] = {
                    filled: 0,
                    vacant: 0,
                    plantilla: null,
                    sections: {}
                  };
                }
                if (!previousHeadcount[division].groups[group].departments[department].sections[section]) {
                  previousHeadcount[division].groups[group].departments[department].sections[section] = {
                    filled: 0,
                    vacant: 0,
                    plantilla: null
                  };
                }
                if (group === '' && department === '' && section === '') {
                  previousHeadcount[division].filled = filled;
                  previousHeadcount[division].vacant = vacant;
                  previousHeadcount[division].plantilla = plantillaValue;
                } else if (group && department === '' && section === '') {
                  previousHeadcount[division].groups[group].filled = filled;
                  previousHeadcount[division].groups[group].vacant = vacant;
                  previousHeadcount[division].groups[group].plantilla = plantillaValue;
                } else if (department && section === '') {
                  previousHeadcount[division].groups[group].departments[department].filled = filled;
                  previousHeadcount[division].groups[group].departments[department].vacant = vacant;
                  previousHeadcount[division].groups[group].departments[department].plantilla = plantillaValue;
                } else if (section) {
                  previousHeadcount[division].groups[group].departments[department].sections[section].filled = filled;
                  previousHeadcount[division].groups[group].departments[department].sections[section].vacant = vacant;
                  previousHeadcount[division].groups[group].departments[department].sections[section].plantilla = plantillaValue;
                }
                if (group === '' && department === '' && section === '' && plantillaValue !== null) {
                  totalApprovedPlantilla += plantillaValue;
                }
              }
            });
          }
        }
      }
    } catch (e) {
      Logger.log('WARNING: Could not parse "Previous Headcount" sheet. Summary data will be unavailable. Error: ' + e.message);
    }


    const snapshotTimestamp = PropertiesService.getScriptProperties().getProperty('snapshotTimestamp');

    return {
      current: employeesToShow.filter(emp => emp.positionId),
      previous: previousHeadcount,
      snapshotTimestamp: snapshotTimestamp,
      previousDateString: previousDateString,
      currentUserEmail: userEmail,
      userCanSeeAnyDepartment: hasReturnedAnyEmployee,
      totalApprovedPlantilla: totalApprovedPlantilla,
      canEdit: canEdit,
      canApprove: canApprove,
      dropdownListData: dropdownListData
    };
  } catch (e) {
    Logger.log('ERROR in getEmployeeData: ' + e.toString() + ' Stack: ' + e.stack);
    return null;
  }
}



function getHistoricalNotes() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = spreadsheet.getSheetByName('change_log_sheet');
  const history = {};
  if (!logSheet || logSheet.getLastRow() < 2) return history;


  const logValues = logSheet.getDataRange().getValues();
  const headers = logValues.shift();
  const posIdIndex = headers.indexOf('Position ID');
  const empIdIndex = headers.indexOf('Employee ID');
  const transferNoteIndex = headers.indexOf('Transfer Note');


  if (posIdIndex === -1 || empIdIndex === -1 || transferNoteIndex === -1) {
    Logger.log("getHistoricalNotes: Could not find required headers in change_log_sheet.");
    return history;
  }


  const filledPositions = new Set(logValues.filter(row => row[empIdIndex]).map(row => row[posIdIndex]));


  logValues.forEach(row => {
    const posId = row[posIdIndex];
    const transferNote = row[transferNoteIndex];
    if (posId) {
      if (!history[posId]) {
        history[posId] = {
          isNewPosition: !filledPositions.has(posId)
        };
      }
      if (transferNote) {
        history[posId].previousRole = transferNote;
      }
    }
  });
  return history;
}


function getApprovalData(department) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const snapshotTimestamp = PropertiesService.getScriptProperties().getProperty('snapshotTimestamp');
    const approvers = getApproversData()[department] || {};
    let approvalStatus = {};
    const approvalsSheet = spreadsheet.getSheetByName('Approvals');
    if (approvalsSheet && snapshotTimestamp) {
      const data = approvalsSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data.shift();
        const snapshotColIndex = headers.indexOf('Snapshot Date');
        const deptColIndex = headers.indexOf('Department');
        const approvalRow = data.find(row => row[snapshotColIndex] === snapshotTimestamp && row[deptColIndex] === department);
        if (approvalRow) {
          headers.forEach((header, index) => {
            const value = approvalRow[index];
            approvalStatus[header] = (value instanceof Date) ? value.toISOString() : value;
          });
        }
      }
    }
    return {
      approvers: approvers,
      approvalStatus: approvalStatus
    };
  } catch (e) {
    Logger.log('ERROR in getApprovalData: ' + e.toString());
    return null;
  }
}


function recordApproval(role, department) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const approvalsSheet = spreadsheet.getSheetByName('Approvals');
    if (!approvalsSheet) throw new Error("Sheet 'Approvals' not found.");
    const user = Session.getActiveUser();
    const userName = user.getUserLoginId().split('@')[0];
    const snapshotTimestamp = PropertiesService.getScriptProperties().getProperty('snapshotTimestamp');
    if (!snapshotTimestamp) throw new Error("No active snapshot found.");
    const data = approvalsSheet.getDataRange().getValues();
    const headers = data[0];
    const snapshotColIndex = headers.indexOf('Snapshot Date');
    const deptColIndex = headers.indexOf('Department');
    const roleColIndex = headers.indexOf(role);
    if (roleColIndex === -1) throw new Error(`Role column "${role}" not found.`);
    for (let i = 1; i < data.length; i++) {
      if (data[i][snapshotColIndex] === snapshotTimestamp && data[i][deptColIndex] === department) {
        approvalsSheet.getRange(i + 1, roleColIndex + 1).setValue(userName);
        approvalsSheet.getRange(i + 1, roleColIndex + 2).setValue(new Date());
        SpreadsheetApp.flush();
        const approversData = getApproversData();
        const currentRoleIndex = APPROVAL_ROLES.indexOf(role);
        const nextRole = APPROVAL_ROLES[currentRoleIndex + 1];
        if (nextRole && !getApprovalData(department).approvalStatus[nextRole]) {
          sendApprovalNotificationEmail(department, snapshotTimestamp, approversData, nextRole);
        }
        return "Approval recorded successfully.";
      }
    }
    throw new Error("Could not find matching approval record.");
  } catch (e) {
    Logger.log('ERROR in recordApproval: ' + e.toString());
    throw new Error('Failed to record approval. ' + e.message);
  }
}


function getListsForDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheets()[0];
  const refSheet = ss.getSheetByName("Reference Data");


  let dynamicLists = {};
  if (mainSheet.getLastRow() > 1) {
    const data = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    const activeEmployees = data
      .filter(row => row[1] && (row[17] || '').toLowerCase() !== 'inactive')
      .map(row => ({
        id: row[1],
        name: row[2]
      }))
      .sort((a, b) => a.name.localeCompare(b.name));


    dynamicLists = {
      managers: activeEmployees,
      divisions: [...new Set(data.map(row => row[6]).filter(String))].sort(),
      groups: [...new Set(data.map(row => row[7]).filter(String))].sort(),
      departments: [...new Set(data.map(row => row[8]).filter(String))].sort(),
      sections: [...new Set(data.map(row => row[9]).filter(String))].sort()
    };
  }


  let staticLists = {};
  if (refSheet) {
    const refData = refSheet.getDataRange().getValues();
    const headers = refData.shift();
    headers.forEach((header, colIndex) => {
      if (header) {
        // Standardize the key: lowercase, remove spaces and special chars.
        // This handles "Reason for Leaving" -> "reasonforleaving"
        const key = header.toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-9]/gi, '');
        const values = refData.map(row => row[colIndex]).filter(String).sort();
        staticLists[key] = values;
      }
    });
  }


  return { ...dynamicLists,
    ...staticLists
  };
}


function generateNewPositionId(division, section) {
  try {
    if (!division || !section) {
      return "ERROR: Division and Section are required.";
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheets()[0];


    const divisionCode = division.split(' ')[0].trim();
    const sectionCode = section.split(' ')[0].trim();


    if (!/^\d+$/.test(divisionCode) || !/^\d+$/.test(sectionCode)) {
      return "ERROR: Division/Section name must start with a numeric code.";
    }


    const prefix = `${divisionCode}-${sectionCode}-`;
    const positionIds = mainSheet.getRange("A2:A").getValues().flat().filter(String);


    let maxSequence = 0;
    positionIds.forEach(id => {
      if (id.startsWith(prefix)) {
        const sequence = parseInt(id.substring(prefix.length), 10);
        if (!isNaN(sequence) && sequence > maxSequence) {
          maxSequence = sequence;
        }
      }
    });


    const newSequence = (maxSequence + 1).toString().padStart(3, '0');
    return prefix + newSequence;
  } catch (e) {
    Logger.log(e);
    return `ERROR: ${e.message}`;
  }
}


function saveEmployeeData(dataObject, mode) {
    logToSheet(`--- saveEmployeeData Started ---`);
    logToSheet(`Mode: ${mode}`);
    logToSheet(`Received dataObject: ${JSON.stringify(dataObject)}`);

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    const scriptProperties = PropertiesService.getScriptProperties();
    try {
        scriptProperties.setProperty('scriptChangeLock', 'true');

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const mainSheet = ss.getSheets()[0];
        logToSheet(`Target sheet: "${mainSheet.getName()}"`);
        const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
        const keyMap = {};
        headers.forEach((header, i) => {
            const key = header.toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-9]/gi, '');
            keyMap[key] = i;
        });
        logToSheet(`Header key map generated: ${JSON.stringify(keyMap)}`);

        // Convert all incoming string data to uppercase for consistency, except for specific fields
        for (const key in dataObject) {
            if (typeof dataObject[key] === 'string') {
                dataObject[key] = dataObject[key].toUpperCase();
            }
        }

        let isTransfer = false;
        let oldPositionIdForTransfer = null;
        let transferredEmployeeId = null;
        let isManualVacate = false;
        let vacatingEmployeeId = null;
        let vacatedPositionId = null;
        let isFillingVacancy = false;
        let filledVacancyPositionId = null;
        let newEmployeeIdForVacancy = null;
        let newEmployeeNameForVacancy = null;

        if (dataObject.employeeid && (dataObject.status.toUpperCase() === 'PROMOTION' || dataObject.status.toUpperCase() === 'INTERNAL TRANSFER' || dataObject.status.toUpperCase() === 'LATERAL TRANSFER')) {
            const allData = mainSheet.getDataRange().getValues();
            const posIdIndex = headers.indexOf('Position ID');
            const empIdIndex = headers.indexOf('Employee ID');
            for (let i = 1; i < allData.length; i++) {
                const row = allData[i];
                if ((String(row[empIdIndex]) || '').toUpperCase() === String(dataObject.employeeid).toUpperCase() && (String(row[posIdIndex]) || '').toUpperCase() !== dataObject.positionid.toUpperCase()) {
                    const oldRowIndex = i + 1;
                    isTransfer = true;
                    oldPositionIdForTransfer = row[posIdIndex];
                    transferredEmployeeId = String(dataObject.employeeid).toUpperCase();
                    if (dataObject.startdateinposition) {
                        scriptProperties.setProperties({
                            'pendingResignationPosId': oldPositionIdForTransfer.toUpperCase(),
                            'pendingResignationDate': dataObject.startdateinposition
                        });
                    }
                    mainSheet.getRange(oldRowIndex, keyMap['employeeid'] + 1).clearContent();
                    mainSheet.getRange(oldRowIndex, keyMap['employeename'] + 1).clearContent();
                    mainSheet.getRange(oldRowIndex, keyMap['gender'] + 1).clearContent();
                    mainSheet.getRange(oldRowIndex, keyMap['datehired'] + 1).clearContent();
                    if (keyMap['dateofbirth'] !== undefined) mainSheet.getRange(oldRowIndex, keyMap['dateofbirth'] + 1).clearContent();
                    mainSheet.getRange(oldRowIndex, keyMap['contractenddate'] + 1).clearContent();
                    mainSheet.getRange(oldRowIndex, keyMap['status'] + 1).setValue('VACANT');
                    break;
                }
            }
        }

        if ((dataObject.status === 'VACANT' || dataObject.status === 'RESIGNED') && dataObject.effectivedate) {
            PropertiesService.getScriptProperties().setProperties({
                'pendingEffectiveDatePosId': dataObject.positionid.toUpperCase(),
                'pendingEffectiveDate': dataObject.effectivedate
            });
        }
        if (dataObject.startdateinposition) {
            PropertiesService.getScriptProperties().setProperty('overrideTimestamp', dataObject.startdateinposition);
        }

        if (mode === 'add') {
            const newRowData = Array(headers.length).fill('');
            for (const key in dataObject) {
                if (keyMap.hasOwnProperty(key)) newRowData[keyMap[key]] = dataObject[key];
            }
            const newPositionId = dataObject.positionid;
            let insertRowIndex = -1;
            if (newPositionId) {
                const positionIdPrefix = newPositionId.substring(0, newPositionId.lastIndexOf('-'));
                if (positionIdPrefix) {
                    const positionIdColValues = mainSheet.getRange("A1:A").getValues().flat();
                    for (let i = positionIdColValues.length - 1; i > 0; i--) {
                        if (positionIdColValues[i] && positionIdColValues[i].startsWith(positionIdPrefix)) {
                            insertRowIndex = i + 1;
                            break;
                        }
                    }
                }
            }
            if (insertRowIndex !== -1) {
                mainSheet.insertRowAfter(insertRowIndex);
                mainSheet.getRange(insertRowIndex + 1, 1, 1, newRowData.length).setValues([newRowData]);
            } else {
                mainSheet.appendRow(newRowData);
            }
        } else if (mode === 'edit') {
            const positionId = dataObject.positionid;
            logToSheet(`EDIT mode: Searching for Position ID "${positionId}" in column A.`);
            const positionIdColValues = mainSheet.getRange("A2:A" + mainSheet.getLastRow()).getValues();
            const rowIndex = positionIdColValues.findIndex(r => r[0] == positionId) + 2;

            if (rowIndex < 2) { // rowIndex will be 1 if not found, since we add 2
              logToSheet(`ERROR: Position ID "${positionId}" not found in column A. Aborting save.`);
              throw new Error(`Position ID ${positionId} not found for editing.`);
            }
            logToSheet(`Position ID found at row index: ${rowIndex}.`);

            const rangeToUpdate = mainSheet.getRange(rowIndex, 1, 1, headers.length);
            logToSheet(`Range to update is: ${rangeToUpdate.getA1Notation()}`);
            const existingRowData = rangeToUpdate.getValues()[0];
            logToSheet(`Existing data in row: ${JSON.stringify(existingRowData)}`);
            const originalStatus = existingRowData[keyMap['status']];
            const originalEmployeeId = existingRowData[keyMap['employeeid']];

            if (dataObject.status.toUpperCase() === 'VACANT' && originalEmployeeId) {
                isManualVacate = true;
                vacatingEmployeeId = originalEmployeeId.toUpperCase();
                vacatedPositionId = positionId;
            }

            if (originalStatus && originalStatus.toUpperCase() === 'VACANT' && dataObject.status && dataObject.status.toUpperCase() !== 'VACANT') {
                isFillingVacancy = true;
                filledVacancyPositionId = positionId;
                newEmployeeIdForVacancy = dataObject.employeeid;
                newEmployeeNameForVacancy = dataObject.employeename;
            }

            if (originalStatus && originalStatus.toUpperCase() === 'VACANT') {
                const isFillingAction = dataObject.status && dataObject.status.toUpperCase() !== 'VACANT' && dataObject.employeeid;
                if (!isFillingAction) {
                    dataObject.employeeid = '';
                    dataObject.employeename = '';
                    dataObject.gender = '';
                    dataObject.datehired = '';
                    dataObject.dateofbirth = '';
                    dataObject.contractenddate = '';
                    dataObject.status = 'VACANT';
                }
            }

            for (const key in dataObject) {
                if (keyMap.hasOwnProperty(key)) {
                    existingRowData[keyMap[key]] = dataObject[key];
                }
            }
            logToSheet(`Data to be written to sheet: ${JSON.stringify(existingRowData)}`);
            rangeToUpdate.setValues([existingRowData]);
            logToSheet(`setValues() called successfully on range ${rangeToUpdate.getA1Notation()}.`);

            if (dataObject.status === 'RESIGNED') {
                let resignationSheet = ss.getSheetByName('Resignation Data');
                if (!resignationSheet) {
                    resignationSheet = ss.insertSheet('Resignation Data');
                    const resignationHeaders = [
                        'Timestamp', 'Position ID', 'Employee ID', 'Employee Name',
                        'Division', 'Group', 'Department', 'Section', 'Job Title',
                        'Level', 'Job Level', 'Gender', 'Contract Type', 'Date Hired', 'Date of Birth', 'Resignation Date', 'Reason for Leaving'
                    ];
                    resignationSheet.appendRow(resignationHeaders);
                    resignationSheet.setFrozenRows(1);
                }
                const resignedEmployeeData = [
                    new Date(), dataObject.positionid, dataObject.employeeid, dataObject.employeename,
                    dataObject.division, dataObject.group, dataObject.department, dataObject.section, dataObject.jobtitle,
                    dataObject.level, dataObject.joblevel, dataObject.gender, dataObject.contracttype,
                    dataObject.datehired, dataObject.dateofbirth, dataObject.effectivedate, dataObject.reasonforleaving || ''
                ];
                resignationSheet.appendRow(resignedEmployeeData);
            }
        }

        SpreadsheetApp.flush();
        logDataChanges();

        let secondaryChangesMade = false;
        const allDataForSecondary = mainSheet.getDataRange().getValues();
        const reportingToIdIndex = keyMap['reportingtoid'];
        const reportingToNameIndex = keyMap['reportingto'];

        if (isTransfer && reportingToIdIndex !== undefined) {
            for (let i = 1; i < allDataForSecondary.length; i++) {
                if ((String(allDataForSecondary[i][reportingToIdIndex]) || '').toUpperCase() === transferredEmployeeId) {
                    mainSheet.getRange(i + 1, reportingToIdIndex + 1).setValue(oldPositionIdForTransfer);
                    mainSheet.getRange(i + 1, reportingToNameIndex + 1).clearContent();
                    secondaryChangesMade = true;
                }
            }
        }
        if (isManualVacate && reportingToIdIndex !== undefined) {
            for (let i = 1; i < allDataForSecondary.length; i++) {
                if ((String(allDataForSecondary[i][reportingToIdIndex]) || '').toUpperCase() === vacatingEmployeeId) {
                    mainSheet.getRange(i + 1, reportingToIdIndex + 1).setValue(vacatedPositionId);
                    mainSheet.getRange(i + 1, reportingToNameIndex + 1).clearContent();
                    secondaryChangesMade = true;
                }
            }
        }
        if (isFillingVacancy && reportingToIdIndex !== undefined) {
            for (let i = 1; i < allDataForSecondary.length; i++) {
                if (allDataForSecondary[i][reportingToIdIndex] === filledVacancyPositionId) {
                    mainSheet.getRange(i + 1, reportingToIdIndex + 1).setValue(newEmployeeIdForVacancy);
                    mainSheet.getRange(i + 1, reportingToNameIndex + 1).setValue(newEmployeeNameForVacancy);
                    secondaryChangesMade = true;
                }
            }
        }

        if (secondaryChangesMade) {
            SpreadsheetApp.flush();
            const finalData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
            scriptProperties.setProperty('lastKnownData', JSON.stringify(finalData));
        }

        return "Data saved successfully.";
    } catch (e) {
        logToSheet('Error in saveEmployeeData: ' + e.message + ' Stack: ' + e.stack);
        throw new Error('Failed to save data. ' + e.message);
    } finally {
        scriptProperties.deleteProperty('scriptChangeLock');
        lock.releaseLock();
    }
}


function deactivatePosition(positionId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheets()[0];
    const positionIdCol = mainSheet.getRange("A:A").getValues();
    const rowIndex = positionIdCol.findIndex(row => row[0] === positionId);
    if (rowIndex === -1) {
      throw new Error(`Position ID ${positionId} not found for deactivation.`);
    }
    // --- REVISED: Changed to uppercase to match data validation rule ---
    mainSheet.getRange(rowIndex + 1, 18).setValue('INACTIVE');
    SpreadsheetApp.flush();
    logDataChanges();


    return "Position deactivated successfully.";
  } catch (e) {
    logToSheet('Error in deactivatePosition: ' + e.message + ' Stack: ' + e.stack);
    throw new Error('Failed to deactivate position. ' + e.message);
  } finally {
    lock.releaseLock();
  }
}


// PASTE THIS ENTIRE CORRECTED CODE BLOCK

/**
 * HELPER FUNCTION to safely parse dates from the spreadsheet.
 * THIS WAS THE MISSING PIECE.
 * @param {any} dateValue - The value from the spreadsheet cell.
 * @returns {Date|null} A valid Date object or null.
 */
function _parseDate(dateValue) {
  if (!dateValue) return null;
  if (dateValue instanceof Date && !isNaN(dateValue)) return dateValue;
  const parsedDate = new Date(dateValue);
  return !isNaN(parsedDate) ? parsedDate : null;
}

/**
 * REVISED - Generates the Incumbency History report sheet.
 */
function generateIncumbencyReport() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheets()[0];
  const mainData = mainSheet.getLastRow() > 1 ? mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 3).getValues() : [];
  const mainDataMap = new Map(mainData.map(row => [row[0], row]));

  const logSheet = spreadsheet.getSheetByName('change_log_sheet');
  const reportSheetName = 'Incumbency History';
  let reportSheet = spreadsheet.getSheetByName(reportSheetName);

  if (!logSheet || logSheet.getLastRow() < 2) {
    ui.alert('The "change_log_sheet" has no data to report.');
    return;
  }

  const allLogData = logSheet.getDataRange().getValues();
  const headers = allLogData.shift();
  const allHistory = calculateIncumbencyEngine(allLogData, headers, mainDataMap);
  const finalHistoryRecords = [];
  const sortedPosIds = Object.keys(allHistory).sort();

  const allLogDataNoHeaders = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();

  for (const posId of sortedPosIds) {
    const records = allHistory[posId];
    if (!records || records.length === 0) continue;

    const lastRecord = records[records.length - 1];
    const currentLiveRow = mainDataMap.get(posId);
    const currentLiveEmployeeId = currentLiveRow ? (currentLiveRow[1] || '').toString().trim() : null;
    const isCurrentlyVacant = !currentLiveEmployeeId;

    // Correction 1: If the last incumbent is the current employee, ensure end date is 'Present'.
    if (lastRecord.incumbentId && currentLiveEmployeeId && lastRecord.incumbentId === currentLiveEmployeeId && lastRecord.endDate) {
      lastRecord.endDate = null;
    }

    // Correction 2 (Failsafe): If history says 'Present' but the position is vacant, find the true end date.
    if (lastRecord.endDate === null && isCurrentlyVacant) {
      const posIdIndex = headers.indexOf('Position ID');
      const effectiveDateIndex = headers.indexOf('Effective Date');
      const timestampIndex = headers.indexOf('Change Timestamp');

      const allEventsForPos = allLogDataNoHeaders
        .filter(row => row[posIdIndex] === posId)
        .map(row => ({ date: _parseDate(row[effectiveDateIndex]) || _parseDate(row[timestampIndex]) }))
        .filter(event => event.date)
        .sort((a, b) => b.date.getTime() - a.date.getTime());

      if (allEventsForPos.length > 0) {
        lastRecord.endDate = allEventsForPos[0].date;
      }
    }

    records.forEach(rec => {
      const tenure = (rec.startDate && (rec.endDate || new Date())) ? Math.round(((rec.endDate || new Date()) - rec.startDate) / (1000 * 60 * 60 * 24)) : 0;
      finalHistoryRecords.push([
        posId,
        rec.jobTitle,
        rec.incumbentName,
        rec.startDate,
        rec.endDate,
        tenure >= 0 ? tenure : 0,
        rec.changeCount
      ]);
    });
  }

  if (finalHistoryRecords.length === 0) {
    ui.alert('No incumbency history could be generated.');
    return;
  }

  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet(reportSheetName);
  }
  reportSheet.clear();
  const reportHeaders = ['Position ID', 'Job Title', 'Incumbent Name', 'Start Date', 'End Date', 'Tenure (Days)', 'Position Change Count'];
  reportSheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]).setFontWeight('bold');

  if (finalHistoryRecords.length > 0) {
    reportSheet.getRange(2, 1, finalHistoryRecords.length, finalHistoryRecords[0].length).setValues(finalHistoryRecords);
  }

  reportSheet.getRange("D:E").setNumberFormat("yyyy-mm-dd");
  reportSheet.setFrozenRows(1);
  reportSheet.autoResizeColumns(1, reportHeaders.length);
  ui.alert(`Success! "${reportSheetName}" sheet has been updated.`);
}

/**
 * =================================================================================================
 * FINAL REWRITE v9 - calculateIncumbencyEngine
 * =================================================================================================
 * This version correctly identifies the end of a tenure by recognizing "effective-dated" events
 * (like promotions/resignations) as definitive termination points. It also correctly handles
 * subsequent "ghost" log entries that might occur for an employee after their tenure has
 * officially ended, preventing these from creating incorrect history records.
 * =================================================================================================
 */
function calculateIncumbencyEngine(allLogData, headers, mainDataMap) {
  const posIdIndex = headers.indexOf('Position ID');
  const empIdIndex = headers.indexOf('Employee ID');
  const nameIndex = headers.indexOf('Employee Name');
  const jobTitleIndex = headers.indexOf('Job Title');
  const timestampIndex = headers.indexOf('Change Timestamp');
  const effectiveDateIndex = headers.indexOf('Effective Date');
  const hireDateIndex = headers.indexOf('Date Hired');

  const isFirstEverEventForEmployee = (employeeId, eventDate, allLogs) => {
    for (const row of allLogs) {
      const logEmpId = (row[empIdIndex] || '').toString().trim();
      if (logEmpId === employeeId) {
        const logEventDate = _parseDate(row[effectiveDateIndex]) || _parseDate(row[timestampIndex]);
        if (logEventDate && logEventDate.getTime() < eventDate.getTime()) {
          return false;
        }
      }
    }
    return true;
  };

  const positions = {};
  allLogData.forEach(row => {
    const posId = row[posIdIndex];
    if (posId) {
      if (!positions[posId]) positions[posId] = [];
      positions[posId].push(row);
    }
  });

  const finalHistory = {};

  for (const posId in positions) {
    const logEntries = positions[posId];
    const allChangeEventsForPos = logEntries
      .filter(row => row[timestampIndex])
      .map(row => ({
        eventDate: _parseDate(row[effectiveDateIndex]) || _parseDate(row[timestampIndex]),
        incumbentId: (row[empIdIndex] || '').toString().trim() || null,
        incumbentName: (row[nameIndex] || '').toString().trim() || 'N/A',
        jobTitle: (row[jobTitleIndex] || '').toString().trim() || 'N/A',
        hireDate: _parseDate(row[hireDateIndex]),
        status: (row[statusIndex] || '').toString().trim().toUpperCase(), // Store the status.
        isEffective: !!_parseDate(row[effectiveDateIndex])
      }))
      .filter(e => e.eventDate)
      .sort((a, b) => a.eventDate.getTime() - b.eventDate.getTime());

    if (allChangeEventsForPos.length === 0) continue;

    let historyRecords = [];
    let i = 0;
    while (i < allChangeEventsForPos.length) {
      const startEvent = allChangeEventsForPos[i];
      if (!startEvent.incumbentId) {
        i++;
        continue;
      }

      let startDate = startEvent.eventDate;
      if (startEvent.hireDate && startEvent.hireDate.getTime() < startEvent.eventDate.getTime()) {
      // *** MODIFIED START DATE LOGIC ***
      if (internalMovementStatus.includes(startEvent.status)) {
        startDate = startEvent.eventDate;
      } else if (startEvent.hireDate && startEvent.hireDate.getTime() < startEvent.eventDate.getTime()) {
        if (isFirstEverEventForEmployee(startEvent.incumbentId, startEvent.eventDate, allLogData)) {
          startDate = startEvent.hireDate;
        }
      }

      let endDate = null;
      let tenureEndingEvent = null;
      let nextEventIndex = i + 1;

      for (let j = i; j < allChangeEventsForPos.length; j++) {
        const currentEvent = allChangeEventsForPos[j];

        if (currentEvent.incumbentId !== startEvent.incumbentId) {
          endDate = currentEvent.eventDate;
          tenureEndingEvent = currentEvent;
          nextEventIndex = j;
          break;
        }

        if (currentEvent.isEffective && currentEvent.incumbentId === startEvent.incumbentId) {
          endDate = currentEvent.eventDate;
          tenureEndingEvent = currentEvent;
          let k = j + 1;
          while (k < allChangeEventsForPos.length && allChangeEventsForPos[k].incumbentId === startEvent.incumbentId) {
            k++;
          }
          nextEventIndex = k;
          break;
        }
      }

      if (!tenureEndingEvent) {
        nextEventIndex = allChangeEventsForPos.length;
      }

      const lastEventOfThisTenure = allChangeEventsForPos[nextEventIndex - 1];

      historyRecords.push({
        startDate: startDate,
        endDate: endDate,
        incumbentId: startEvent.incumbentId,
        incumbentName: lastEventOfThisTenure.incumbentName,
        jobTitle: lastEventOfThisTenure.jobTitle,
        hireDate: startEvent.hireDate
      });

      i = nextEventIndex;
    }

    const changeCount = historyRecords.length;
    historyRecords.forEach(rec => rec.changeCount = changeCount);
    finalHistory[posId] = historyRecords;
  }
  return finalHistory;
}

/**
 * REVISED - Fetches and formats incumbency history for the web app.
 */
function getDetailedIncumbencyHistory(posId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `incumbency_history_${posId}`;
  const cachedHistory = cache.get(cacheKey);

  if (cachedHistory) {
    Logger.log(`Cache HIT for Position ID: ${posId}`);
    return JSON.parse(cachedHistory);
  }

  Logger.log(`Cache MISS for Position ID: ${posId}. Calculating from scratch.`);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = spreadsheet.getSheets()[0];
    const logSheet = spreadsheet.getSheetByName('change_log_sheet');
    if (!logSheet || !mainSheet || logSheet.getLastRow() < 2) return [];

    const mainData = mainSheet.getLastRow() > 1 ? mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 3).getValues() : [];
    const mainDataMap = new Map(mainData.map(row => [row[0], row]));

    const allLogDataWithHeaders = logSheet.getDataRange().getValues();
    const headers = allLogDataWithHeaders.shift();
    const allLogData = allLogDataWithHeaders;

    // Run the main history engine
    const allHistory = calculateIncumbencyEngine(allLogData, headers, mainDataMap);
    let positionHistory = allHistory[posId] || [];

    // --- START: DATA CORRECTION AND FAILSAFE LOGIC ---
    if (positionHistory.length > 0) {
      const lastRecord = positionHistory[positionHistory.length - 1];
      const liveRow = mainDataMap.get(posId);
      const currentLiveEmployeeId = liveRow ? (liveRow[1] || '').toString().trim() : null;
      const isCurrentlyVacant = !currentLiveEmployeeId;

      // Correction 1: If the last incumbent in the history is the current, active employee,
      // ensure their end date is null (i.e., 'Present'), overriding any erroneous log entry.
      if (lastRecord.incumbentId && currentLiveEmployeeId && lastRecord.incumbentId === currentLiveEmployeeId && lastRecord.endDate) {
        Logger.log(`Position ${posId} is currently held by ${currentLiveEmployeeId}, but history shows an end date. Correcting to 'Present'.`);
        lastRecord.endDate = null;
      }

      // Correction 2 (Failsafe): If the position is actually vacant, but the history shows 'Present',
      // find the last known event for that position and use its date as the end date.
      if (lastRecord.endDate === null && isCurrentlyVacant) {
        Logger.log(`Position ${posId} is vacant but history shows "Present". Applying final failsafe.`);
        const posIdIndex = headers.indexOf('Position ID');
        const effectiveDateIndex = headers.indexOf('Effective Date');
        const timestampIndex = headers.indexOf('Change Timestamp');
        
        const allEventsForPos = allLogData
          .filter(row => row[posIdIndex] === posId)
          .map(row => ({ date: _parseDate(row[effectiveDateIndex]) || _parseDate(row[timestampIndex]) }))
          .filter(event => event.date)
          .sort((a, b) => b.date.getTime() - a.date.getTime());

        if (allEventsForPos.length > 0) {
          lastRecord.endDate = allEventsForPos[0].date;
          Logger.log(`Failsafe applied. Corrected End Date to: ${lastRecord.endDate}`);
        }
      }
    }
    // --- END OF DATA CORRECTION AND FAILSAFE LOGIC ---

    const finalHistory = positionHistory
      .filter(rec => rec.incumbentId)
      .map(rec => {
        const startDate = rec.startDate;
        const endDateForCalc = rec.endDate || new Date();

        let tenureDays = 0;
        if (startDate && endDateForCalc) {
          const diffMillis = endDateForCalc.getTime() - startDate.getTime();
          tenureDays = Math.max(0, Math.floor(diffMillis / (1000 * 60 * 60 * 24)));
        }

        let tenureString = "0 days";
        if (tenureDays > 0) {
          const years = Math.floor(tenureDays / 365.25);
          const months = Math.floor((tenureDays % 365.25) / 30.44);
          const days = Math.round((tenureDays % 365.25) % 30.44);

          let parts = [];
          if (years > 0) parts.push(`${years} year${years > 1 ? 's' : ''}`);
          if (months > 0) parts.push(`${months} month${months > 1 ? 's' : ''}`);
          if (days > 0 || (years === 0 && months === 0)) parts.push(`${days} day${days !== 1 ? 's' : ''}`);
          tenureString = parts.join(', ');
        }

        return {
          name: rec.incumbentName,
          startDate: rec.startDate ? Utilities.formatDate(rec.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'N/A',
          endDate: rec.endDate ? Utilities.formatDate(rec.endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'Present',
          tenure: tenureString,
          employeeHireDate: rec.hireDate ? Utilities.formatDate(rec.hireDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'N/A',
        };
      });

    const reversedHistory = finalHistory.reverse();
    cache.put(cacheKey, JSON.stringify(reversedHistory), 21600);
    return reversedHistory;

  } catch (e) {
    Logger.log(`Error in getDetailedIncumbencyHistory: ${e.toString()}\nStack: ${e.stack}`);
    return [];
  }
}


/**
 * REVISED NOTIFICATION FUNCTION
 */
// PASTE THIS ENTIRE CORRECTED FUNCTION
function getUpcomingDues() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheets()[0];
  const logSheet = spreadsheet.getSheetByName('change_log_sheet');

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const upcoming = [];
  const overdue = [];

  if (mainSheet.getLastRow() < 2) {
    return {
      upcoming,
      overdue
    };
  }

  const mainData = mainSheet.getDataRange().getValues();
  const mainHeaders = mainData.shift();
  const mainDataMap = new Map(mainData.map(row => [row[mainHeaders.indexOf('Position ID')], row]));
  const nameIndex = mainHeaders.indexOf('Employee Name');
  const contractTypeIndex = mainHeaders.indexOf('Contract Type');
  const contractEndIndex = mainHeaders.indexOf('Contract End Date');
  const statusIndex = mainHeaders.indexOf('Status');
  const posStatusIndex = mainHeaders.indexOf('Position Status');
  const dateHiredIndex = mainHeaders.indexOf('Date Hired');

  mainDataMap.forEach((row, posId) => {
    const positionStatus = (row[posStatusIndex] || '').toString().trim().toUpperCase();
    if (positionStatus === 'INACTIVE') return;

    const contractType = (row[contractTypeIndex] || '').toString().trim().toUpperCase();
    const endDate = row[contractEndIndex];
    if (contractType === 'JPRO' && endDate instanceof Date) {
      const normalizedEndDate = new Date(endDate.getTime());
      normalizedEndDate.setHours(0, 0, 0, 0);
      const timeDiff = normalizedEndDate.getTime() - today.getTime();
      const days = Math.round(timeDiff / (1000 * 60 * 60 * 24));
      const employeeName = row[nameIndex];

      if (days >= 0 && days <= 30) {
        const message = `${employeeName}'s JPRO contract ends in ${days} day${days !== 1 ? 's' : ''}.`;
        upcoming.push({
          days,
          message
        });
      } else if (days < 0) {
        const daysAgo = Math.abs(days);
        const message = `${employeeName}'s JPRO contract expired ${daysAgo} day${daysAgo !== 1 ? 's' : ''} ago. Please update their status.`;
        overdue.push({
          days: daysAgo,
          message
        });
      }
    }
  });

  if (statusIndex > -1 && dateHiredIndex > -1) {
    mainData.forEach(row => {
        const positionStatus = (row[posStatusIndex] || '').toString().trim().toUpperCase();
        if (positionStatus === 'INACTIVE') return;

        const status = (row[statusIndex] || '').toString().trim().toUpperCase();
        const startDate = row[dateHiredIndex];

        if ((status === 'PROBATIONARY' || status === 'NEW HIRE') && startDate instanceof Date) {
            const employeeName = row[nameIndex];

            // 3-Month Evaluation Logic
            const evaluationDate = new Date(startDate.getTime());
            evaluationDate.setMonth(evaluationDate.getMonth() + 3);
            evaluationDate.setHours(0, 0, 0, 0);

            const evalTimeDiff = evaluationDate.getTime() - today.getTime();
            const evalDays = Math.round(evalTimeDiff / (1000 * 60 * 60 * 24));

            // --- MODIFIED LOGIC HERE ---
            // Only show notification from 30 days before to 30 days after the due date.
            if (evalDays >= 0 && evalDays <= 30) {
                const message = `${employeeName} is due for 3-month evaluation in ${evalDays} day${evalDays !== 1 ? 's' : ''}.`;
                upcoming.push({ days: evalDays, message });
            } else if (evalDays < 0 && evalDays >= -30) { // Condition changed
                const evalDaysAgo = Math.abs(evalDays);
                const message = `${employeeName}'s 3-month evaluation was due ${evalDaysAgo} day${evalDaysAgo !== 1 ? 's' : ''} ago.`;
                overdue.push({ days: evalDaysAgo, message });
            }

            // 6-Month Regularization Logic (remains the same)
            const regularizationDate = new Date(startDate.getTime());
            regularizationDate.setMonth(regularizationDate.getMonth() + 6);
            regularizationDate.setHours(0, 0, 0, 0);

            const regTimeDiff = regularizationDate.getTime() - today.getTime();
            const regDays = Math.round(regTimeDiff / (1000 * 60 * 60 * 24));
            if (regDays >= 0 && regDays <= 30) {
                const message = `${employeeName} is due for regularization in ${regDays} day${regDays !== 1 ? 's' : ''}.`;
                upcoming.push({ days: regDays, message });
            } else if (regDays < 0) {
                const regDaysAgo = Math.abs(regDays);
                const message = `${employeeName}'s regularization was due ${regDaysAgo} day${regDaysAgo !== 1 ? 's' : ''} ago. Please update their status.`;
                overdue.push({ days: regDaysAgo, message });
            }
        }
    });
  }

  if (logSheet && logSheet.getLastRow() > 1) {
    const logData = logSheet.getDataRange().getValues();
    const logHeaders = logData.shift();
    const logPosIdIndex = logHeaders.indexOf('Position ID');
    const logNameIndex = logHeaders.indexOf('Employee Name');
    const logStatusIndex = logHeaders.indexOf('Status');
    const logEffectiveDateIndex = logHeaders.indexOf('Effective Date');
    if (logPosIdIndex > -1 && logStatusIndex > -1 && logEffectiveDateIndex > -1) {
      const latestResignations = new Map();
      for (let i = logData.length - 1; i >= 0; i--) {
        const row = logData[i];
        const posId = row[logPosIdIndex];
        const logStatus = (row[logStatusIndex] || '').trim().toUpperCase();
        if (posId && logStatus === 'RESIGNED' && !latestResignations.has(posId)) {
          latestResignations.set(posId, {
            date: row[logEffectiveDateIndex],
            name: row[logNameIndex]
          });
        }
      }

      latestResignations.forEach((resignation, posId) => {
        const currentPosData = mainDataMap.get(posId);
        if (!currentPosData || (currentPosData[statusIndex] || '').toUpperCase() !== 'RESIGNED') {
          return;
        }

        const effectiveDate = resignation.date;
        if (effectiveDate instanceof Date) {
          const normalizedEffectiveDate = new Date(effectiveDate.getTime());
          normalizedEffectiveDate.setHours(0, 0, 0, 0);
          const timeDiff = normalizedEffectiveDate.getTime() - today.getTime();
          const days = Math.round(timeDiff / (1000 * 60 * 60 * 24));

          if (days >= 0 && days <= 30) {
            const message = `${resignation.name}'s resignation is effective in ${days} day${days !== 1 ? 's' : ''}.`;
            upcoming.push({
              days,
              message
            });
          } else if (days < 0) {
            const daysAgo = Math.abs(days);
            const message = `${resignation.name}'s resignation was ${daysAgo} day${daysAgo !== 1 ? 's' : ''} ago. Please update the position to VACANT.`;
            overdue.push({
              days: daysAgo,
              message
            });
          }
        }
      });
    }
  }

  const sortedUpcoming = upcoming.sort((a, b) => a.days - b.days).map(d => d.message);
  const sortedOverdue = overdue.sort((a, b) => a.days - b.days).map(d => d.message);
  return {
    upcoming: sortedUpcoming,
    overdue: sortedOverdue
  };
}


// PASTE THIS ENTIRE CORRECTED FUNCTION
function getResignationData(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheets()[0];
  const resignationSheet = ss.getSheetByName('Resignation Data');
  const emptyResult = { reasonCounts: {}, resignationGenderCounts: {}, resignationContractCounts: {}, resignationDivisionCounts: {}, resignationJobGroupCounts: {}, monthlyTurnover: [], yearlyHiresLeavers: { hires: 0, leavers: 0 }, ytdTurnover: 0, attritionRate: 0, overallHeadcount: 0, filteredResignationsCount: 0, retentionRate: 0 };

  if (!mainSheet || mainSheet.getLastRow() < 2) return emptyResult;

  const mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
  const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  const mainHeaderMap = new Map(mainHeaders.map((h, i) => [h.trim(), i]));
  const dateHiredIndex = mainHeaderMap.get('Date Hired');
  const empIdIndex = mainHeaderMap.get('Employee ID');
  const overallHeadcount = mainData.filter(row => row[empIdIndex]).length;

  // --- MOVED VARIABLE DEFINITIONS HERE TO FIX ERROR ---
  const selectedYear = (filters.year && filters.year !== 'All Years') ? parseInt(filters.year) : new Date().getFullYear();
  const monthIndex = (filters.month && !String(filters.month).toLowerCase().startsWith('all')) ? new Date(Date.parse(filters.month +" 1, 2012")).getMonth() : -1;

  if (!resignationSheet || resignationSheet.getLastRow() < 2) {
    return { ...emptyResult, overallHeadcount };
  }

  const resignationData = resignationSheet.getRange(2, 1, resignationSheet.getLastRow() - 1, resignationSheet.getLastColumn()).getValues();
  const resignationHeaders = resignationSheet.getRange(1, 1, 1, resignationSheet.getLastColumn()).getValues()[0];
  const resHeaderMap = new Map(resignationHeaders.map((h, i) => [h.trim(), i]));

  // --- NEW: ATTRITION CALCULATION LOGIC ---
  const logSheet = ss.getSheetByName('change_log_sheet');
  let leaversForAttrition = 0;
  if (logSheet && logSheet.getLastRow() > 1) {
      const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();
      const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
      const logHeaderMap = new Map(logHeaders.map((h, i) => [h.trim(), i]));

      const posIdIndex = logHeaderMap.get('Position ID');
      const posStatusIndex = logHeaderMap.get('Position Status');
      const timestampIndex = logHeaderMap.get('Change Timestamp');

      const deactivatedPositionIds = new Set();
      logData.forEach(row => {
          const eventDate = new Date(row[timestampIndex]);
          const positionStatus = (row[posStatusIndex] || '').toString().trim().toUpperCase();

          if (positionStatus === 'INACTIVE' && eventDate.getFullYear() === selectedYear) {
              if (filters.division && !String(filters.division).toLowerCase().startsWith('all') && row[logHeaderMap.get('Division')] !== filters.division) return;
              if (filters.group && !String(filters.group).toLowerCase().startsWith('all') && row[logHeaderMap.get('Group')] !== filters.group) return;
              if (filters.department && !String(filters.department).toLowerCase().startsWith('all') && row[logHeaderMap.get('Department')] !== filters.department) return;
              if (filters.section && !String(filters.section).toLowerCase().startsWith('all') && row[logHeaderMap.get('Section')] !== filters.section) return;
              if (filters.jobLevel && !String(filters.jobLevel).toLowerCase().startsWith('all') && (row[logHeaderMap.get('Job Level')] || '').toString().trim().toLowerCase() !== filters.jobLevel.toLowerCase()) return;
              if (filters.gender && !String(filters.gender).toLowerCase().startsWith('all') && (row[logHeaderMap.get('Gender')] || '').toString().trim().toLowerCase() !== filters.gender.toLowerCase()) return;

              deactivatedPositionIds.add(row[posIdIndex]);
          }
      });
      leaversForAttrition = deactivatedPositionIds.size;
  }
  // --- END OF NEW LOGIC ---

  const jobGroupMapping = { 1: 'Executives', 2: 'Director', 3: 'Managerial', 4: 'Supervisory', 5: 'Rank & File', 6: 'Jobcon' };

  const filteredResignations = resignationData.filter(row => {
    const resDate = new Date(row[resHeaderMap.get('Resignation Date')]);
    if (filters.year && !String(filters.year).toLowerCase().startsWith('all') && resDate.getFullYear() != filters.year) return false;
    if (monthIndex > -1 && resDate.getMonth() !== monthIndex) return false;
    if (filters.division && !String(filters.division).toLowerCase().startsWith('all') && row[resHeaderMap.get('Division')] !== filters.division) return false;
    if (filters.group && !String(filters.group).toLowerCase().startsWith('all') && row[resHeaderMap.get('Group')] !== filters.group) return false;
    if (filters.department && !String(filters.department).toLowerCase().startsWith('all') && row[resHeaderMap.get('Department')] !== filters.department) return false;
    if (filters.section && !String(filters.section).toLowerCase().startsWith('all') && row[resHeaderMap.get('Section')] !== filters.section) return false;
    if (filters.jobTitle && !String(filters.jobTitle).toLowerCase().startsWith('all') && row[resHeaderMap.get('Job Title')] !== filters.jobTitle) return false;
    if (filters.jobLevel && !String(filters.jobLevel).toLowerCase().startsWith('all') && (row[resHeaderMap.get('Job Level')] || '').toString().trim().toLowerCase() !== filters.jobLevel.toLowerCase()) return false;
    if (filters.gender && !String(filters.gender).toLowerCase().startsWith('all') && (row[resHeaderMap.get('Gender')] || '').toString().trim().toLowerCase() !== filters.gender.toLowerCase()) return false;
    return true;
  });

  const reasonCounts = {}, resignationGenderCounts = {}, resignationContractCounts = {}, resignationDivisionCounts = {}, resignationJobGroupCounts = {};
  filteredResignations.forEach(row => {
    reasonCounts[row[resHeaderMap.get('Reason for Leaving')] || 'Unknown'] = (reasonCounts[row[resHeaderMap.get('Reason for Leaving')] || 'Unknown'] || 0) + 1;
    resignationGenderCounts[row[resHeaderMap.get('Gender')] || 'Unknown'] = (resignationGenderCounts[row[resHeaderMap.get('Gender')] || 'Unknown'] || 0) + 1;
    resignationContractCounts[row[resHeaderMap.get('Contract Type')] || 'Unknown'] = (resignationContractCounts[row[resHeaderMap.get('Contract Type')] || 'Unknown'] || 0) + 1;
    resignationDivisionCounts[row[resHeaderMap.get('Division')] || 'Unknown'] = (resignationDivisionCounts[row[resHeaderMap.get('Division')] || 'Unknown'] || 0) + 1;
    resignationJobGroupCounts[jobGroupMapping[row[resHeaderMap.get('Level')]] || 'Unknown'] = (resignationJobGroupCounts[jobGroupMapping[row[resHeaderMap.get('Level')]] || 'Unknown'] || 0) + 1;
  });

  const hiresThisYear = mainData.filter(r => {
      if (!r[dateHiredIndex]) return false;
      const hiredDate = new Date(r[dateHiredIndex]);
      if (hiredDate.getFullYear() !== selectedYear) return false;
      if (monthIndex > -1 && hiredDate.getMonth() !== monthIndex) return false;
      if (filters.division && !String(filters.division).toLowerCase().startsWith('all') && r[mainHeaderMap.get('Division')] !== filters.division) return false;
      if (filters.group && !String(filters.group).toLowerCase().startsWith('all') && r[mainHeaderMap.get('Group')] !== filters.group) return false;
      if (filters.department && !String(filters.department).toLowerCase().startsWith('all') && r[mainHeaderMap.get('Department')] !== filters.department) return false;
      if (filters.section && !String(filters.section).toLowerCase().startsWith('all') && r[mainHeaderMap.get('Section')] !== filters.section) return false;
      if (filters.jobLevel && !String(filters.jobLevel).toLowerCase().startsWith('all') && (r[mainHeaderMap.get('Job Level')] || '').toString().trim().toLowerCase() !== filters.jobLevel.toLowerCase()) return false;
      if (filters.gender && !String(filters.gender).toLowerCase().startsWith('all') && (r[mainHeaderMap.get('Gender')] || '').toString().trim().toLowerCase() !== filters.gender.toLowerCase()) return false;
      return true;
  });

  const leaversThisYear = filteredResignations;
  const endOfYearHeadcount = mainData.length;
  const startOfYearHeadcount = endOfYearHeadcount - hiresThisYear.length + leaversThisYear.length;
  const averageHeadcount = (startOfYearHeadcount + endOfYearHeadcount) / 2;
  const ytdTurnover = averageHeadcount > 0 ? (leaversThisYear.length / averageHeadcount) * 100 : 0;
  const retentionRate = startOfYearHeadcount > 0 ? ((startOfYearHeadcount - leaversThisYear.length) / startOfYearHeadcount) * 100 : 0;

  // --- CORRECTED ATTRITION RATE CALCULATION ---
  const attritionRate = averageHeadcount > 0 ? (leaversForAttrition / averageHeadcount) * 100 : 0;

  const monthlyTurnover = Array(12).fill(null).map((_, month) => {
    const monthlySeparations = leaversThisYear.filter(r => new Date(r[resHeaderMap.get('Resignation Date')]).getMonth() === month).length;
    const monthlyRate = averageHeadcount > 0 ? (monthlySeparations / averageHeadcount) * 100 : 0;
    return {
      month: new Date(selectedYear, month, 1).toLocaleString('default', { month: 'short' }),
      separations: monthlySeparations,
      rate: parseFloat(monthlyRate.toFixed(2))
    };
  });

  return {
    reasonCounts,
    resignationGenderCounts,
    resignationContractCounts,
    resignationDivisionCounts,
    resignationJobGroupCounts,
    monthlyTurnover,
    yearlyHiresLeavers: { hires: hiresThisYear.length, leavers: leaversThisYear.length },
    ytdTurnover: parseFloat(ytdTurnover.toFixed(2)),
    attritionRate: parseFloat(attritionRate.toFixed(2)),
    overallHeadcount,
    filteredResignationsCount: filteredResignations.length,
    retentionRate: parseFloat(retentionRate.toFixed(2))
  };
}

function getAnalyticsData(filters) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheets()[0];
  const previousSheet = spreadsheet.getSheetByName('Previous Headcount');
  let headers = [];
  let mainData = [];

  if (mainSheet.getLastRow() >= 1) {
    headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    if (mainSheet.getLastRow() > 1) {
      mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    }
  } else {
    return { statusCounts: {}, contractCounts: {}, trendData: [], totalHeadcount: 0, overallHeadcount: 0, ageGenerationCounts: {} };
  }

  const headerMap = {
      division: headers.indexOf('Division'),
      group: headers.indexOf('Group'),
      department: headers.indexOf('Department'),
      section: headers.indexOf('Section'),
      jobLevel: headers.indexOf('Job Level'),
      positionStatus: headers.indexOf('Position Status'),
      gender: headers.indexOf('Gender'),
      jobTitle: headers.indexOf('Job Title'),
      status: headers.indexOf('Status'),
      contractType: headers.indexOf('Contract Type'),
      empId: headers.indexOf('Employee ID'),
      level: headers.indexOf('Level'),
      dateHired: headers.indexOf('Date Hired'),
      dateOfBirth: headers.indexOf('Date of Birth')
  };

  const overallHeadcount = mainData.filter(row =>
      (row[headerMap.positionStatus] || '').toString().trim().toUpperCase() !== 'INACTIVE' && row[headerMap.empId]
  ).length;

  let filteredData = [];
  const statusCounts = {};
  const contractCounts = {};
  let totalHeadcount = 0;

  if (mainData.length > 0) {
    filteredData = mainData.filter(row => {
        if ((row[headerMap.positionStatus] || '').toString().trim().toUpperCase() === 'INACTIVE') {
            return false;
        }
        return Object.keys(filters).every(key => {
            if (!filters[key] || String(filters[key]).toLowerCase().startsWith('all')) return true;
            const colIndex = headerMap[key];
            return colIndex !== -1 && (row[colIndex] || '').toString().trim().toLowerCase() === filters[key].toLowerCase();
        });
    });
  }

  const genderCounts = {};
  const jobGroupCounts = {};
  const losCounts = { '< 1 Year': 0, '1-3 Years': 0, '3-5 Years': 0, '5-10 Years': 0, '10+ Years': 0 };
  const ageGenerationCounts = { 'Gen Z': 0, 'Millennials': 0, 'Gen X': 0, 'Baby Boomers': 0, 'Unknown': 0 };

  const jobGroupMapping = { 1: 'Executives', 2: 'Director', 3: 'Managerial', 4: 'Supervisory', 5: 'Rank & File', 6: 'Jobcon' };
  const today = new Date();

  filteredData.forEach(row => {
    if (row[headerMap.empId]) {
      totalHeadcount++;

      const status = (row[headerMap.status] || 'Unknown').toString().trim();
      statusCounts[status] = (statusCounts[status] || 0) + 1;

      const contract = (row[headerMap.contractType] || 'Unknown').toString().trim();
      contractCounts[contract] = (contractCounts[contract] || 0) + 1;

      const rawGender = (row[headerMap.gender] || 'Unknown').toString().trim().toLowerCase();
      let gender;
      if (rawGender === 'male') gender = 'Male';
      else if (rawGender === 'female') gender = 'Female';
      else gender = 'Unknown';
      genderCounts[gender] = (genderCounts[gender] || 0) + 1;

      const level = row[headerMap.level];
      const jobGroup = jobGroupMapping[level] || 'Unknown';
      jobGroupCounts[jobGroup] = (jobGroupCounts[jobGroup] || 0) + 1;

      const hiredDate = row[headerMap.dateHired] ? new Date(row[headerMap.dateHired]) : null;
      if (hiredDate && !isNaN(hiredDate)) {
        const years = (today.getFullYear() - hiredDate.getFullYear());
        if (years < 1) losCounts['< 1 Year']++;
        else if (years < 3) losCounts['1-3 Years']++;
        else if (years < 5) losCounts['3-5 Years']++;
        else if (years < 10) losCounts['5-10 Years']++;
        else losCounts['10+ Years']++;
      }

      const dob = row[headerMap.dateOfBirth] ? new Date(row[headerMap.dateOfBirth]) : null;
      if (dob && !isNaN(dob)) {
          const birthYear = dob.getFullYear();
          if (birthYear >= 1997) ageGenerationCounts['Gen Z']++;
          else if (birthYear >= 1981) ageGenerationCounts['Millennials']++;
          else if (birthYear >= 1965) ageGenerationCounts['Gen X']++;
          else if (birthYear >= 1946) ageGenerationCounts['Baby Boomers']++;
          else ageGenerationCounts['Unknown']++;
      } else {
          ageGenerationCounts['Unknown']++;
      }
    }
  });

  const trendData = [];
  if (previousSheet && previousSheet.getLastRow() > 1) {
    const prevData = previousSheet.getDataRange().getValues();
    const prevHeaders = prevData.shift();
    const divIdx = 0, grpIdx = 1, dptIdx = 2, secIdx = 3;
    for (let i = 0; i < prevHeaders.length; i++) {
      const header = prevHeaders[i];
      if (header.includes(' Filled')) {
        const month = header.replace(' Filled', '').trim();
        const vacantHeader = `${month} Vacant`;
        const vacantIndex = prevHeaders.indexOf(vacantHeader);
        if (vacantIndex !== -1) {
          let totalFilled = 0;
          let totalVacant = 0;
          const sectionFilter = (filters.section && !String(filters.section).toLowerCase().startsWith('all')) ? filters.section : null;
          const departmentFilter = (filters.department && !String(filters.department).toLowerCase().startsWith('all')) ? filters.department : null;
          const groupFilter = (filters.group && !String(filters.group).toLowerCase().startsWith('all')) ? filters.group : null;
          const divisionFilter = (filters.division && !String(filters.division).toLowerCase().startsWith('all')) ? filters.division : null;
          let targetRow = null;
          if (sectionFilter) targetRow = prevData.find(r => r[secIdx] === sectionFilter && r[dptIdx] === departmentFilter && r[grpIdx] === groupFilter && r[divIdx] === divisionFilter);
          else if (departmentFilter) targetRow = prevData.find(r => r[dptIdx] === departmentFilter && r[grpIdx] === groupFilter && r[divIdx] === divisionFilter && !r[secIdx]);
          else if (groupFilter) targetRow = prevData.find(r => r[grpIdx] === groupFilter && r[divIdx] === divisionFilter && !r[dptIdx] && !r[secIdx]);
          else if (divisionFilter) targetRow = prevData.find(r => r[divIdx] === divisionFilter && !r[grpIdx] && !r[dptIdx] && !r[secIdx]);
          if (targetRow) {
              totalFilled = parseInt(targetRow[i] || 0);
              totalVacant = parseInt(targetRow[vacantIndex] || 0);
          } else if (!sectionFilter && !departmentFilter && !groupFilter && !divisionFilter) {
              const divisionTotalRows = prevData.filter(r => r[divIdx] && !r[grpIdx] && !r[dptIdx] && !r[secIdx]);
              divisionTotalRows.forEach(row => {
                  totalFilled += parseInt(row[i] || 0);
                  totalVacant += parseInt(row[vacantIndex] || 0);
              });
          }
          trendData.push({ month: month, filled: totalFilled, vacant: totalVacant });
        }
      }
    }
  }

  const newHiresByMonth = {};
  const logSheet = spreadsheet.getSheetByName('change_log_sheet');
  if (logSheet && logSheet.getLastRow() > 1) {
    const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();
    const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    const logHeaderMap = new Map(logHeaders.map((h, i) => [h.trim(), i]));
    const logStatusIndex = logHeaderMap.get('Status');
    const logEffectiveDateIndex = logHeaderMap.get('Effective Date');
    const logTimestampIndex = logHeaderMap.get('Change Timestamp');
    const twelveMonthsAgo = new Date();
    twelveMonthsAgo.setMonth(twelveMonthsAgo.getMonth() - 12);
    logData.forEach(row => {
      if ((row[logStatusIndex] || '').toUpperCase() === 'NEW HIRE') {
        const eventDate = row[logEffectiveDateIndex] || row[logTimestampIndex];
        if (eventDate && new Date(eventDate) >= twelveMonthsAgo) {
          const date = new Date(eventDate);
          const monthYear = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM yyyy');
          newHiresByMonth[monthYear] = (newHiresByMonth[monthYear] || 0) + 1;
        }
       }
    });
  }

  return {
    statusCounts: statusCounts,
    contractCounts: contractCounts,
    genderCounts: genderCounts,
    jobGroupCounts: jobGroupCounts,
    losCounts: losCounts,
    ageGenerationCounts: ageGenerationCounts,
    trendData: trendData,
    totalHeadcount: totalHeadcount,
    filteredPositionsCount: filteredData.length,
    overallHeadcount: overallHeadcount,
    newHiresByMonth: newHiresByMonth
  };
}

function getEmployeeMovementData(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheets()[0];
  const logSheet = ss.getSheetByName('change_log_sheet');

  const emptyResult = {
    promotionCount: 0,
    transferCount: 0,
    promotionRate: 0,
    transferRate: 0,
    promotionsByDept: {},
    transfersByDept: {}
  };

  if (!logSheet) {
    Logger.log('Warning: change_log_sheet not found. Returning empty data for Employee Movement.');
    return emptyResult;
  }

  if (!mainSheet || mainSheet.getLastRow() < 2 || logSheet.getLastRow() < 2) {
    return emptyResult;
  }

  // Get main data for headcount calculation
  const mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
  const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  const mainHeaderMap = new Map(mainHeaders.map((h, i) => [h.trim(), i]));

  // Get log data for movement calculation
  const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();
  const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  const logHeaderMap = new Map(logHeaders.map((h, i) => [h.trim(), i]));

  const statusIndex = logHeaderMap.get('Status');
  const effectiveDateIndex = logHeaderMap.get('Effective Date');
  const timestampIndex = logHeaderMap.get('Change Timestamp');
  const departmentIndex = logHeaderMap.get('Department');

  const selectedYear = (filters.year && !String(filters.year).toLowerCase().startsWith('all')) ? parseInt(filters.year) : new Date().getFullYear();

  // Filter log data based on filters
  const filteredLogData = logData.filter(row => {
    const eventDate = row[effectiveDateIndex] || row[timestampIndex];
    if (!eventDate) return false;

    const date = new Date(eventDate);
    if (filters.year && !String(filters.year).toLowerCase().startsWith('all') && date.getFullYear() != filters.year) return false;
    if (filters.month && !String(filters.month).toLowerCase().startsWith('all')) {
        const monthIndex = new Date(Date.parse(filters.month +" 1, 2012")).getMonth();
        if (date.getMonth() != monthIndex) return false;
    }
    
    // Apply other location filters
    if (filters.division && !String(filters.division).toLowerCase().startsWith('all') && row[logHeaderMap.get('Division')] !== filters.division) return false;
    if (filters.group && !String(filters.group).toLowerCase().startsWith('all') && row[logHeaderMap.get('Group')] !== filters.group) return false;
    if (filters.department && !String(filters.department).toLowerCase().startsWith('all') && row[logHeaderMap.get('Department')] !== filters.department) return false;
    if (filters.section && !String(filters.section).toLowerCase().startsWith('all') && row[logHeaderMap.get('Section')] !== filters.section) return false;
    if (filters.jobLevel && !String(filters.jobLevel).toLowerCase().startsWith('all') && row[logHeaderMap.get('Job Level')] !== filters.jobLevel) return false;
    if (filters.gender && !String(filters.gender).toLowerCase().startsWith('all') && row[logHeaderMap.get('Gender')] !== filters.gender) return false;
    if (filters.jobTitle && !String(filters.jobTitle).toLowerCase().startsWith('all') && row[logHeaderMap.get('Job Title')] !== filters.jobTitle) return false;
    
    return true;
  });

  let promotionCount = 0;
  let transferCount = 0;
  const promotionsByDept = {};
  const transfersByDept = {};

  filteredLogData.forEach(row => {
    const status = (row[statusIndex] || '').toUpperCase();
    const department = row[departmentIndex] || 'Unknown';

    if (status === 'PROMOTION') {
      promotionCount++;
      promotionsByDept[department] = (promotionsByDept[department] || 0) + 1;
    } else if (status === 'INTERNAL TRANSFER' || status === 'LATERAL TRANSFER') {
      transferCount++;
      transfersByDept[department] = (transfersByDept[department] || 0) + 1;
    }
  });

  // Calculate Rates (similar to resignation rate)
  const leaversThisYear = logData.filter(r => (r[logHeaderMap.get('Status')] || '').toUpperCase() === 'RESIGNED' && new Date(r[effectiveDateIndex] || r[timestampIndex]).getFullYear() === selectedYear);
  const hiresThisYear = mainData.filter(r => r[mainHeaderMap.get('Date Hired')] && new Date(r[mainHeaderMap.get('Date Hired')]).getFullYear() === selectedYear);

  const endOfYearHeadcount = mainData.filter(r => r[mainHeaderMap.get('Employee ID')]).length;
  const startOfYearHeadcount = endOfYearHeadcount - hiresThisYear.length + leaversThisYear.length;
  const averageHeadcount = (startOfYearHeadcount + endOfYearHeadcount) / 2;

  const promotionRate = averageHeadcount > 0 ? (promotionCount / averageHeadcount) * 100 : 0;
  const transferRate = averageHeadcount > 0 ? (transferCount / averageHeadcount) * 100 : 0;

  return {
    promotionCount,
    transferCount,
    promotionRate: parseFloat(promotionRate.toFixed(2)),
    transferRate: parseFloat(transferRate.toFixed(2)),
    promotionsByDept,
    transfersByDept
  };
}


function generateMasterlistSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = spreadsheet.getSheets()[0];

    if (mainSheet.getLastRow() < 2) {
      throw new Error("The source data sheet is empty.");
    }

    const data = mainSheet.getDataRange().getValues();
    
    const newSpreadsheet = SpreadsheetApp.create(`Employee Masterlist - ${new Date().toLocaleDateString()}`);
    const newSheet = newSpreadsheet.getSheets()[0];
    
    newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    newSheet.setFrozenRows(1);
    newSheet.autoResizeColumns(1, data[0].length);
    
    return newSpreadsheet.getUrl();

  } catch (e) {
    Logger.log(`Error in generateMasterlistSheet: ${e.toString()}`);
    throw new Error(`Failed to generate masterlist. Error: ${e.message}`);
  }
}

function generateMasterlistSheetWithPrompt() {
  const ui = SpreadsheetApp.getUi();
  try {
    const url = generateMasterlistSheet();
    const htmlOutput = HtmlService.createHtmlOutput(`<p>Masterlist generated successfully. <a href="${url}" target="_blank">Click here to open the new spreadsheet.</a></p>`)
        .setWidth(400)
        .setHeight(100);
    ui.showModalDialog(htmlOutput, 'Masterlist Generated');
  } catch (e) {
    ui.alert('Error', `Failed to generate masterlist: ${e.message}`, ui.ButtonSet.OK);
  }
}

function getJdFileUrl(positionId, employeeId, jobTitle, type) {
  try {
    let folder;
    let files;
    
    if (type === 'general') {
      folder = DriveApp.getFolderById(JD_GENERAL_FOLDER_ID);
      const searchPrefix = positionId;
      files = folder.getFiles();
      // Loop through all files in the folder to find a match
      while (files.hasNext()) {
        const file = files.next();
        // If a file starts with the Position ID, we've found it
        if (file.getName().startsWith(searchPrefix)) {
          // Use the correct embedding URL format
          return `https://drive.google.com/file/d/${file.getId()}/preview`;
        }
      }
    } else if (type === 'incumbent') {
      if (!employeeId) {
        return null; // Still require an employeeId for this type
      }
      folder = DriveApp.getFolderById(JD_INCUMBENT_FOLDER_ID);
      // The incumbent filename is more specific, so we can do a more targeted search
      const fileName = `${positionId}-${jobTitle}-${employeeId}.pdf`;
      files = folder.getFilesByName(fileName);
      if (files.hasNext()) {
        const file = files.next();
        // Use the correct embedding URL format
        return `https://drive.google.com/file/d/${file.getId()}/preview`;
      }
    } else {
      throw new Error("Invalid job description type specified.");
    }

    return null; // Return null if no file was found after searching

  } catch (e) {
    Logger.log(`Error in getJdFileUrl: ${e.toString()}`);
    return `Error: ${e.message}`;
  }
}

function getStartDateForLastMovement(positionId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('change_log_sheet');
    if (!logSheet || logSheet.getLastRow() < 2) {
      return null;
    }

    const logData = logSheet.getDataRange().getValues();
    const headers = logData.shift();
    const posIdIndex = headers.indexOf('Position ID');
    const statusIndex = headers.indexOf('Status');
    const effectiveDateIndex = headers.indexOf('Effective Date');
    const timestampIndex = headers.indexOf('Change Timestamp');

    if ([posIdIndex, statusIndex, effectiveDateIndex, timestampIndex].includes(-1)) {
      return null; // A required column is missing
    }

    const movementStatuses = ['PROMOTION', 'INTERNAL TRANSFER', 'LATERAL TRANSFER', 'FILLED VACANCY', 'NEW HIRE', 'PROBATIONARY'];

    const relevantEvents = logData
      .filter(row => row[posIdIndex] === positionId && movementStatuses.includes((row[statusIndex] || '').toUpperCase()))
      .map(row => new Date(row[effectiveDateIndex] || row[timestampIndex]))
      .filter(date => !isNaN(date.getTime()));

    if (relevantEvents.length === 0) {
      return null;
    }

    // Sort descending to get the most recent date first
    relevantEvents.sort((a, b) => b.getTime() - a.getTime());

    return Utilities.formatDate(relevantEvents[0], Session.getScriptTimeZone(), 'yyyy-MM-dd');

  } catch (e) {
    Logger.log(`Error in getStartDateForLastMovement: ${e.toString()}`);
    return null;
  }
}

function getDateHired(employeeId) {
  if (!employeeId) {
    return null;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheets()[0];
    if (mainSheet.getLastRow() < 2) {
      return null;
    }
    const data = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    const empIdIndex = headers.indexOf('Employee ID');
    const dateHiredIndex = headers.indexOf('Date Hired');

    if (empIdIndex === -1 || dateHiredIndex === -1) {
      return null;
    }

    const employeeRow = data.find(row => (row[empIdIndex] || '').toString().trim() === employeeId.trim());

    if (employeeRow && employeeRow[dateHiredIndex] instanceof Date) {
      return Utilities.formatDate(employeeRow[dateHiredIndex], Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    return null;
  } catch (e) {
    Logger.log(`Error in getDateHired: ${e.toString()}`);
    return null;
  }
}

function reactivatePosition(positionId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheets()[0];
    const positionIdCol = mainSheet.getRange("A:A").getValues();
    const rowIndex = positionIdCol.findIndex(row => row[0] === positionId);

    if (rowIndex === -1) {
      throw new Error(`Position ID ${positionId} not found for reactivation.`);
    }
    // --- REVISED: Changed to uppercase to match data validation rule ---
    mainSheet.getRange(rowIndex + 1, 18).setValue('ACTIVE'); 
    SpreadsheetApp.flush();
    logDataChanges(); // Log this change to keep history accurate

    return "Position reactivated successfully.";
  } catch (e) {
    Logger.log('Error in reactivatePosition: ' + e.message + ' Stack: ' + e.stack);
    throw new Error('Failed to reactivate position. ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function getMasterlistData(filters) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = spreadsheet.getSheets()[0];
    
    if (mainSheet.getLastRow() < 1) {
      return { headers: ["Info"], rows: [["The main data sheet is empty."]] };
    }

    const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    
    // If filters is null, it's a request for just the headers
    if (filters === null) {
        return { headers: headers, rows: [] };
    }

    if (mainSheet.getLastRow() < 2) {
        return { headers: headers, rows: [] };
    }
    
    const allData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    const headerMap = new Map(headers.map((h, i) => [h.trim(), i]));

    const filteredRows = allData.filter(row => {
      const positionStatusIndex = headerMap.get('Position Status');
      if ((row[positionStatusIndex] || '').toString().trim().toUpperCase() === 'INACTIVE') {
        return false;
      }
      
      return Object.keys(filters).every(key => {
        if (!filters[key] || String(filters[key]).toLowerCase().startsWith('all')) {
          return true;
        }
        const colIndex = headerMap.get(key);
        if (colIndex === undefined) return true; 
        return (row[colIndex] || '').toString().trim().toLowerCase() === filters[key].toLowerCase();
      });
    });

    const serializableRows = filteredRows.map(row => {
        return row.map(cell => {
            if (cell instanceof Date) {
                return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            }
            return cell;
        });
    });

    return { headers: headers, rows: serializableRows };

  } catch (e) {
    Logger.log('FATAL Error in getMasterlistData: ' + e.message + ' Stack: ' + e.stack);
    return { headers: ["Fatal Error"], rows: [[`A critical error occurred on the server: ${e.message}`]] };
  }
}

function generateFilteredMasterlist(payload) {
  try {
    const filters = payload.filters;
    const visibleColumns = payload.visibleColumns;
    
    const masterlistData = getMasterlistData(filters);
    
    if (masterlistData.rows.length === 0) {
      throw new Error("No data matches the current filters to export.");
    }
    
    const originalHeaders = masterlistData.headers;
    const headerIndexMap = originalHeaders.map((header, index) => visibleColumns.includes(header) ? index : -1).filter(index => index !== -1);
    
    const finalHeaders = headerIndexMap.map(index => originalHeaders[index]);
    const finalRows = masterlistData.rows.map(row => headerIndexMap.map(index => row[index]));

    const newSpreadsheet = SpreadsheetApp.create(`Filtered Masterlist - ${new Date().toLocaleDateString()}`);
    const newSheet = newSpreadsheet.getSheets()[0];
    
    newSheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
    
    if (finalRows.length > 0) {
        newSheet.getRange(2, 1, finalRows.length, finalRows[0].length).setValues(finalRows);
    }
    
    newSheet.setFrozenRows(1);
    newSheet.autoResizeColumns(1, finalHeaders.length);
    
    // --- THIS IS THE FIX ---
    // Get the file and move it to the specified folder
    const file = DriveApp.getFileById(newSpreadsheet.getId());
    const folder = DriveApp.getFolderById(MASTERLIST_EXPORT_FOLDER_ID);
    file.moveTo(folder);
    
    return newSpreadsheet.getUrl();

  } catch (e) {
    Logger.log(`Error in generateFilteredMasterlist: ${e.toString()}`);
    throw new Error(`Failed to generate masterlist. Error: ${e.message}`);
  }
}

// PASTE THIS NEW FUNCTION AT THE END OF THE FILE
function getDateOfBirth(employeeId) {
  if (!employeeId) {
    return null;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheets()[0];
    if (mainSheet.getLastRow() < 2) {
      return null;
    }
    const data = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    const empIdIndex = headers.indexOf('Employee ID');
    const dobIndex = headers.indexOf('Date of Birth');

    if (empIdIndex === -1 || dobIndex === -1) {
      return null; // A required column is missing
    }

    // Find the first occurrence of this employee ID in the sheet
    const employeeRow = data.find(row => (row[empIdIndex] || '').toString().trim() === employeeId.trim());

    if (employeeRow && employeeRow[dobIndex] instanceof Date) {
      return Utilities.formatDate(employeeRow[dobIndex], Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    return null;
  } catch (e) {
    Logger.log(`Error in getDateOfBirth: ${e.toString()}`);
    return null;
  }
}

// --- START: NEW PREDICTIVE INSIGHTS FUNCTION ---

// --- START: REVISED PREDICTIVE INSIGHTS FUNCTION (v2) ---

function getAttritionRiskData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheets()[0];
    const resignationSheet = ss.getSheetByName('Resignation Data');
    const logSheet = ss.getSheetByName('change_log_sheet');

    // 1. Get Current Employee & Log Data
    const mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    const headerMap = new Map(headers.map((h, i) => [h.trim(), i]));
    const currentEmployees = mainData.filter(row => row[headerMap.get('Employee ID')]);
    
    // Get historical log data for promotion/transfer analysis
    let logData = [];
    let logHeaderMap = new Map();
    if (logSheet && logSheet.getLastRow() > 1) {
        logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();
        const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
        logHeaderMap = new Map(logHeaders.map((h, i) => [h.trim(), i]));
    }

    // 2. Analyze Historical Resignation Data
    let highTurnoverDepts = [];
    if (resignationSheet && resignationSheet.getLastRow() > 1) {
      const resignationData = resignationSheet.getDataRange().getValues();
      const resHeaders = resignationData.shift();
      const resDeptIndex = resHeaders.indexOf('Department');
      const deptCounts = resignationData.reduce((acc, row) => {
        const dept = row[resDeptIndex];
        if (dept) acc[dept] = (acc[dept] || 0) + 1;
        return acc;
      }, {});
      const turnoverThreshold = 3;
      highTurnoverDepts = Object.keys(deptCounts).filter(dept => deptCounts[dept] >= turnoverThreshold);
    }

    // 3. Score Each Current Employee
    const employeesAtRisk = [];
    const today = new Date();

    currentEmployees.forEach(row => {
      let riskScore = 0;
      let riskFactors = [];

      const empId = row[headerMap.get('Employee ID')];
      const dept = row[headerMap.get('Department')];
      const dateHired = new Date(row[headerMap.get('Date Hired')]);
      const contractType = row[headerMap.get('Contract Type')] || '';

      // Factor 1: Tenure
      if (!isNaN(dateHired)) {
        const tenureMonths = (today.getFullYear() - dateHired.getFullYear()) * 12 + (today.getMonth() - dateHired.getMonth());
        if (tenureMonths < 12) { riskScore += 3; riskFactors.push("Tenure < 1 Year"); } 
        else if (tenureMonths < 24) { riskScore += 2; riskFactors.push("Tenure < 2 Years"); }
      }
      
      // Factor 2: High-Turnover Department
      if (highTurnoverDepts.includes(dept)) { riskScore += 2; riskFactors.push("High-Turnover Dept"); }

      // Factor 3: Contract Type
      if (contractType.toUpperCase() === 'JPRO') { riskScore += 3; riskFactors.push("JPRO Contract"); }

      // Factor 4: Time Since Last Promotion/Movement
      if (empId && logData.length > 0) {
          const movementEvents = logData.filter(logRow => 
              (logRow[logHeaderMap.get('Employee ID')] || '') == empId &&
              ['PROMOTION', 'INTERNAL TRANSFER', 'LATERAL TRANSFER'].includes((logRow[logHeaderMap.get('Status')] || '').toUpperCase())
          ).map(logRow => new Date(logRow[logHeaderMap.get('Effective Date')] || logRow[logHeaderMap.get('Change Timestamp')]))
          .filter(date => !isNaN(date.getTime()));
          
          let lastEventDate = dateHired; // Default to hire date if no other movement
          if (movementEvents.length > 0) {
              lastEventDate = new Date(Math.max.apply(null, movementEvents));
          }

          if (!isNaN(lastEventDate)) {
              const monthsSinceMovement = (today.getFullYear() - lastEventDate.getFullYear()) * 12 + (today.getMonth() - lastEventDate.getMonth());
              if (monthsSinceMovement >= 36) {
                  riskScore += 2;
                  riskFactors.push("No Role Change > 3 Years");
              }
          }
      }

      if (riskScore > 2) { // Only show employees with a moderate to high risk score
        employeesAtRisk.push({
          name: row[headerMap.get('Employee Name')],
          position: row[headerMap.get('Job Title')],
          department: dept,
          score: riskScore,
          factors: riskFactors.join(', ')
        });
      }
    });

    employeesAtRisk.sort((a, b) => b.score - a.score);
    return employeesAtRisk.slice(0, 20);

  } catch (e) {
    Logger.log('Error in getAttritionRiskData: ' + e.message);
    return { error: e.message };
  }
}

/**
 * Fetches the data required to build a team competency heatmap.
 * @param {string} department The department to filter by.
 * @returns {Object} An object containing employee scores and competency headers.
 */
function getTeamCompetencyHeatmap(department) {
  if (!department) return { error: 'Department not specified.' };

  try {
    const hubSs = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = hubSs.getSheets()[0];
    const compSs = SpreadsheetApp.openById(COMPETENCY_SPREADSHEET_ID);
    const compSheet = compSs.getSheetByName('Competency Matrix');

    if (!mainSheet || !compSheet) {
      return { error: 'Required sheets not found.' };
    }

    // Get department data from main sheet
    const mainData = mainSheet.getDataRange().getValues();
    const mainHeaders = mainData.shift();
    const mainEmpIdIndex = mainHeaders.indexOf('Employee ID');
    const mainDeptIndex = mainHeaders.indexOf('Department');
    const employeeToDeptMap = new Map();
    mainData.forEach(row => {
      if (row[mainEmpIdIndex] && row[mainDeptIndex]) {
        employeeToDeptMap.set(String(row[mainEmpIdIndex]).trim(), row[mainDeptIndex]);
      }
    });

    // Get competency data
    const compData = compSheet.getDataRange().getValues();
    const rawCompHeaders = compData.shift();
    const compHeaders = rawCompHeaders.map(h => (h || '').toString().replace(/\n|\r/g, ' ').trim());
    const compEmpIdIndex = compHeaders.indexOf('EMPLOYEE ID');
    const compNameIndex = compHeaders.indexOf('EMPLOYEE NAME');

    // Filter employees by the selected department
    const departmentEmployees = compData.filter(row => {
      const empId = String(row[compEmpIdIndex]).trim();
      return employeeToDeptMap.get(empId) === department;
    });

    if (departmentEmployees.length === 0) {
      return { headers: [], employeeScores: [] }; // No employees in this department
    }

    const competencyColumns = [];
    const competencyHeaderNames = [];

    // --- REVISED: More robust competency column identification ---
    const allDefinedCompetencies = [
        "TRUSTWORTHINESS", "ENTREPRENEURIAL SPIRIT", "INNOVATION", "LEADERSHIP", "RESPECT FOR THE INDIVIDUAL",
        "LEADERSHIP BY EXAMPLE", "DRIVE FOR RESULTS", "COACHING FOR SUCCESS", "INSPIRING LOYAL AND TRUST",
        "WORKING ACROSS TEAMS", "TALENT MANAGEMENT AND DEVELOPMENT", "EMPOWERMENT", "COMMUNICATION", "EXECUTIVE DISPOSITION",
        "PROJECT MANAGEMENT", "LEAN THINKING PRINCIPLES", "PROCESS STANDARDIZATION", "OPERATIONAL EXPERTISE",
        "COST MANAGEMENT", "DATA-DRIVEN DECISION MAKING", "MANAGEMENT OF WORK SYSTEMS/BUSINESS PROCESS ORIENTATION"
    ];

    allDefinedCompetencies.forEach(competencyName => {
        let actualIndex = compHeaders.indexOf(`${competencyName} [Actual]`);

        // Fallback for headers that might just have the competency name
        if (actualIndex === -1) {
            const indices = compHeaders.map((h, i) => (h.trim().toUpperCase() === competencyName ? i : -1)).filter(i => i !== -1);
            if (indices.length >= 2) {
                actualIndex = indices[1]; // Assume the second column is 'Actual'
            }
        }
        
        if (actualIndex !== -1) {
            competencyColumns.push({ name: competencyName, index: actualIndex });
            competencyHeaderNames.push(competencyName);
        }
    });
    // --- END REVISED SECTION ---

    // Structure the data for the frontend
    const employeeScores = departmentEmployees.map(row => {
      const scores = competencyColumns.map(col => {
        return parseFloat(row[col.index]) || 0;
      });
      return {
        name: row[compNameIndex],
        scores: scores
      };
    });

    return {
      headers: competencyHeaderNames,
      employeeScores: employeeScores
    };

  } catch (e) {
    Logger.log('Error in getTeamCompetencyHeatmap: ' + e.message);
    return { error: e.message };
  }
}

/**
 * Fetches data for a deep dive into a specific competency for an employee.
 * @param {string} employeeId The ID of the employee.
 * @param {string} competencyName The name of the competency.
 * @returns {Object} An object containing the deep dive data.
 */
function getCompetencyDeepDive(employeeId, competencyName) {
  if (!employeeId || !competencyName) return { error: 'Employee ID or competency name not specified.' };

  try {
    const hubSs = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = hubSs.getSheets()[0];
    const compSs = SpreadsheetApp.openById(COMPETENCY_SPREADSHEET_ID);
    const compSheet = compSs.getSheetByName('Competency Matrix');

    if (!mainSheet || !compSheet) {
      return { error: 'Required sheets not found.' };
    }

    // Get department data from main sheet
    const mainData = mainSheet.getDataRange().getValues();
    const mainHeaders = mainData.shift();
    const mainEmpIdIndex = mainHeaders.indexOf('Employee ID');
    const mainDeptIndex = mainHeaders.indexOf('Department');
    const employeeToDeptMap = new Map();
    mainData.forEach(row => {
      if (row[mainEmpIdIndex] && row[mainDeptIndex]) {
        employeeToDeptMap.set(String(row[mainEmpIdIndex]).trim(), row[mainDeptIndex]);
      }
    });
    
    // Get competency data
    const compData = compSheet.getDataRange().getValues();
    const rawCompHeaders = compData.shift();
    const compHeaders = rawCompHeaders.map(h => (h || '').toString().replace(/\n|\r/g, ' ').trim());
    const compEmpIdIndex = compHeaders.indexOf('EMPLOYEE ID');
    const compNameIndex = compHeaders.indexOf('EMPLOYEE NAME');
    
    let actualScoreIndex = -1;
    let requiredScoreIndex = -1;

    // Flexible header finding
    const requiredHeader = `${competencyName} [Required]`;
    const actualHeader = `${competencyName} [Actual]`;
    
    actualScoreIndex = compHeaders.indexOf(actualHeader);
    
    // Fallback for headers that might just have the competency name
    if (actualScoreIndex === -1) {
        const indices = compHeaders.map((h, i) => (h.trim() === competencyName ? i : -1)).filter(i => i !== -1);
        if (indices.length >= 2) {
          requiredScoreIndex = indices[0];
          actualScoreIndex = indices[1];
        }
    }

    if (actualScoreIndex === -1) {
        return { error: `Competency '${competencyName}' score column not found.`};
    }
    
    // --- CALCULATIONS ---
    let employeeScore = 0;
    let teamTotalScore = 0;
    let teamMemberCount = 0;
    let companyTotalScore = 0;
    let companyMemberCount = 0;
    const topPerformers = [];

    const employeeDept = employeeToDeptMap.get(String(employeeId).trim());

    compData.forEach(row => {
      const currentEmpId = String(row[compEmpIdIndex]).trim();
      const score = parseFloat(row[actualScoreIndex]) || 0;
      
      if(score > 0) {
        // Company-wide calculation
        companyTotalScore += score;
        companyMemberCount++;

        // Team calculation
        if (employeeDept && employeeToDeptMap.get(currentEmpId) === employeeDept) {
          teamTotalScore += score;
          teamMemberCount++;
        }

        // Employee's specific score
        if (currentEmpId === String(employeeId).trim()) {
          employeeScore = score;
        }

        // Top performers list
        topPerformers.push({ name: row[compNameIndex], score: score });
      }
    });

    const teamAverage = teamMemberCount > 0 ? (teamTotalScore / teamMemberCount) : 0;
    const companyAverage = companyMemberCount > 0 ? (companyTotalScore / companyMemberCount) : 0;

    // Sort and get top 5 performers
    const sortedTopPerformers = topPerformers
      .sort((a, b) => b.score - a.score)
      .slice(0, 5);

    return {
      employeeScore: employeeScore,
      teamAverage: parseFloat(teamAverage.toFixed(2)),
      companyAverage: parseFloat(companyAverage.toFixed(2)),
      topPerformers: sortedTopPerformers
    };

  } catch (e) {
    Logger.log('Error in getCompetencyDeepDive: ' + e.message);
    return { error: e.message };
  }
}

function testBackendConnection() {
  Logger.log('SUCCESS: The backend connection test function was called successfully.');
  return "Connection successful!";
}
// --- END: REVISED PREDICTIVE INSIGHTS FUNCTION ---

// --- START: NEW HIRING PREDICTIONS FUNCTION ---

function getHiringPredictions() {
  try {
    const recruitmentSheetId = "1SfJbXTqtN5Bu7y4ikt0G-3hqiwZLIhezXqxv_zRxRFA"; // Make sure this is correct!
    const spreadsheet = SpreadsheetApp.openById(recruitmentSheetId);
    const sheet = spreadsheet.getSheets()[0];
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { timeToHire: {}, sourceQuality: {} };
    }
    
    const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = new Map(headers.map((h, i) => [h.trim(), i]));

    // 1. Calculate Average Time to Hire per Position
    const timeToHireData = {};
    allData.forEach(row => {
      const position = row[headerMap.get('Position Applied For')];
      const status = (row[headerMap.get('Application Status')] || '').toUpperCase();
      const appDate = new Date(row[headerMap.get('Application Date')]);
      const hireDate = new Date(row[headerMap.get('Hiring Date')]);

      if (position && status === 'HIRED' && !isNaN(appDate) && !isNaN(hireDate)) {
        const diffTime = Math.abs(hireDate - appDate);
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        
        if (!timeToHireData[position]) {
          timeToHireData[position] = { totalDays: 0, count: 0 };
        }
        timeToHireData[position].totalDays += diffDays;
        timeToHireData[position].count++;
      }
    });
    
    const timeToHirePrediction = {};
    for (const pos in timeToHireData) {
      timeToHirePrediction[pos] = Math.round(timeToHireData[pos].totalDays / timeToHireData[pos].count);
    }

    // 2. Calculate Source Quality (Conversion Rate)
    const sourceData = {};
    allData.forEach(row => {
        const source = row[headerMap.get('Source')];
        const status = (row[headerMap.get('Application Status')] || '').toUpperCase();
        if (source) {
            if (!sourceData[source]) {
                sourceData[source] = { applications: 0, hires: 0 };
            }
            sourceData[source].applications++;
            if (status === 'HIRED') {
                sourceData[source].hires++;
            }
        }
    });

    const sourceQualityPrediction = {};
    for (const source in sourceData) {
        const rate = (sourceData[source].hires / sourceData[source].applications) * 100;
        sourceQualityPrediction[source] = {
            rate: parseFloat(rate.toFixed(1)),
            applications: sourceData[source].applications,
            hires: sourceData[source].hires
        };
    }

    return {
      timeToHire: timeToHirePrediction,
      sourceQuality: sourceQualityPrediction
    };

  } catch (e) {
    Logger.log('Error in getHiringPredictions: ' + e.message);
    return { error: e.message };
  }
}
// --- END: NEW HIRING PREDICTIONS FUNCTION ---

// --- START: NEW USER AUTHENTICATION FUNCTION ---

/**
 * Checks if the current user's email is in the 'Permissions' sheet.
 * This acts as the secure gatekeeper for the web app.
 * @returns {object} An object containing authorization status and user email.
 */
function checkUserAccess() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const permissionsSheet = ss.getSheetByName('Permissions');
    
    if (!permissionsSheet) {
      // If there's no Permissions sheet, deny access by default for security.
      return { isAuthorized: false, userEmail: userEmail };
    }

    const emailColumn = permissionsSheet.getRange("A2:A").getValues();
    const authorizedEmails = new Set(emailColumn.flat().map(email => String(email).trim().toLowerCase()).filter(String));

    if (authorizedEmails.has(userEmail)) {
      return { isAuthorized: true, userEmail: userEmail };
    } else {
      return { isAuthorized: false, userEmail: userEmail };
    }
  } catch (e) {
    Logger.log('Error in checkUserAccess: ' + e.message);
    return { isAuthorized: false, error: e.message };
  }
}

// --- END: NEW USER AUTHENTICATION FUNCTION ---

// --- START: CHANGE REQUEST FUNCTIONS ---

function submitChangeRequest(requestData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Org Chart Requests');
    if (!sheet) {
      throw new Error('"Org Chart Requests" sheet not found.');
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const requestId = 'REQ-' + new Date().getTime();
    let folderUrl = '';

    // Handle file uploads
    if (requestData.files && requestData.files.length > 0) {
      const parentFolder = DriveApp.getFolderById(CHANGE_REQUESTS_FOLDER_ID);
      const requestFolder = parentFolder.createFolder(requestId);
      
      requestData.files.forEach(file => {
        const decodedContent = Utilities.base64Decode(file.content);
        const blob = Utilities.newBlob(decodedContent, file.mimeType, file.name);
        requestFolder.createFile(blob);
      });
      
      folderUrl = requestFolder.getUrl();
    }

    const newRow = headers.map(header => {
      switch (header) {
        case 'RequestID':
          return requestId;
        case 'RequestorEmail':
          return Session.getActiveUser().getEmail();
        case 'SubmissionTimestamp':
          return new Date();
        case 'Status':
          return 'Pending';
        case 'SupportingDocuments':
          return folderUrl;
        // --- NEW: Explicitly handle the new fields ---
        case 'EmployeeID':
          return requestData.EmployeeID || '';
        case 'DateHired':
          return requestData.DateHired || '';
        case 'DateOfBirth':
          return requestData.DateOfBirth || '';
        case 'PositionStatus':
            return requestData.PositionStatus || '';
        default:
          // Use a more robust check for other fields
          return requestData.hasOwnProperty(header) ? requestData[header] : '';
      }
    });
    
    sheet.appendRow(newRow);
    return 'Request submitted successfully.';
  } catch (e) {
    Logger.log('Error in submitChangeRequest: ' + e.message);
    throw new Error('Failed to submit request. ' + e.message);
  }
}

function getApproverList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Permissions');
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const emailIndex = headers.indexOf('EMAIL');
    const approverIndex = headers.indexOf('Is Approver');

    if (emailIndex === -1 || approverIndex === -1) {
      return [];
    }

    const approvers = data
      .filter(row => (row[approverIndex] || '').toString().trim().toLowerCase() === 'x')
      .map(row => ({ email: row[emailIndex] }))
      .sort((a, b) => a.email.localeCompare(b.email));
      
    return approvers;
  } catch (e) {
    Logger.log('Error in getApproverList: ' + e.message);
    throw new Error('Could not retrieve approver list.');
  }
}
// --- END: CHANGE REQUEST FUNCTIONS ---

/**
 * =================================================================================================
 * TALENT & SUCCESSION PLANNING FUNCTIONS (Reads from separate sheet)
 * =================================================================================================
 */

/**
 * Main function to get talent analytics. Reads employee data from the main hub
 * and performance/succession data from the separate Talent Data spreadsheet.
 * @returns {object} An object containing lists of employees for promotion, high potential, and succession.
 */
function getTalentAnalyticsData() {
  try {
    // Open both spreadsheets
    const hubSs = SpreadsheetApp.getActiveSpreadsheet();
    const talentSs = SpreadsheetApp.openById(TALENT_DATA_SPREADSHEET_ID);

    // Get current active employees from the main hub sheet
    const mainSheet = hubSs.getSheets()[0];
    const mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    const headerMap = new Map(headers.map((h, i) => [h.trim(), i]));
    const currentEmployees = mainData.filter(row => row[headerMap.get('Employee ID')] && (row[headerMap.get('Position Status')] || '').toUpperCase() !== 'INACTIVE');

    // Get performance data from the separate Talent Data sheet
    const performanceSheet = talentSs.getSheetByName('Performance Data');
    const performanceMap = new Map();
    if (performanceSheet && performanceSheet.getLastRow() > 1) {
      const perfData = performanceSheet.getRange(2, 1, performanceSheet.getLastRow() - 1, performanceSheet.getLastColumn()).getValues();
      const perfHeader = performanceSheet.getRange(1, 1, 1, performanceSheet.getLastColumn()).getValues()[0];
      const empIdIndex = perfHeader.indexOf('Employee ID');
      const scoreIndex = perfHeader.indexOf('Overall Score (1-5)');
      const competencyIndex = perfHeader.indexOf('Competency Score (1-5)');

      perfData.forEach(row => {
        const empId = row[empIdIndex];
        if (empId) { // Store the most recent record for each employee
          performanceMap.set(String(empId).trim(), {
            overallScore: row[scoreIndex],
            competencyScore: row[competencyIndex]
          });
        }
      });
    }

    const promotionReady = [];
    const highPotentials = [];
    const today = new Date();

    currentEmployees.forEach(row => {
      const empId = String(row[headerMap.get('Employee ID')]).trim();
      const performance = performanceMap.get(empId);
      const dateHired = new Date(row[headerMap.get('Date Hired')]);
      const tenureYears = !isNaN(dateHired) ? (today.getTime() - dateHired.getTime()) / (31557600000) : 0;

      const employeeRecord = {
          name: row[headerMap.get('Employee Name')],
          jobTitle: row[headerMap.get('Job Title')],
          department: row[headerMap.get('Department')],
          tenure: tenureYears.toFixed(1) + ' years',
          performance: 'N/A',
          competency: 'N/A'
      };

      if (performance) {
        employeeRecord.performance = performance.overallScore;
        employeeRecord.competency = performance.competencyScore;

        if (performance.overallScore >= 4 && performance.competencyScore >= 4 && tenureYears >= 1) {
          promotionReady.push(employeeRecord);
        }

        if (performance.overallScore === 5 && performance.competencyScore === 5) {
          highPotentials.push(employeeRecord);
        }
      }
    });

    // Get manually defined succession plan from the Talent Data sheet
    const employeeNameMap = new Map(currentEmployees.map(r => [String(r[headerMap.get('Employee ID')]).trim(), r[headerMap.get('Employee Name')]]));
    const successionPlan = getSuccessionPlanData(talentSs, employeeNameMap);

    return {
      promotionReady: promotionReady.sort((a,b) => b.performance - a.performance),
      highPotentials: highPotentials,
      successionPlan: successionPlan
    };

  } catch (e) {
    Logger.log('Error in getTalentAnalyticsData: ' + e.message);
    if (e.message.includes("You do not have permission")) {
      return { error: "Permission denied. Please ensure the hub has access to the Talent Data Google Sheet." };
    }
    return { error: e.message };
  }
}

/**
 * Helper function to read 'Succession Planning' sheet from the talent spreadsheet.
 */
function getSuccessionPlanData(talentSpreadsheet, employeeNameMap) {
    const sheet = talentSpreadsheet.getSheetByName('Succession Planning');
    if (!sheet || sheet.getLastRow() < 2) {
        return [];
    }
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    return data.map(row => ({
        keyPosition: row[1],
        successorId: String(row[2]).trim(),
        successorName: employeeNameMap.get(String(row[2]).trim()) || 'Unknown Employee',
        readiness: row[3],
        notes: row[4]
    }));
}

/**
 * Fetches a list of all employees from the competency sheet.
 * This version correctly finds employees using 'EMPLOYEE ID' and 'EMPLOYEE NAME'.
 * @returns {Array<Object>} An array of objects with employee code and name.
 */
function getCompetencyEmployeeList() {
  try {
    const ss = SpreadsheetApp.openById(COMPETENCY_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Competency Matrix');
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Get headers from the first row.

    // Use the correct header names from your spreadsheet.
    const codeIndex = headers.indexOf('EMPLOYEE ID');
    const nameIndex = headers.indexOf('EMPLOYEE NAME');

    if (codeIndex === -1 || nameIndex === -1) {
      // If the columns aren't found, return a descriptive error.
      return { error: "Could not find 'EMPLOYEE ID' or 'EMPLOYEE NAME' columns in the Competency Matrix sheet." };
    }

    const employees = data
      .filter(row => row[codeIndex]) // Filter out rows without an Employee ID.
      .map(row => ({
        code: row[codeIndex],
        name: row[nameIndex]
      }));

    return employees.sort((a, b) => a.name.localeCompare(b.name));
  } catch (e) {
    Logger.log('Error in getCompetencyEmployeeList: ' + e.message);
    return { error: `Could not access the Competency Spreadsheet. Details: ${e.message}` };
  }
}

/**
 * Fetches the detailed competency profile for a single employee.
 * This version dynamically filters to only show competencies that have a required or actual score for the employee.
 * @param {string} employeeId The Employee ID to look up.
 * @returns {Object} An object containing the structured competency data for the charts.
 */
function getEmployeeCompetencyProfile(employeeId) {
  if (!employeeId) return { error: 'No employee ID provided.' };

  try {
    const ss = SpreadsheetApp.openById(COMPETENCY_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Competency Matrix');
    if (!sheet) {
      return { error: "The 'Competency Matrix' sheet was not found." };
    }
    const data = sheet.getDataRange().getValues();
    const rawHeaders = data.shift();
    const headers = rawHeaders.map(header => header.replace(/\n|\r/g, ' ').trim());

    const empIdIndex = headers.indexOf('EMPLOYEE ID');
    if (empIdIndex === -1) {
      return { error: "Header 'EMPLOYEE ID' not found in Competency Matrix." };
    }

    const employeeRow = data.find(row => String(row[empIdIndex]).trim() === String(employeeId).trim());

    if (!employeeRow) {
      return { error: 'Employee not found in Competency Matrix.' };
    }

    const profile = {
      core: { labels: [], actual: [], required: [] },
      leadership: { labels: [], actual: [], required: [] },
      technical: { labels: [], actual: [], required: [] }
    };

    const coreCompetencies = ["TRUSTWORTHINESS", "ENTREPRENEURIAL SPIRIT", "INNOVATION", "LEADERSHIP", "RESPECT FOR THE INDIVIDUAL"];
    const leadershipCompetencies = ["LEADERSHIP BY EXAMPLE", "DRIVE FOR RESULTS", "COACHING FOR SUCCESS", "INSPIRING LOYAL AND TRUST", "WORKING ACROSS TEAMS", "TALENT MANAGEMENT AND DEVELOPMENT", "EMPOWERMENT", "COMMUNICATION", "EXECUTIVE DISPOSITION"];
    const technicalCompetencies = ["PROJECT MANAGEMENT", "LEAN THINKING PRINCIPLES", "PROCESS STANDARDIZATION", "OPERATIONAL EXPERTISE", "COST MANAGEMENT", "DATA-DRIVEN DECISION MAKING", "MANAGEMENT OF WORK SYSTEMS/BUSINESS PROCESS ORIENTATION"];
    
    const allCompetencyNames = [...coreCompetencies, ...leadershipCompetencies, ...technicalCompetencies];
    
    allCompetencyNames.forEach(competencyName => {
      let requiredIndex = -1;
      let actualIndex = -1;

      requiredIndex = headers.indexOf(`${competencyName} [Required]`);
      actualIndex = headers.indexOf(`${competencyName} [Actual]`);

      if (requiredIndex === -1 || actualIndex === -1) {
        const indices = headers.map((h, i) => (h.trim() === competencyName ? i : -1)).filter(i => i !== -1);
        if (indices.length >= 2) {
          requiredIndex = indices[0];
          actualIndex = indices[1];
        }
      }
      
      if (requiredIndex !== -1 && actualIndex !== -1) {
        const requiredValue = parseFloat(employeeRow[requiredIndex]) || 0;
        const actualValue = parseFloat(employeeRow[actualIndex]) || 0;
        
        // --- THIS IS THE KEY IMPROVEMENT ---
        // Only include the competency if it's relevant to the employee (score > 0)
        if (requiredValue > 0 || actualValue > 0) {
          let category;
          if (coreCompetencies.includes(competencyName)) category = profile.core;
          else if (leadershipCompetencies.includes(competencyName)) category = profile.leadership;
          else if (technicalCompetencies.includes(competencyName)) category = profile.technical;

          if (category && !category.labels.includes(competencyName)) {
            category.labels.push(competencyName);
            category.required.push(requiredValue);
            category.actual.push(actualValue);
          }
        }
      }
    });

    return profile;

  } catch (e) {
    Logger.log('Error in getEmployeeCompetencyProfile: ' + e.message + ' Stack: ' + e.stack);
    return { error: e.message };
  }
}

/**
 * REVISED - Fetches and calculates competency analytics.
 * This version is made more robust to handle hidden characters and case sensitivity in headers.
 * It also includes logging for easier debugging.
 * @param {string} employeeId The Employee ID to look up.
 * @returns {Object} An object containing all competency analytics.
 */
function getCompetencyAnalytics(employeeId) {
  if (!employeeId) return { error: 'No employee ID provided.' };

  try {
    const hubSs = SpreadsheetApp.getActiveSpreadsheet();
    const compSs = SpreadsheetApp.openById(COMPETENCY_SPREADSHEET_ID);
    const mainSheet = hubSs.getSheets()[0];
    const compSheet = compSs.getSheetByName('Competency Matrix');
    const historySheet = compSs.getSheetByName('Competency History');

    // (This part for current data remains the same)
    if (!compSheet) return { error: "The 'Competency Matrix' sheet was not found." };
    const compData = compSheet.getDataRange().getValues();
    const rawHeaders = compData.shift();
    const headers = rawHeaders.map(header => header.replace(/\n|\r/g, ' ').trim());
    const compEmpIdIndex = headers.indexOf('EMPLOYEE ID');
    if (compEmpIdIndex === -1) return { error: "Header 'EMPLOYEE ID' not found in Competency Matrix." };
    const employeeRow = compData.find(row => String(row[compEmpIdIndex]).trim() === String(employeeId).trim());
    if (!employeeRow) return { error: 'Employee not found in Competency Matrix.' };

    const employeeDepartmentMap = new Map();
    let selectedEmployeeDept = null;
    if (mainSheet) {
      const mainData = mainSheet.getDataRange().getValues();
      const mainHeaders = mainData.shift();
      const mainEmpIdIndex = mainHeaders.indexOf('Employee ID');
      const mainDeptIndex = mainHeaders.indexOf('Department');
      if (mainEmpIdIndex !== -1 && mainDeptIndex !== -1) {
        mainData.forEach(row => {
          const empId = String(row[mainEmpIdIndex]).trim();
          const dept = row[mainDeptIndex];
          if (empId && dept) employeeDepartmentMap.set(empId, dept);
        });
      }
      selectedEmployeeDept = employeeDepartmentMap.get(String(employeeId).trim());
    }

    const profile = {
      core: { labels: [], actual: [], required: [], team: [] },
      leadership: { labels: [], actual: [], required: [], team: [] },
      technical: { labels: [], actual: [], required: [], team: [] }
    };
    const allGaps = [];
    const coreCompetencies = ["TRUSTWORTHINESS", "ENTREPRENEURIAL SPIRIT", "INNOVATION", "LEADERSHIP", "RESPECT FOR THE INDIVIDUAL"];
    const leadershipCompetencies = ["LEADERSHIP BY EXAMPLE", "DRIVE FOR RESULTS", "COACHING FOR SUCCESS", "INSPIRING LOYAL AND TRUST", "WORKING ACROSS TEAMS", "TALENT MANAGEMENT AND DEVELOPMENT", "EMPOWERMENT", "COMMUNICATION", "EXECUTIVE DISPOSITION"];
    const technicalCompetencies = ["PROJECT MANAGEMENT", "LEAN THINKING PRINCIPLES", "PROCESS STANDARDIZATION", "OPERATIONAL EXPERTISE", "COST MANAGEMENT", "DATA-DRIVEN DECISION MAKING", "MANAGEMENT OF WORK SYSTEMS/BUSINESS PROCESS ORIENTATION"];
    const allCompetencyNames = [...coreCompetencies, ...leadershipCompetencies, ...technicalCompetencies];
    allCompetencyNames.forEach(competencyName => {
      let requiredIndex = -1, actualIndex = -1;
      const indices = headers.map((h, i) => (h.trim() === competencyName ? i : -1)).filter(i => i !== -1);
      if (indices.length >= 2) {
        requiredIndex = indices[0];
        actualIndex = indices[1];
      }
      if (requiredIndex !== -1 && actualIndex !== -1) {
        const requiredValue = parseFloat(employeeRow[requiredIndex]) || 0;
        const actualValue = parseFloat(employeeRow[actualIndex]) || 0;
        if (requiredValue > 0 || actualValue > 0) {
          allGaps.push({ name: competencyName, gap: actualValue - requiredValue });
          let teamTotal = 0, teamMemberCount = 0;
          if (selectedEmployeeDept) {
            compData.forEach(row => {
              const currentEmpId = String(row[compEmpIdIndex]).trim();
              if (employeeDepartmentMap.get(currentEmpId) === selectedEmployeeDept) {
                teamTotal += parseFloat(row[actualIndex]) || 0;
                teamMemberCount++;
              }
            });
          }
          const teamAverage = (teamMemberCount > 0) ? (teamTotal / teamMemberCount) : 0;
          let category;
          if (coreCompetencies.includes(competencyName)) category = profile.core;
          else if (leadershipCompetencies.includes(competencyName)) category = profile.leadership;
          else if (technicalCompetencies.includes(competencyName)) category = profile.technical;
          if (category && !category.labels.includes(competencyName)) {
            category.labels.push(competencyName);
            category.required.push(requiredValue);
            category.actual.push(actualValue);
            category.team.push(parseFloat(teamAverage.toFixed(2)));
          }
        }
      }
    });
    allGaps.sort((a, b) => b.gap - a.gap);
    const strengths = allGaps.filter(item => item.gap > 0).slice(0, 3);
    const gaps = allGaps.filter(item => item.gap < 0).reverse().slice(0, 3);
    const allActuals = [...profile.core.actual, ...profile.leadership.actual, ...profile.technical.actual];
    const allRequired = [...profile.core.required, ...profile.leadership.required, ...profile.technical.required];
    const overallActual = allActuals.length > 0 ? (allActuals.reduce((a, b) => a + b, 0) / allActuals.length) : 0;
    const overallRequired = allRequired.length > 0 ? (allRequired.reduce((a, b) => a + b, 0) / allRequired.length) : 0;

    // --- REVISED: Process Historical Data with robust header matching and logging ---
    const historicalData = { labels: [], datasets: [] };
    if (historySheet) {
      Logger.log("Processing 'Competency History' sheet.");
      const histData = historySheet.getDataRange().getValues();
      const rawHistHeaders = histData.shift();
      // Clean headers: remove newlines, trim, and convert to uppercase for robust matching.
      const histHeaders = rawHistHeaders.map(h => (h || '').toString().replace(/\n|\r/g, ' ').trim().toUpperCase());
      
      Logger.log("Cleaned History Headers: " + JSON.stringify(histHeaders));

      // Find the correct column indices using the cleaned, uppercase headers.
      const histEmpIdIndex = histHeaders.indexOf('EMPLOYEE ID');
      const histYearIndex = histHeaders.indexOf('ASSESSMENT YEAR');
      const histActualAvgIndex = histHeaders.indexOf('ACTUAL AVERAGE');
      const histGapAvgIndex = histHeaders.indexOf('GAP AVERAGE');

      Logger.log(`Header Indices Found: empId=${histEmpIdIndex}, year=${histYearIndex}, actualAvg=${histActualAvgIndex}, gapAvg=${histGapAvgIndex}`);

      if (histEmpIdIndex !== -1 && histYearIndex !== -1 && histActualAvgIndex !== -1 && histGapAvgIndex !== -1) {
        Logger.log("All required history headers found. Filtering for employee: " + employeeId);

        const employeeHistoryRows = histData
          .filter(row => String(row[histEmpIdIndex]).trim() === String(employeeId).trim())
          .sort((a, b) => a[histYearIndex] - b[histYearIndex]);

        Logger.log("Found " + employeeHistoryRows.length + " historical rows for this employee.");

        if (employeeHistoryRows.length > 0) {
            historicalData.labels = employeeHistoryRows.map(row => row[histYearIndex].toString());
            
            historicalData.datasets.push({
                label: "Overall Actual Average",
                data: employeeHistoryRows.map(row => parseFloat(row[histActualAvgIndex]) || null)
            });

            historicalData.datasets.push({
                label: "Overall Gap Average",
                data: employeeHistoryRows.map(row => parseFloat(row[histGapAvgIndex]) || null)
            });

            Logger.log("Successfully prepared historical data for chart: " + JSON.stringify(historicalData));
        }
      } else {
        Logger.log("One or more required headers were NOT found in 'Competency History' sheet.");
      }
    } else {
        Logger.log("'Competency History' sheet not found.");
    }

    return {
      radarData: profile,
      strengths: strengths,
      gaps: gaps,
      overall: {
        actual: parseFloat(overallActual.toFixed(2)),
        required: parseFloat(overallRequired.toFixed(2))
      },
      history: historicalData
    };

  } catch (e) {
    Logger.log('Error in getCompetencyAnalytics: ' + e.message + ' Stack: ' + e.stack);
    return { error: e.message };
  }
}

/**
 * Gets the org chart data as it would look *if* a specific change request were applied.
 * Does NOT save any changes to the main sheet.
 * @param {string} requestId The ID of the change request to preview.
 * @returns {Array<Object>} An array of employee/position objects representing the preview state.
 * @throws {Error} If the request ID is not found or simulation fails.
 */
function getPreviewOrgChartData(requestId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheets()[0];
    const requestSheet = ss.getSheetByName('Org Chart Requests');

    if (!requestSheet) throw new Error('"Org Chart Requests" sheet not found.');
    if (mainSheet.getLastRow() < 2) throw new Error('Main org chart data sheet is empty.');

    // --- Get Current Live Data as Objects ---
    const lastCol = mainSheet.getLastColumn();
    const mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, lastCol).getValues();
    const mainHeaders = mainSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    const liveObjects = mainData.map(row => {
      let employee = {};
      mainHeaders.forEach((header, i) => {
        const key = header.toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-9]/gi, '');
        // FIX: Ensure dates are formatted here before they are even used
        if (row[i] instanceof Date) {
          employee[key] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          employee[key] = row[i];
        }
      });
      employee.nodeId = String(employee['positionid'] || '');
      employee.managerId = String(employee['reportingtoid'] || '');
      return employee;
    });
    
    // --- Get Request Details ---
    const requestDataRange = requestSheet.getDataRange();
    const requestValues = requestDataRange.getValues();
    const requestHeaders = requestValues[0];
    const reqIdCol = requestHeaders.indexOf('RequestID');
    let requestRowData = null;
    for (let i = 1; i < requestValues.length; i++) {
      if (requestValues[i][reqIdCol] === requestId) {
        requestRowData = {};
        requestHeaders.forEach((header, index) => {
          requestRowData[header] = requestValues[i][index];
        });
        break;
      }
    }
    if (!requestRowData) throw new Error(`Request ID ${requestId} not found.`);

    // --- Simulate the Change on a deep copy of the objects ---
    let previewObjects = JSON.parse(JSON.stringify(liveObjects));
    const requestType = requestRowData['RequestType'];
    let changedPositionIds = new Set();
    let changeDescription = '';

    if (requestType.includes('Transfer') || requestType.includes('Promotion')) {
      const employeeId = requestRowData['EmployeeID'];
      const newPositionId = requestRowData['NewPositionID'];
      const employeeName = requestRowData['EmployeeName'];
      const effectiveDate = requestRowData['EffectiveDate'];
      changeDescription = `${requestType}: ${employeeName} to ${newPositionId}`;

      const oldPosition = previewObjects.find(p => p.positionid === requestRowData['CurrentPositionID']);
      const newPosition = previewObjects.find(p => p.positionid === newPositionId);

      if (oldPosition) {
        // oldPosition.employeeid = ''; // DO NOT CLEAR - This was the bug
        // oldPosition.employeename = ''; // DO NOT CLEAR - This was the bug
        // oldPosition.status = 'VACANT'; // DO NOT CHANGE STATUS
        oldPosition.isPreviewChange = true;
        oldPosition.changeType = 'VACATED BY ' + employeeName;
        changedPositionIds.add(oldPosition.positionid);
      }

      if (newPosition) {
        newPosition.employeeid = employeeId;
        newPosition.employeename = employeeName;
        newPosition.status = requestType.toUpperCase();
        newPosition.isPreviewChange = true;
        newPosition.changeType = requestType.toUpperCase() + ' - ' + employeeName;
        newPosition.effectiveDate = effectiveDate ? Utilities.formatDate(new Date(effectiveDate), Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
        changedPositionIds.add(newPosition.positionid);
      } else {
        Logger.log(`Warning: New position ${newPositionId} not found during preview generation.`);
      }

    } else if (requestType.includes('replacement for vacancy')) {
      const vacantPositionId = requestRowData['VacantPositionID'];
      const newEmployeeId = requestRowData['NewEmployeeID'];
      const newEmployeeName = requestRowData['NewEmployeeName'];
      const effectiveDate = requestRowData['EffectiveDate'];
      changeDescription = `Fill Vacancy: ${newEmployeeName} in ${vacantPositionId}`;

      const position = previewObjects.find(p => p.positionid === vacantPositionId);
      if (position) {
        position.employeeid = newEmployeeId;
        position.employeename = newEmployeeName;
        position.status = 'FILLED VACANCY';
        position.isPreviewChange = true;
        position.changeType = 'Filled Vacancy';
        position.effectiveDate = effectiveDate ? Utilities.formatDate(new Date(effectiveDate), Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
        changedPositionIds.add(vacantPositionId);
      } else {
        Logger.log(`Warning: Vacant position ${vacantPositionId} not found during preview generation.`);
      }

    } else if (requestType.includes('newly created position')) {
      changeDescription = `New Position: ${requestRowData['NewJobTitle']} filled by ${requestRowData['NewEmployeeName']}`;
      const tempNewPositionId = `PREVIEW-${requestId}`;
      
      const newPositionObject = {
        positionid: tempNewPositionId,
        jobtitle: requestRowData['NewJobTitle'],
        level: requestRowData['NewLevel'],
        division: requestRowData['Division'],
        group: requestRowData['Group'],
        department: requestRowData['Department'],
        section: requestRowData['Section'],
        reportingtoid: requestRowData['ReportingToId'],
        employeeid: requestRowData['NewEmployeeID'],
        employeename: requestRowData['NewEmployeeName'],
        status: 'NEW HIRE',
        isPreviewChange: true,
        changeType: 'New Position',
        nodeId: tempNewPositionId,
        managerId: requestRowData['ReportingToId']
      };
      
      previewObjects.push(newPositionObject);
      changedPositionIds.add(tempNewPositionId);
    }
    
    // --- START: NEW HIERARCHY REBUILD LOGIC ---
    // After simulating changes, the reporting structure might be broken.
    // We need to rebuild the managerId links based on the new state of previewObjects.
    const newEmployeeIdToPositionIdMap = new Map();
    previewObjects.forEach(p => {
      if (p.employeeid) {
        newEmployeeIdToPositionIdMap.set(String(p.employeeid).trim(), p.positionid);
      }
    });

    previewObjects.forEach(p => {
      const managerEmployeeId = (p.reportingtoid || '').toString().trim();
      if (managerEmployeeId) {
        // Find the manager's NEW position ID from our fresh map
        const newManagerPositionId = newEmployeeIdToPositionIdMap.get(managerEmployeeId);
        if (newManagerPositionId) {
          p.managerId = newManagerPositionId;
        } else {
          // If the manager ID points to an employee who no longer exists in a position
          // (e.g., they were the one transferred out), this link should be broken.
          p.managerId = ''; 
          Logger.log(`Preview Warning: Could not find new position for manager with Employee ID ${managerEmployeeId}. Breaking link for ${p.positionid}.`);
        }
      } else {
        // No manager employee ID, so no manager link.
        p.managerId = '';
      }
    });
    // --- END: NEW HIERARCHY REBUILD LOGIC ---

    // --- FINAL DATA SANITIZATION ---
    // Ensure all objects in the array are clean for JSON serialization, especially dates.
    const sanitizedObjects = previewObjects.map(obj => {
      const sanitizedObj = {};
      for (const key in obj) {
        if (obj[key] instanceof Date) {
          sanitizedObj[key] = Utilities.formatDate(obj[key], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          sanitizedObj[key] = obj[key];
        }
      }
      return sanitizedObj;
    });


    // --- Return the final object ---
    return {
        chartData: sanitizedObjects, // Use the sanitized data
        highlightIds: Array.from(changedPositionIds),
        changeDescription: changeDescription || requestType
     };

  } catch (e) {
    Logger.log('Error in getPreviewOrgChartData for request ID ' + requestId + ': ' + e.message + ' Stack: ' + e.stack);
    return { error: 'Failed to generate preview data. ' + e.message };
  }
}
