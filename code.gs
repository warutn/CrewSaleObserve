/**
 * Crew Random Inspection Web App - Backend
 * Developed by Antigravity (Google DeepMind Team)
 * 
 * Manages Google Sheets integration, dynamic schema mapping,
 * server-side whitelist validation, and secure batch synchronization.
 */

const CONFIG = {
  SPREADSHEET_ID: '1m7NLUTS9QTA8oZL7A6tLU_p5vF0PVt5J54YlUr3f4ag',
  SHEET_RECORD: 'Record',
  SHEET_CHECKLIST: 'CheckList',
  SHEET_USERS: 'Users',
  CHECKLIST_RANGE: 'A7:A26'
};

/**
 * Serves the HTML Single Page Application (SPA) with Server-Side Whitelist Validation.
 */
function doGet(e) {
  const activeUserEmail = Session.getActiveUser().getEmail();
  
  if (!activeUserEmail) {
    return HtmlService.createHtmlOutput("<p style='font-family:sans-serif; text-align:center; margin-top:50px; color:#555;'>No email session found. Please log into your Google account.</p>");
  }

  const isAuthorized = checkUserAuthorization(activeUserEmail);
  if (!isAuthorized) {
    return HtmlService.createHtmlOutput("<p style='font-family:sans-serif; text-align:center; margin-top:50px; color:#555;'>Access Denied. You do not have permissions to access this portal.</p>")
      .setTitle('Access Denied');
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Crew Inspection WebApp')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Includes HTML/CSS/JS files inside the main template.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Safely opens the target Google Spreadsheet.
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch (e) {
    console.error('Error opening Spreadsheet: ' + e.message);
    try {
      return SpreadsheetApp.getActiveSpreadsheet();
    } catch (err) {
      throw new Error('Could not open spreadsheet. Please check SPREADSHEET_ID and Apps Script permissions: ' + e.message);
    }
  }
}

/**
 * Server-side function to check if an email address is whitelisted.
 */
function checkUserAuthorization(emailAddress) {
  try {
    const ss = getSpreadsheet();
    let usersSheet = ss.getSheetByName(CONFIG.SHEET_USERS);
    if (!usersSheet) {
      // Auto create users sheet and register the active user
      usersSheet = ss.insertSheet(CONFIG.SHEET_USERS);
      usersSheet.getRange(1, 1, 1, 4).setValues([['Email', 'Name', 'Role', 'Is Active']]);
      usersSheet.getRange(2, 1, 1, 4).setValues([[emailAddress, 'Initial Inspector', 'Administrator', 'Y']]);
      SpreadsheetApp.flush();
      return true; // Authorized since we just created it for them
    }

    const usersData = usersSheet.getDataRange().getValues();
    const headers = usersData[0].map(h => h.toString().toLowerCase().trim());
    const emailColIndex = headers.indexOf('email');
    const activeColIndex = headers.indexOf('is active');

    if (emailColIndex !== -1) {
      for (let i = 1; i < usersData.length; i++) {
        const row = usersData[i];
        const email = row[emailColIndex].toString().toLowerCase().trim();
        if (email === emailAddress.toLowerCase().trim()) {
          const isActive = activeColIndex !== -1 ? row[activeColIndex].toString().toUpperCase().trim() : 'Y';
          if (isActive === 'Y' || isActive === 'YES' || isActive === 'TRUE') {
            return true;
          }
        }
      }
    }
    return false;
  } catch (e) {
    console.error('Error checking authorization: ' + e.toString());
    return false; // Safely deny access on error
  }
}

/**
 * Initializes and fetches application configurations (Checklist and user details).
 */
function getAppInitializationData() {
  const data = {
    checklist: [],
    userEmail: '',
    userName: '',
    isAuthorized: true, // Already validated via doGet server check
    error: null
  };

  try {
    const ss = getSpreadsheet();
    const activeUserEmail = Session.getActiveUser().getEmail();
    data.userEmail = activeUserEmail || 'offline_developer@airasia.com';

    // Fetch User's name from sheet
    let usersSheet = ss.getSheetByName(CONFIG.SHEET_USERS);
    if (usersSheet) {
      const usersData = usersSheet.getDataRange().getValues();
      const headers = usersData[0].map(h => h.toString().toLowerCase().trim());
      const emailColIndex = headers.indexOf('email');
      const nameColIndex = headers.indexOf('name');
      
      if (emailColIndex !== -1 && nameColIndex !== -1) {
        for (let i = 1; i < usersData.length; i++) {
          const row = usersData[i];
          if (row[emailColIndex].toString().toLowerCase().trim() === data.userEmail.toLowerCase().trim()) {
            data.userName = row[nameColIndex].toString().trim();
            break;
          }
        }
      }
    }
    
    if (!data.userName) data.userName = 'Observer';

    // Fetch Checklist rows A7:A26
    const checklistSheet = ss.getSheetByName(CONFIG.SHEET_CHECKLIST);
    if (!checklistSheet) {
      throw new Error(`Sheet "${CONFIG.SHEET_CHECKLIST}" not found in spreadsheet. Please create it and place questions in rows A7:A26.`);
    }

    const rangeValues = checklistSheet.getRange(CONFIG.CHECKLIST_RANGE).getValues();
    const questions = [];
    
    rangeValues.forEach((row, idx) => {
      const qText = row[0] ? row[0].toString().trim() : '';
      if (qText) {
        questions.push({
          id: idx + 1,
          question: qText
        });
      }
    });

    data.checklist = questions;
    return data;

  } catch (e) {
    console.error('Initialization error: ' + e.toString());
    data.error = e.toString();
    return data;
  }
}

/**
 * Submits a batch of inspection audits to the 'Record' sheet.
 */
function submitInspections(inspectionArray) {
  if (!inspectionArray || !Array.isArray(inspectionArray) || inspectionArray.length === 0) {
    return { success: false, message: 'No inspection records received.' };
  }

  try {
    const ss = getSpreadsheet();
    let recordSheet = ss.getSheetByName(CONFIG.SHEET_RECORD);
    
    if (!recordSheet) {
      recordSheet = ss.insertSheet(CONFIG.SHEET_RECORD);
    }

    const initData = getAppInitializationData();
    const checklist = initData.checklist;
    
    if (checklist.length === 0) {
      throw new Error('Checklist is empty. Cannot determine record schema.');
    }

    const standardHeaders = [
      'Timestamp',
      'Observation Date',
      'Flight No.',
      'Destination',
      'Observer Email',
      'Observer Name'
    ];

    const dynamicHeaders = [...standardHeaders];
    checklist.forEach(item => {
      dynamicHeaders.push(`Q${item.id}: ${item.question}`);
      dynamicHeaders.push(`Q${item.id} Comment`);
    });
    dynamicHeaders.push('Overall Comments');

    const currentData = recordSheet.getDataRange().getValues();
    const isSheetEmpty = (currentData.length === 1 && currentData[0][0] === '') || currentData.length === 0;

    if (isSheetEmpty) {
      recordSheet.getRange(1, 1, 1, dynamicHeaders.length).setValues([dynamicHeaders]);
      recordSheet.getRange(1, 1, 1, dynamicHeaders.length)
        .setFontWeight('bold')
        .setBackground('#E11B22')
        .setFontColor('#FFFFFF')
        .setHorizontalAlignment('center');
      recordSheet.setFrozenRows(1);
    }

    const rowsToWrite = inspectionArray.map(audit => {
      const row = [
        audit.timestamp ? new Date(audit.timestamp) : new Date(),
        audit.observationDate || '',
        (audit.flightNo || '').toUpperCase().trim(),
        (audit.destination || '').toUpperCase().trim(),
        audit.observerEmail || '',
        audit.observerName || ''
      ];

      checklist.forEach(item => {
        const answerObj = (audit.checklistAnswers || []).find(ans => ans.id === item.id) || {};
        row.push(answerObj.status || 'N/A');
        row.push(answerObj.comment || '');
      });

      row.push(audit.overallComments || '');
      return row;
    });

    const startRow = recordSheet.getLastRow() + 1;
    recordSheet.getRange(startRow, 1, rowsToWrite.length, dynamicHeaders.length).setValues(rowsToWrite);
    
    recordSheet.getRange(startRow, 1, rowsToWrite.length, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
    recordSheet.getRange(startRow, 2, rowsToWrite.length, 1).setNumberFormat('yyyy-MM-dd');

    SpreadsheetApp.flush();
    return { success: true, count: inspectionArray.length };

  } catch (e) {
    console.error('Submission error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}
