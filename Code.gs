function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Salesforce')
    .addItem('Update All Sheets', 'updateAllSheets')
    .addItem('Update Single Sheet (Prompt)', 'updateSingleSheetPrompted')
    .addItem('Compile Sheets for Domo', 'compileSheetsToDomo')
    .addToUi();
}

function updateSingleSheetPrompted() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Enter Parent Sheet Name', 'Please input the parent sheet name you wish to update:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const parentSheetName = result.getResponseText();
    refreshToken();
    makeRequestSoql(parentSheetName);  // <-- Pass the user input to the function.
    setReportDateRange();
  }
}

function updateAllSheets() {
  refreshToken();
  makeRequestSoql();
  setReportDateRange();
}

function refreshToken() {
  const scriptProps = PropertiesService.getScriptProperties();
  const userProps = PropertiesService.getUserProperties();

  const clientId = scriptProps.getProperty("SFClientID");
  const clientSecret = scriptProps.getProperty("SFClientSecret");
  
  // Check if clientId or clientSecret are missing
  if (!clientId || !clientSecret) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.prompt('Configuration Missing', 'Please add SFClientID and SFClientSecret in script properties.', ui.ButtonSet.OK);
      if (response.getSelectedButton() == ui.Button.OK) {
          Logger.log(`User didn't provide the necessary info.`)
      }
      return;
  }

  let currentOAuthProps = userProps.getProperty('oauth2.salesforce');
  
  // If currentOAuthProps is missing or empty, initiate authentication flow
  if (!currentOAuthProps || Object.keys(JSON.parse(currentOAuthProps)).length === 0) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Authentication Required', 'You need to authenticate with Salesforce before proceeding.', ui.ButtonSet.OK);
      getSfService();
      showSidebar();
      return;
  }

  currentOAuthProps = JSON.parse(currentOAuthProps);

  const refreshTokenEndpoint = `login.salesforce.com/services/oauth2/token?grant_type=refresh_token&client_id=${clientId}&client_secret=${clientSecret}&refresh_token=${currentOAuthProps.refresh_token}`;
  const response = UrlFetchApp.fetch(refreshTokenEndpoint, { method: 'POST' });

  const newOAuthProps = JSON.parse(response);
  Object.assign(currentOAuthProps, {
    signature: newOAuthProps.signature,
    access_token: newOAuthProps.access_token,
    issued_at: newOAuthProps.issued_at
  });

  userProps.setProperty('oauth2.salesforce', JSON.stringify(currentOAuthProps));
}

function getPastDate() {
  let today = new Date();
  let year = today.getFullYear() - 2;
  let month = today.getMonth();  // Months are 0-indexed in JavaScript.

  if (month === 0) {
    year -= 1;
    month = 11;  // December of the previous year.
  } else {
    month -= 1;  // Subtract 1 month.
  }

  // Convert single-digit months to two-digit format.
  let formattedMonth = (month + 1).toString().padStart(2, '0');

  return `${year}-${formattedMonth}-01T00:00:00Z`;
}

function makeRequestSoql(parentSheetFilter = "") {
  const ss = SpreadsheetApp.getActive();
  const settingsSheet = ss.getSheetByName('Settings');
  const dataRange = settingsSheet.getDataRange().getValues();
  const sheetsUpdated = new Set(); // Track which sheets have been updated
  const numRows = settingsSheet.getLastRow();

  for (let i = 3; i < numRows; i++) {
    const rowData = dataRange[i];
    const parentSheetName = rowData[0];
    const sfObjectType = rowData[1];
    const sfFieldsList = rowData[2];
    const supportProduct = rowData[3];
    const types = rowData[4];
    const modulePortalInclude = rowData[5];
    const modulePortalExclude = rowData[6];
    const shouldUpdate = rowData[7];

    if (parentSheetFilter && parentSheetFilter !== parentSheetName) {
      continue;  // If a filter is provided and it doesn't match the current parentSheetName, skip this iteration
    }

    if (shouldUpdate !== "Yes") {
      continue;  // Skip the current iteration and move to the next row
    }

    let whereClause = `WHERE CreatedDate > ${getPastDate()}`;
    if (sfObjectType === "Case") {
      whereClause = `WHERE Status != 'Cancelled' AND CreatedDate > ${getPastDate()}`;
    }

    if (supportProduct !== undefined) {
      whereClause += ` AND Support_Product_Portal__c = '${supportProduct}'`;
    }

    if (types) {
      const fieldName = sfObjectType === "Case" ? "Type" : (sfObjectType === "Issues__c" ? "Issue_Type__c" : null);

      if (fieldName) {
        const typeList = types.split(",").map(type => type.trim());
        const typeConditions = typeList.map(type => `${fieldName} = '${type}'`).join(" OR ");
        whereClause += ` AND (${typeConditions})`;
      }
    }

    if (modulePortalInclude) {
      const includedModules = modulePortalInclude.split(",").map(value => `Module_Portal__c = '${value.trim()}'`).join(" OR ");
      whereClause += ` AND (${includedModules})`;
    }

    if (modulePortalExclude) {
      const excludedModules = modulePortalExclude.split(",").map(value => `Module_Portal__c != '${value.trim()}'`).join(" AND ");
      whereClause += ` AND (${excludedModules})`;
    }

    const soqlTemplate = `SELECT ${sfFieldsList} FROM ${sfObjectType} ${whereClause}`;
    const sfService = getSfService();
    const oauthData = JSON.parse(PropertiesService.getUserProperties().getProperty('oauth2.salesforce'));
    let queryUrl = `${oauthData.instance_url}/services/data/v51.0/query?q=${encodeURIComponent(soqlTemplate)}`;

    do {
      const response = UrlFetchApp.fetch(queryUrl, {
        method: "GET",
        headers: { "Authorization": `OAuth ${sfService.getAccessToken()}` },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) {
        Logger.log(response.getContentText());
        return;
      }

      const responseData = JSON.parse(response.getContentText());
      const resultArray = toArray(responseData.records);

      let responseSheet = ss.getSheetByName(parentSheetName);
      if (!responseSheet) {
        Logger.log(`Sheet ${parentSheetName} not found. Ensure the sheet exists.`);
        return;
      }

      if (!sheetsUpdated.has(parentSheetName)) {
        if (responseSheet.getLastRow() > 1) {
          responseSheet.getRange(2, 1, responseSheet.getLastRow() - 1, responseSheet.getLastColumn()).clearContent();
        }
        sheetsUpdated.add(parentSheetName);
      }

      settingsSheet.getRange(i + 1, 10).setValue(responseData.totalSize);

      if (resultArray.length > 1) {
        const lastRow = responseSheet.getLastRow();
        responseSheet.getRange(lastRow + 1, 1, resultArray.length - 1, resultArray[0].length).setValues(resultArray.slice(1));
        Logger.log(`${resultArray.length - 1} results added to ${parentSheetName} for (Product: ${supportProduct})`);
      } else {
        Logger.log(`No results to write for sheet ${parentSheetName} - (Product: ${supportProduct})`);
      }

      queryUrl = responseData.nextRecordsUrl ? `${oauthData.instance_url}${responseData.nextRecordsUrl}` : null;

    } while (queryUrl);
  }
}

function toArray(items) {
  // Check if items is empty or the first item is undefined/null
  if (!items.length || !items[0]) {
    return [];
  }

  const headers = Object.keys(items[0]).filter(key => key !== "attributes").concat(["FormattedDate"]);
  const data = items.map(item => {
    let values = headers.slice(0, -1).map(header => item[header]);
    let formattedDate = formatCreatedDate(item.CreatedDate); // Add formatted date at the end.
    values.push(formattedDate);
    return values;
  });

  return [headers, ...data];
}

function formatCreatedDate(dateStr) {
  let date = new Date(dateStr);
  let month = (date.getMonth() + 1).toString().padStart(2, '0'); // Months are 0-indexed.
  let year = date.getFullYear();
  return `${month}-${year}`;
}


function getFormattedCurrentDate() {
    let today = new Date();
    let day = today.getDate().toString().padStart(2, '0');
    let month = (today.getMonth() + 1).toString().padStart(2, '0'); // Months are 0-indexed.
    let year = today.getFullYear();
    return `${month}-${day}-${year}`;
}

function setReportDateRange() {
    const ss = SpreadsheetApp.getActive();
    const settingsSheet = ss.getSheetByName('Settings');
    const summarySheet = ss.getSheetByName('Summary'); // Get 'Summary' sheet
    
    let currentDate = new Date();
    let startDate = new Date(currentDate.getFullYear() - 2, currentDate.getMonth() - 1, 1); // 2 years and 1 month ago from the current date.
    let endDate = new Date(); // Current date.

    settingsSheet.getRange('B1').setValue(endDate);
    
    // Reset the hours, minutes, seconds, and milliseconds
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    // Set the two dates in the 'Settings' sheet at D1 and E1
    settingsSheet.getRange('H1').setValue(startDate);
    settingsSheet.getRange('I1').setValue(endDate);

    startDate = new Date(currentDate.getFullYear() - 1, currentDate.getMonth() - 1, 1); // 2 years and 1 month ago from the current date.
    startDate.setHours(0, 0, 0, 0);

    // Set the two dates in the 'Summary' sheet at B1 and C1
    summarySheet.getRange('B1').setValue(startDate);
    summarySheet.getRange('C1').setValue(endDate);

    // Generate and write the list of month start dates between startDate and endDate
    let monthDates = getMonthStartDates(startDate, endDate);
    summarySheet.getRange(5, 1, monthDates.length, 1).setValues(monthDates.map(date => [date]));
}

function getMonthStartDates(startDate, endDate) {
    let dates = [];
    let currentDate = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
    
    while (currentDate <= endDate) {
        let month = (currentDate.getMonth() + 1).toString().padStart(2, '0'); // Months are 0-indexed.
        let year = currentDate.getFullYear();
        dates.push(`${month}-${year}`);
        
        currentDate.setMonth(currentDate.getMonth() + 1); // Move to the next month.
    }
    
    return dates;
}

function compileSheetsToDomo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  const lastRow = settingsSheet.getLastRow();
  const settingsData = settingsSheet.getRange(4, 1, lastRow - 3, 9).getValues();

  const processedSheets = new Set();
  
  const destId = PropertiesService.getScriptProperties().getProperty("domoDestSheetID");
  const destSS = SpreadsheetApp.openById(destId);
  
  const hasCleared = {
    "SCR_01": false,
    "SCR_02": false,
    "SIR_01": false
  };

  const domoMappings = {
    "Yes, SCR_01": "SCR_01",
    "Yes, SCR_02": "SCR_02",
    "Yes, SIR_01": "SIR_01"
  };
  
  settingsData.forEach(row => {
    const sourceSheetName = row[0];
    const domoSetting = row[8];

    if (!processedSheets.has(sourceSheetName) && domoMappings[domoSetting]) {
      const targetSheetName = domoMappings[domoSetting];
      addDataToTargetSheet(ss, sourceSheetName, targetSheetName, destSS, hasCleared);
      processedSheets.add(sourceSheetName);
      Logger.log(`${sourceSheetName} has been added to ${targetSheetName} in the destination sheet`);
    }
  });
}

function addDataToTargetSheet(ss, sourceSheetName, targetSheetName, destSS, hasCleared) {
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const targetSheet = destSS.getSheetByName(targetSheetName); 

  if (!sourceSheet || !targetSheet) {
    Logger.log(`Source or Target sheet not found. Source: ${sourceSheetName}, Target: ${targetSheetName}`);
    return;
  }

  if (!hasCleared[targetSheetName]) {
    const lastRowTarget = targetSheet.getLastRow();
    if (lastRowTarget > 1) {
      targetSheet.getRange(2, 1, lastRowTarget - 1, targetSheet.getLastColumn()).clearContent();
      hasCleared[targetSheetName] = true;
    }
  }

  const lastRowSource = sourceSheet.getLastRow();

  if (lastRowSource <= 1) { // Only headers, no data
    return;
  }

  const sourceData = sourceSheet.getRange(2, 1, lastRowSource - 1, sourceSheet.getLastColumn()).getValues();

  if (sourceData.length > 0) {
    const lastRowTargetAfterClear = targetSheet.getLastRow();
    targetSheet.getRange(lastRowTargetAfterClear + 1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
  }
}